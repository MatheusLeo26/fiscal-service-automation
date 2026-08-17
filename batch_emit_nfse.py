import os
import time
from datetime import datetime
import re
import json
import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Load credentials from .env
load_dotenv()

# Garantir segurança cibernética: remover qualquer variável com prefixo NEXT_PUBLIC_
for key in list(os.environ.keys()):
    if key.upper().startswith("NEXT_PUBLIC_"):
        del os.environ[key]

# Configurações de Execução
SALVAR_EVIDENCIA_ERRO = True # Guardar fotos de erro para ajudar no suporte (Sugerido: True)

def emit_nfse_batch(headless=False):
    # 1. Setup paths and Load Excel early
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, "clientes.xlsx")
    checkpoint_path = os.path.join(current_dir, "progresso.json")
    
    if not os.path.exists(excel_path):
        print(f"[ERROR] Planilha não encontrada em: {excel_path}")
        return
        
    df = pd.read_excel(excel_path)
    # Ensure CNPJ is string for comparison
    df['CNPJ'] = df['CNPJ'].astype(str).str.strip()
    
    # Check for previous progress
    cnpj_concluidos = []
    if os.path.exists(checkpoint_path):
        print("\n" + "!"*40)
        print("      HISTÓRICO DE EMISSÕES")
        print("!"*40)
        resumo = input("Encontrei notas já emitidas anteriormente. Deseja CONTINUAR de onde parou? (S/N): ").strip().upper()
        if resumo == 'S':
            try:
                with open(checkpoint_path, "r", encoding="utf-8") as f:
                    cnpj_concluidos = json.load(f)
                print(f"[INFO] Resumindo lote. {len(cnpj_concluidos)} CNPJs serão pulados.")
            except:
                print("[WARN] Erro ao ler progresso anterior. Iniciando do zero.")
        else:
            print("[INFO] Iniciando novo lote (histórico descartado).")
            try: os.remove(checkpoint_path)
            except: pass
    
    # Filter pending clients
    df_pendentes = df[~df['CNPJ'].isin(cnpj_concluidos)].copy().reset_index(drop=True)
    
    if df_pendentes.empty:
        print("[INFO] Todas as notas da planilha já foram emitidas conforme o histórico.")
        return

    # 2. Get Alíquota input
    print("\n" + "="*40)
    print("      CONFIGURAÇÃO DA AUTOMAÇÃO")
    print("="*40)
    aliquota = input("Por favor, digite a alíquota de ISS do mês (ex: 2.24): ").replace(',', '.')
    
    # 3. Manual Overrides Menu (only for pending clients)
    skip_indices = []
    print("\nREVISÃO DE DADOS (NOTAS PENDENTES):")
    opcao_alterar = input("Deseja alterar valores/descrições ou pular algum cliente pendente? (S/N): ").strip().upper()
    
    if opcao_alterar == 'S':
        while True:
            print("\nLISTA DE CLIENTES PENDENTES:")
            for i, row in df_pendentes.iterrows():
                status = "[PULAR]" if i in skip_indices else "[OK]"
                print(f"{i+1}. {status} {row['Nome da empresa']} - R$ {row['VALOR']}")
            
            escolha = input("\nNúmero para editar, 'N' p/ remover, ou 'F' p/ FINALIZAR e começar: ").strip().upper()
            
            if escolha == 'F':
                break
            
            try:
                idx = int(escolha) - 1
                if idx < 0 or idx >= len(df_pendentes):
                    print("[!] Número inválido.")
                    continue
                
                print(f"\nEditando: {df_pendentes.iloc[idx]['Nome da empresa']}")
                acao = input("Ação? (V = Valor, D = Descrição, P = Pular, C = Cancelar): ").strip().upper()
                
                if acao == 'V':
                    df_pendentes.at[idx, 'VALOR'] = input(f"Novo valor (atual: {df_pendentes.iloc[idx]['VALOR']}): ").strip()
                elif acao == 'D':
                    df_pendentes.at[idx, 'Discriminação dos Serviços'] = input(f"Nova descrição: ").strip()
                elif acao == 'P':
                    if idx not in skip_indices: skip_indices.append(idx)
                    else: skip_indices.remove(idx)
            except ValueError:
                print("[!] Entrada inválida.")

    # Remove skipped clients from pending list
    df_execucao = df_pendentes.drop(skip_indices).reset_index(drop=True)
    
    if df_execucao.empty:
        print("[INFO] Nenhum cliente restou para processar. Encerrando.")
        return

    # 4. Setup Evidence Folder
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    evidence_dir = os.path.join(current_dir, "evidencias", f"execucao_{timestamp}")
    os.makedirs(evidence_dir, exist_ok=True)
    print(f"\n[INFO] Pasta de evidências criada: {evidence_dir}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()
        
        # 3.1 Initialize Report List
        relatorio_notas = []
        
        # 3. Login
        cpf = os.getenv("CPF")
        password = os.getenv("PASSWORD")
        
        if not cpf or not password:
            print("[ERROR] CPF ou Senha não encontrados no arquivo .env")
            return

        print("[INFO] Fazendo login...")
        page.goto("https://itu.giss.com.br/#/")
        
        # Handle Chrome profile selector if it appears (from user's screenshot)
        try:
            # The selector "Seu Chrome" or similar based on user's image
            if page.query_selector("text='Seu Chrome'") or page.query_selector("div[aria-label*='Chrome']"):
                print("[INFO] Tela de perfil do Chrome detectada. Ignorando...")
                page.click("text='Seu Chrome'")
                page.wait_for_load_state("networkidle")
        except:
            pass
            
        # Optimized route: Click directly on 'Emitir NFS-e' to go to login
        try:
            selector_emitir = "text='Emitir NFS-e'"
            page.wait_for_selector(selector_emitir, timeout=10000)
            
            # This action opens a NEW tab
            with context.expect_page() as new_page_info:
                page.click(selector_emitir)
            
            # Switch to the new tab
            page = new_page_info.value
            page.wait_for_load_state("networkidle")
            print("[INFO] Botão 'Emitir NFS-e' clicado e nova aba detectada.")
        except Exception as e:
            print(f"[WARN] Falha na rota direta, tentando alternativa: {str(e)}")
            try:
                page.click("button:has-text('Entrar'), a:has-text('Entrar')")
                page.wait_for_load_state("networkidle")
            except:
                pass

        # -- UPDATE: Handle pre-login popups on the Login Tab (e.g., "EMPRESAS QUE EMITEM NFS-e POR WEBSERVICE")
        print("[INFO] Verificando popups na tela de login...")
        try:
            # Wait up to 5 seconds for the popup 'OK' button to appear before trying to close
            try:
                page.wait_for_selector("text='OK', text='Ok'", timeout=5000)
            except:
                pass
                
            # Esc key is very effective for modal dialogs
            page.keyboard.press("Escape")
            time.sleep(0.5)
                
            # Use multiple robust locator strategies to find and click the OK/Close buttons
            locators_to_try = [
                page.get_by_role("button", name="OK"),
                page.get_by_text("OK", exact=True),
                page.get_by_text("Ok", exact=True),
                page.locator("button:has-text('OK')"),
                page.locator(".close"),
                page.locator(".btn-close"),
                page.locator("button[aria-label='Close']")
            ]
            
            for loc in locators_to_try:
                try:
                    for element in loc.all():
                        if element.is_visible():
                            print("[INFO] Botão de popup detectado! Emulando clique...")
                            element.click(force=True)
                            time.sleep(1)
                except:
                    pass
        except Exception as e:
            print(f"[WARN] Erro silencioso ao lidar com popups: {e}")

        # Wait for login fields (CPF/Usuario and Password)
        try:
            # Check for IDs commonly used in this portal
            page.wait_for_selector("#usuario, input[name='cpf'], input[placeholder*='Usuário']", timeout=20000)
            
            # Fill CPF/Usuario
            if page.query_selector("#usuario"):
                page.fill("#usuario", cpf)
            elif page.query_selector("input[name='cpf']"):
                page.fill("input[name='cpf']", cpf)
            else:
                page.fill("input[placeholder*='Usuário']", cpf)
            
            # Click Advance or Acessar depending on the screen
            if page.query_selector("button:has-text('Avançar')"):
                page.click("button:has-text('Avançar')")
                time.sleep(1)

            page.wait_for_selector("#senha, input[name='password'], input[type='password']", timeout=10000)
            
            if page.query_selector("#senha"):
                page.fill("#senha", password)
            elif page.query_selector("input[name='password']"):
                page.fill("input[name='password']", password)
            else:
                page.fill("input[type='password']", password)
                
            page.click("button:has-text('Acessar'), button.btn-primary")
        except Exception as e:
            print(f"[ERROR] Falha ao localizar campos de login na aba ativa: {str(e)}")
            page.screenshot(path=os.path.join(evidence_dir, "erro_login_aba.png"))
            return
        
        # Handle "Ir para o sistema" landing page
        try:
            page.wait_for_selector("button:has-text('Ir para o sistema')", timeout=10000)
            page.click("button:has-text('Ir para o sistema')")
            print("[INFO] Botão 'Ir para o sistema' clicado.")
            page.wait_for_load_state("networkidle")
        except:
            pass
            
        # 4. Handle Popups
        print("[INFO] Removendo avisos iniciais...")
        time.sleep(2) # Give some time for SPA to settle
        popups = page.query_selector_all("button[aria-label='Close'], .modal-header .close, button:has-text('OK'), .btn-close")
        for popup in popups:
            try:
                if popup.is_visible():
                    popup.click()
                    time.sleep(1)
            except:
                pass

        # 5. Select Company SGR
        print("[INFO] Selecionando empresa SGR CONTABIL...")
        try:
            # Check for 404 or profile screen
            if "Página não encontrada" in page.content():
                print("[WARN] 404 detectado. Tentando recarregar...")
                page.goto("https://itu.giss.com.br/portal/home#/login-portal?pageRedirect=emitir-nfs")
                page.wait_for_load_state("networkidle")

            selector_busca = "input[placeholder*='Pesquisar'], #pesquisar, input[type='search']"
            page.wait_for_selector(selector_busca, timeout=20000)
            target_busca = page.locator(selector_busca).first
            target_busca.fill("SGR CONTABIL E ASSESSORIA LTDA")
            page.keyboard.press("Enter")
            time.sleep(2) # Wait for table update
            
            # Selector for the specific button based on DOM inspection
            selector_seta = "button[title='Selecionar Empresa'], i.fa-arrow-right, .btn-entrar-empresa"
            page.wait_for_selector(selector_seta, timeout=10000)
            print("[INFO] Clicando na 'setinha azul' de seleção (botão)...")
            
            # Use locator for a more precise click
            page.locator(selector_seta).first.click()
            
            # CRITICAL: Wait for the specific target URL or page content
            print("[INFO] Aguardando transição para o painel de emissão...")
            try:
                page.wait_for_url("**/operacao/servicos-prestados/emitir-nfse", timeout=20000)
                page.wait_for_load_state("networkidle")
            except:
                print("[WARN] URL esperada não carregou no tempo previsto. Verificando conteúdo...")
                if "Página não encontrada" in page.content():
                    print("[ERROR] 404 após clique na seta. Forçando navegação...")
                    page.goto("https://itu.giss.com.br/portal/home#/operacao/servicos-prestados/emitir-nfse")
                    page.wait_for_load_state("networkidle")

        except Exception as e_sgr:
            print(f"[WARN] Falha ao selecionar SGR CONTABIL: {str(e_sgr)}")
            page.screenshot(path=os.path.join(evidence_dir, "erro_sgr_contabil.png"))

        # 6. Batch Loop
        for index, row in df_execucao.iterrows():
            # Dynamic mapping to handle encoding issues in column names
            col_empresa = [c for c in df_execucao.columns if 'empresa' in c or 'RAZAO' in c][0]
            col_descricao = [c for c in df_execucao.columns if 'Discrimina' in c][0]
            
            cnpj = str(row['CNPJ']).strip()
            raw_valor = row['VALOR']
            if isinstance(raw_valor, (int, float)):
                valor = f"{float(raw_valor):.2f}".replace('.', ',')
            else:
                raw_str = str(raw_valor).replace('R$', '').strip()
                if ',' in raw_str and '.' in raw_str:
                    raw_str = raw_str.replace('.', '')
                elif '.' in raw_str and len(raw_str.split('.')[-1]) <= 2:
                    raw_str = raw_str.replace('.', ',')
                if ',' not in raw_str:
                    raw_str += ",00"
                valor = raw_str
            
            nome_empresa = str(row[col_empresa]).strip()
            
            # OVERRIDE: Tratamento específico para a Primeira Igreja Batista
            # Na planilha está 'EM ITU' e no sistema está 'DE ITU'. Isso faz a Estratégia A falhar e cair no CNPJ.
            if cnpj == "48994164000108":
                nome_empresa = "PRIMEIRA IGREJA BATISTA"
                
            descricao = str(row[col_descricao]).strip()
            
            print(f"\n[PROCESS] Emitindo nota ({index+1}/{len(df_execucao)}): {nome_empresa} ({cnpj})")
            print(f"[INFO] Dados da Planilha: Valor=R$ {valor}, Descrição='{descricao[:30]}...'")
            
            try:
                # Check if we are already on the emission page (via URL or header)
                if "/emitir-nfse" not in page.url:
                    print("[INFO] Navegando para o menu de emissão (Início do loop)...")
                    # Use robust selector to avoid strict mode violations
                    target_sidebar = page.get_by_text("Serviços Prestados", exact=False).first
                    target_sidebar.scroll_into_view_if_needed()
                    target_sidebar.click()
                    
                    emitir_card = page.get_by_text("Emitir NFS-e", exact=True).first
                    emitir_card.scroll_into_view_if_needed()
                    emitir_card.click()
                    page.wait_for_load_state("networkidle")
                else:
                    print("[INFO] Já está na página de emissão. Prosseguindo...")
                
                # EMISSION SEQUENCE (12 PASSOS)
                print(f"[INFO] Processando Tomador: {nome_empresa} ({cnpj})")

                # Passo 1: Serviço / Atividade
                print("[INFO] Passo 1: Selecionando Atividade...")
                try:
                    # Specific ID found by subagent
                    selector_atividade = "select#atividadeServico, select[name='atividadeServico'], select[name='atividade']"
                    page.wait_for_selector(f"{selector_atividade} option", timeout=20000)
                    
                    # Busca dinâmica pela atividade correta usando JavaScript puro para contornar problemas de visibilidade do HTMLOptionElement
                    target_value = page.evaluate('''() => {
                        const select = document.querySelector("select#atividadeServico") || 
                                       document.querySelector("select[name='atividadeServico']") || 
                                       document.querySelector("select[name='atividade']");
                        if (!select) return null;
                        for (let i = 0; i < select.options.length; i++) {
                            const text = select.options[i].text;
                            if (text.includes("17.19") || text.includes("CONTABILIDADE")) {
                                return select.options[i].value;
                            }
                        }
                        return null;
                    }''')
                    
                    if target_value:
                        page.select_option(selector_atividade, value=target_value)
                        print(f"[INFO] Atividade selecionada por valor ({target_value}) via JS.")
                    else:
                        print("[WARN] JS não achou '17.19'. Tentando Fallback para índice fixo...")
                        page.select_option(selector_atividade, index=3)
                        
                except Exception as e_step1:
                    print(f"[ERROR] Falha na seleção da Atividade: {str(e_step1)}. Isso pode quebrar os próximos passos.")

                # Passo 2: NBS (Só após passo 1)
                print("[INFO] Passo 2: Selecionando NBS (Apenas clique)...")
                time.sleep(3) # Wait for site to react to Step 1
                
                # Click field to open the Angular dropdown - with retry logic
                selector_nbs_input = "#ibs_nbs_value, input[name='nbsAuto'], input[placeholder*='NBS'], input[id*='nbs'], input[ng-model*='nbs']"
                
                nbs_found = False
                for tentativa_nbs in range(3):
                    try:
                        page.wait_for_selector(selector_nbs_input, timeout=20000)
                        nbs_found = True
                        break
                    except:
                        print(f"[WARN] NBS não apareceu (tentativa {tentativa_nbs+1}/3). Aguardando mais...")
                        time.sleep(3)
                        # Tenta fechar popups que podem estar bloqueando
                        try:
                            page.keyboard.press("Escape")
                            time.sleep(1)
                        except:
                            pass
                        # Se for a última tentativa, tenta recarregar apenas a seção
                        if tentativa_nbs == 1:
                            try:
                                # Re-seleciona a atividade para forçar o Angular a renderizar o campo NBS
                                print("[INFO] Re-selecionando Atividade para forçar carregamento do NBS...")
                                selector_atividade = "select#atividadeServico, select[name='atividadeServico'], select[name='atividade']"
                                page.select_option(selector_atividade, index=0)
                                time.sleep(1)
                                target_value = page.evaluate('''() => {
                                    const select = document.querySelector("select#atividadeServico") || 
                                                   document.querySelector("select[name='atividadeServico']") || 
                                                   document.querySelector("select[name='atividade']");
                                    if (!select) return null;
                                    for (let i = 0; i < select.options.length; i++) {
                                        const text = select.options[i].text;
                                        if (text.includes("17.19") || text.includes("CONTABILIDADE")) {
                                            return select.options[i].value;
                                        }
                                    }
                                    return null;
                                }''')
                                if target_value:
                                    page.select_option(selector_atividade, value=target_value)
                                time.sleep(3)
                            except:
                                pass
                
                if not nbs_found:
                    raise Exception("Campo NBS não apareceu após 3 tentativas. Pulando cliente.")
                
                page.locator(selector_nbs_input).first.scroll_into_view_if_needed()
                
                # Limpa o campo de forma mais robusta para SPAs
                page.fill(selector_nbs_input, "")
                page.locator(selector_nbs_input).first.blur()
                time.sleep(1)
                
                # Digita apenas parte do código LENTAMENTE para forçar o Angular a buscar
                page.locator(selector_nbs_input).first.click()
                page.locator(selector_nbs_input).first.press_sequentially("1.1302.21", delay=200)
                
                # Wait for the specific dropdown row to appear
                print("[INFO] Aguardando lista NBS aparecer...")
                time.sleep(2) # Dar tempo pro dropdown Angular renderizar
                
                # Tenta múltiplos seletores para a opção do dropdown NBS
                nbs_clicked = False
                
                # Estratégia 1: Seletores genéricos de dropdown/autocomplete
                dropdown_selectors = [
                    ".angucomplete-row:has-text('1.1302.21')",
                    ".angucomplete-row:has-text('contabilidade')",
                    "div[class*='angucomplete'] div:has-text('1.1302.21')",
                    "div[class*='autocomplete'] div:has-text('1.1302.21')",
                    "ul[class*='dropdown'] li:has-text('1.1302.21')",
                    "div[class*='suggestion']:has-text('1.1302.21')",
                    "div[class*='dropdown']:has-text('contabilidade')",
                ]
                
                for sel in dropdown_selectors:
                    try:
                        loc = page.locator(sel).first
                        if loc.is_visible(timeout=1000):
                            loc.click(timeout=3000)
                            nbs_clicked = True
                            print(f"[INFO] NBS selecionado via seletor: {sel}")
                            break
                    except:
                        continue
                
                # Estratégia 2: Busca por texto visível genérico na página
                if not nbs_clicked:
                    try:
                        # Procura qualquer elemento visível que contenha o texto da opção
                        nbs_text_loc = page.get_by_text("1.1302.21.00", exact=False).first
                        # Garante que não é o próprio input (queremos o item da lista)
                        tag = nbs_text_loc.evaluate("el => el.tagName")
                        if tag.upper() != "INPUT":
                            nbs_text_loc.click(timeout=3000)
                            nbs_clicked = True
                            print("[INFO] NBS selecionado via texto visível.")
                    except:
                        pass
                
                # Estratégia 3: Confirma com Enter (o autocomplete já preencheu o campo)
                if not nbs_clicked:
                    print("[WARN] Nenhum seletor de dropdown funcionou. Confirmando via Enter...")
                    page.locator(selector_nbs_input).first.press("Enter")
                    time.sleep(1)
                    # Verifica se o campo ainda tem o valor (se o Enter limpou, tenta Tab)
                    valor_campo = page.locator(selector_nbs_input).first.input_value()
                    if "1.1302" in valor_campo:
                        print("[INFO] NBS confirmado via tecla Enter.")
                        nbs_clicked = True
                    else:
                        page.locator(selector_nbs_input).first.press("Tab")
                        print("[INFO] NBS confirmado via tecla Tab.")
                        nbs_clicked = True
                
                time.sleep(1)

                # Passo 2b: Código Indicador da Operação (novo campo obrigatório do site)
                print("[INFO] Passo 2b: Selecionando Código Indicador da Operação (030101)...")
                time.sleep(1)
                try:
                    # Localiza o campo de autocomplete do Código Indicador
                    selector_cod_ind = "input[placeholder*='Selecione'], input[name*='indicador'], input[name*='codigoIndicador'], input[ng-model*='indicador']"
                    # Tenta localizar o campo pelo label próximo
                    cod_ind_field = None
                    
                    # Estratégia 1: Pelo label "Código Indicador da Operação"
                    try:
                        label_cod = page.get_by_text("Código Indicador da Operação", exact=False).first
                        # Pega o input mais próximo (irmão ou filho do container pai)
                        parent = label_cod.locator("xpath=..")
                        cod_ind_field = parent.locator("input").first
                        if not cod_ind_field.is_visible(timeout=2000):
                            cod_ind_field = None
                    except:
                        pass
                    
                    # Estratégia 2: Todos os campos "Selecione..." na página e pegar o segundo (primeiro é CST)
                    if not cod_ind_field:
                        selecione_inputs = page.locator("input[placeholder='Selecione...']").all()
                        if len(selecione_inputs) >= 1:
                            cod_ind_field = selecione_inputs[0]  # Primeiro "Selecione..." após NBS
                    
                    if cod_ind_field:
                        cod_ind_field.scroll_into_view_if_needed()
                        cod_ind_field.click()
                        cod_ind_field.fill("")
                        cod_ind_field.press_sequentially("030101", delay=150)
                        time.sleep(2)
                        
                        # Clicar na opção do dropdown
                        cod_clicked = False
                        dropdown_sels = [
                            ".angucomplete-row:has-text('030101')",
                            "div[class*='angucomplete'] div:has-text('030101')",
                            "div[class*='autocomplete'] div:has-text('030101')",
                            "div[class*='dropdown']:has-text('030101')",
                            "ul li:has-text('030101')",
                        ]
                        for sel in dropdown_sels:
                            try:
                                loc = page.locator(sel).first
                                if loc.is_visible(timeout=1000):
                                    loc.click(timeout=3000)
                                    cod_clicked = True
                                    print(f"[INFO] Código Indicador selecionado via: {sel}")
                                    break
                            except:
                                continue
                        
                        if not cod_clicked:
                            # Fallback: texto visível
                            try:
                                opt = page.get_by_text("030101", exact=False).first
                                tag = opt.evaluate("el => el.tagName")
                                if tag.upper() != "INPUT":
                                    opt.click(timeout=3000)
                                    cod_clicked = True
                                    print("[INFO] Código Indicador selecionado via texto visível.")
                            except:
                                pass
                        
                        if not cod_clicked:
                            # Último recurso: Enter/Tab
                            cod_ind_field.press("Enter")
                            time.sleep(0.5)
                            print("[WARN] Código Indicador confirmado via Enter.")
                    else:
                        print("[WARN] Campo 'Código Indicador da Operação' não encontrado. Pulando...")
                except Exception as e_cod:
                    print(f"[WARN] Erro ao preencher Código Indicador: {e_cod}")
                
                time.sleep(1)

                # Passo 2c: Classificação Tributária (novo campo obrigatório do site)
                print("[INFO] Passo 2c: Selecionando Classificação Tributária (200052)...")
                try:
                    # Localiza o campo de autocomplete da Classificação Tributária
                    class_trib_field = None
                    
                    # Estratégia 1: Pelo label
                    try:
                        label_class = page.get_by_text("Classificação Tributária", exact=False).first
                        parent = label_class.locator("xpath=..")
                        class_trib_field = parent.locator("input").first
                        if not class_trib_field.is_visible(timeout=2000):
                            class_trib_field = None
                    except:
                        pass
                    
                    # Estratégia 2: Campos "Selecione..." restantes
                    if not class_trib_field:
                        selecione_inputs = page.locator("input[placeholder='Selecione...']").all()
                        # O primeiro agora é Classificação Tributária (Código Indicador já foi preenchido)
                        for inp in selecione_inputs:
                            try:
                                val = inp.input_value()
                                if not val or val.strip() == "":
                                    class_trib_field = inp
                                    break
                            except:
                                continue
                    
                    if class_trib_field:
                        class_trib_field.scroll_into_view_if_needed()
                        class_trib_field.click()
                        class_trib_field.fill("")
                        class_trib_field.press_sequentially("200052", delay=150)
                        time.sleep(2)
                        
                        # Clicar na opção do dropdown
                        class_clicked = False
                        dropdown_sels = [
                            ".angucomplete-row:has-text('200052')",
                            "div[class*='angucomplete'] div:has-text('200052')",
                            "div[class*='autocomplete'] div:has-text('200052')",
                            "div[class*='dropdown']:has-text('200052')",
                            "ul li:has-text('200052')",
                        ]
                        for sel in dropdown_sels:
                            try:
                                loc = page.locator(sel).first
                                if loc.is_visible(timeout=1000):
                                    loc.click(timeout=3000)
                                    class_clicked = True
                                    print(f"[INFO] Classificação Tributária selecionada via: {sel}")
                                    break
                            except:
                                continue
                        
                        if not class_clicked:
                            try:
                                opt = page.get_by_text("200052", exact=False).first
                                tag = opt.evaluate("el => el.tagName")
                                if tag.upper() != "INPUT":
                                    opt.click(timeout=3000)
                                    class_clicked = True
                                    print("[INFO] Classificação Tributária selecionada via texto visível.")
                            except:
                                pass
                        
                        if not class_clicked:
                            class_trib_field.press("Enter")
                            time.sleep(0.5)
                            print("[WARN] Classificação Tributária confirmada via Enter.")
                    else:
                        print("[WARN] Campo 'Classificação Tributária' não encontrado. Pulando...")
                except Exception as e_class:
                    print(f"[WARN] Erro ao preencher Classificação Tributária: {e_class}")
                
                # Aguarda o CST-IBS/CBS ser preenchido automaticamente
                print("[INFO] Aguardando CST-IBS/CBS ser preenchido automaticamente...")
                time.sleep(2)

                # Passo 3: Dados do Tomador de Serviço
                print(f"[INFO] Passo 3: Pesquisando Tomador: {cnpj}...")
                selector_busca_tomador = "input#buscarTomador, input[name='cnpjTomador'], input[placeholder*='Pesquisar'], input[placeholder*='Tomador']"
                page.wait_for_selector(selector_busca_tomador, timeout=10000)
                
                # PASSO 1: CLICAR NO CAMPO DE DIGITAÇÃO E DIGITAR O CNPJ DO CLIENTE
                page.fill(selector_busca_tomador, "")
                time.sleep(0.5)
                page.locator(selector_busca_tomador).click()
                page.locator(selector_busca_tomador).press_sequentially(cnpj, delay=100)
                # Removemos o "Enter" pois ele estava ativando a validação de todo o formulário (ex: acusando Valor vazio)
                page.locator(selector_busca_tomador).blur()
                time.sleep(0.5)
                
                # PASSO 2: CLICAR NO BOTÃO PESQUISAR LOGO AO LADO
                print("[INFO] Clicando no botão Pesquisar e aguardando resultado...")
                btn_pesquisar = page.locator("button:has-text('Pesquisar')").first
                
                # Removemos o force=True para garantir que ele só clique quando o botão for habilitado pelo Angular
                btn_pesquisar.wait_for(state="visible", timeout=5000)
                try:
                    btn_pesquisar.click(timeout=5000)
                except:
                    # Se não habilitou, tenta dar um 'Enter' no campo como fallback
                    page.locator(selector_busca_tomador).press("Enter")
                    
                time.sleep(3) # Aguarda retorno da pesquisa (cascata Angular)
                
                # PASSO 3: BUSCAR A OPÇÃO NA LISTA SUSPENSA E CLICAR
                print("[INFO] Buscando a opção na lista suspensa...")
                clicado = False
                
                # 1. Tenta localizar o container do dropdown de resultados
                container = None
                for sel in [".popover", ".dropdown-menu", ".ui-autocomplete", "div[class*='autocomplete']", "div[class*='angucomplete']", "div.cadastro"]:
                    try:
                        loc = page.locator(sel)
                        if loc.is_visible():
                            container = loc
                            break
                    except:
                        pass
                
                if not container:
                    try:
                        candidate = page.locator("div").filter(has_text="Cadastro de Empresas").last
                        if candidate.is_visible():
                            container = candidate
                    except:
                        pass
                        
                if container:
                    print("[INFO] Container de dropdown/cadastro de empresas detectado.")
                    # Estratégia A: Clicar no texto do nome da empresa dentro do container
                    try:
                        opcoes_nome = container.get_by_text(nome_empresa, exact=False)
                        if opcoes_nome.count() > 0:
                            opcoes_nome.first.click()
                            print("[INFO] Tomador clicado com sucesso dentro do container pelo nome.")
                            clicado = True
                    except:
                        pass
                    
                    # Estratégia B: Clicar pelo CNPJ formatado dentro do container
                    if not clicado:
                        try:
                            cnpj_pad = cnpj.strip().zfill(14)
                            cnpj_formatado = f"{cnpj_pad[:2]}.{cnpj_pad[2:5]}.{cnpj_pad[5:8]}/{cnpj_pad[8:12]}-{cnpj_pad[12:]}"
                            opcoes_cnpj = container.get_by_text(cnpj_formatado, exact=False).all()
                            for opt in opcoes_cnpj:
                                if opt.is_visible():
                                    txt = opt.inner_text().strip()
                                    if ("CONVENCAO" in txt.upper() or "CONVENÇÃO" in txt.upper()) and "IGREJA" in nome_empresa.upper():
                                        continue
                                    opt.click()
                                    print("[INFO] Tomador clicado com sucesso dentro do container pelo CNPJ.")
                                    clicado = True
                                    break
                        except:
                            pass
                            
                    # Estratégia C: Clicar no primeiro item que tenha alguma palavra em comum (evita clicar errado em matriz/filial)
                    if not clicado:
                        try:
                            # Tenta achar alguma opção que tenha pelo menos a primeira palavra forte do nome
                            palavra_chave = nome_empresa.split()[0].upper()
                            if len(palavra_chave) < 3 and len(nome_empresa.split()) > 1:
                                palavra_chave = nome_empresa.split()[1].upper()
                                
                            paragraphs = container.locator("p, div, span").all()
                            for p in paragraphs:
                                if p.is_visible():
                                    txt = p.inner_text().strip()
                                    if txt and "Cadastro" not in txt and "Cadastrar" not in txt and len(txt) > 5:
                                        # Regra específica para o problema da Igreja x Convenção
                                        if "CONVENCAO" in txt.upper() or "CONVENÇÃO" in txt.upper():
                                            if "IGREJA" in nome_empresa.upper():
                                                continue # Pula a convenção e tenta o próximo
                                                
                                        if palavra_chave in txt.upper():
                                            p.click()
                                            print(f"[INFO] Tomador clicado por fallback de palavra-chave no container: '{txt}'")
                                            clicado = True
                                            break
                                            
                            # Se ainda não clicou, usa o fallback cego original (apenas como último recurso real do container)
                            if not clicado:
                                for p in paragraphs:
                                    if p.is_visible():
                                        txt = p.inner_text().strip()
                                        if txt and "Cadastro" not in txt and "Cadastrar" not in txt and len(txt) > 5:
                                            # Evita convenção novamente
                                            if ("CONVENCAO" in txt.upper() or "CONVENÇÃO" in txt.upper()) and "IGREJA" in nome_empresa.upper():
                                                continue
                                            p.click()
                                            print(f"[INFO] Tomador clicado por fallback cego no container: '{txt}'")
                                            clicado = True
                                            break
                        except:
                            pass

                # 2. Busca Global na página se o container não foi detectado ou falhou
                if not clicado:
                    print("[WARN] Container não detectado. Iniciando busca global na página...")
                    try:
                        opcoes_globais = page.get_by_text(nome_empresa, exact=False).all()
                        for opt in opcoes_globais:
                            try:
                                if opt.is_visible():
                                    tag = opt.evaluate("el => el.tagName")
                                    # Evita clicar em inputs da própria página ou botões indesejados
                                    if tag.upper() not in ["INPUT", "TEXTAREA", "BUTTON"]:
                                        opt.click()
                                        print(f"[INFO] Tomador clicado na busca global (tag: {tag}).")
                                        clicado = True
                                        break
                            except:
                                pass
                    except Exception as e_glob:
                        print(f"[WARN] Falha na busca global por nome: {e_glob}")

                # 3. Busca Global por CNPJ formatado
                if not clicado:
                    try:
                        cnpj_pad = cnpj.strip().zfill(14)
                        cnpj_formatado = f"{cnpj_pad[:2]}.{cnpj_pad[2:5]}.{cnpj_pad[5:8]}/{cnpj_pad[8:12]}-{cnpj_pad[12:]}"
                        opcoes_globais_cnpj = page.get_by_text(cnpj_formatado, exact=False).all()
                        for opt in opcoes_globais_cnpj:
                            try:
                                if opt.is_visible():
                                    tag = opt.evaluate("el => el.tagName")
                                    if tag.upper() not in ["INPUT", "TEXTAREA", "BUTTON"]:
                                        txt = opt.inner_text().strip()
                                        if ("CONVENCAO" in txt.upper() or "CONVENÇÃO" in txt.upper()) and "IGREJA" in nome_empresa.upper():
                                            continue
                                        opt.click()
                                        print(f"[INFO] Tomador clicado na busca global por CNPJ (tag: {tag}).")
                                        clicado = True
                                        break
                            except:
                                pass
                    except Exception as e_glob_cnpj:
                        print(f"[WARN] Falha na busca global por CNPJ: {e_glob_cnpj}")

                # 4. Fallback original (Word fuzzy)
                if not clicado:
                    print("[WARN] Tentando estratégias de fallback legadas...")
                    try:
                        # Pega as primeiras 3 palavras significativas
                        keywords = [p for p in nome_empresa.split() if len(p) >= 2][:3]
                        fuzzy_name = " ".join(keywords)
                        if fuzzy_name:
                            opcoes_fuzzy = page.get_by_text(fuzzy_name, exact=False).all()
                            for opt in opcoes_fuzzy:
                                if opt.is_visible() and opt.evaluate("el => el.tagName").upper() not in ["INPUT", "TEXTAREA", "BUTTON"]:
                                    txt = opt.inner_text().strip()
                                    if ("CONVENCAO" in txt.upper() or "CONVENÇÃO" in txt.upper()) and "IGREJA" in nome_empresa.upper():
                                        continue
                                    opt.click()
                                    print(f"[INFO] Tomador selecionado via palavras-chave: '{fuzzy_name}'")
                                    clicado = True
                                    break
                    except:
                        pass

                if not clicado:
                    # Se nada funcionou, lança exceção para acionar a captura de tela e fluxo de erro
                    raise Exception("Falha definitiva ao selecionar Tomador.")
                
                time.sleep(2) # Aguarda o painel inferior de valores carregar

                # Passo 4: Valor do Serviço
                print(f"[INFO] Passo 4: Preenchendo Valor (R$ {valor})...")
                selector_valor = "input#valorServico, input[name='valorServico']"
                page.wait_for_selector(selector_valor, timeout=10000)
                # Formato final com vírgula para respeitar a máscara e uso sequencial
                page.fill(selector_valor, "")
                time.sleep(0.5)
                page.locator(selector_valor).click()
                page.locator(selector_valor).press_sequentially(valor, delay=100)
                page.locator(selector_valor).blur()

                # Passo 5: Discriminação do Serviço
                print("[INFO] Passo 5: Preenchendo Descrição...")
                selector_desc = "textarea#discriminacaoServico, textarea[name='discriminacao']"
                page.wait_for_selector(selector_desc, timeout=10000)
                page.locator(selector_desc).scroll_into_view_if_needed()
                page.fill(selector_desc, descricao)

                # Passo 6: Situação Tributária do PIS/COFINS
                print("[INFO] Passo 6: Selecionando PIS/COFINS (Nenhum)...")
                selector_pis = "select#cstPisCofins, select[name='pisCofins']"
                page.wait_for_selector(selector_pis, timeout=10000)
                page.locator(selector_pis).scroll_into_view_if_needed()
                
                # Use JS to ensure selection in Angular
                page.evaluate("""(sel) => {
                    const el = document.querySelector(sel);
                    if (el) {
                        el.value = '00';
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }""", selector_pis)
                
                # Fallback standard selection
                try:
                    page.select_option(selector_pis, value="00")
                except:
                    pass

                # Passo 7: Clicar em "Próximo"
                print("[INFO] Passo 7: Clicando em Próximo...")
                page.locator("button:has-text('Próximo')").scroll_into_view_if_needed()
                page.click("button:has-text('Próximo')")

                # Passo 8: Alíquota (pode ter mudado de lugar com a atualização do site)
                print(f"[INFO] Passo 8: Preenchendo Alíquota ({aliquota})...")
                selector_aliq = "input#aliquotaValor, input[name='aliquota'], input[name='aliquotaIss'], input[placeholder*='líquota'], input[id*='aliquota']"
                
                try:
                    page.wait_for_selector(selector_aliq, timeout=15000)
                    page.locator(selector_aliq).first.scroll_into_view_if_needed()
                    
                    # O site brasileiro normalmente usa vírgula para decimais na máscara
                    aliquota_limpa = aliquota.replace(".", ",")
                    
                    # Critical: Clear field first (Ctrl+A + Backspace) as per site mask
                    page.locator(selector_aliq).first.click()
                    page.keyboard.press("Control+A")
                    page.keyboard.press("Backspace")
                    time.sleep(0.5)
                    
                    # Fill clean value (usando type com delay simulando usuário para não bugar a máscara)
                    page.locator(selector_aliq).first.press_sequentially(aliquota_limpa, delay=150)
                    time.sleep(1)
                    print(f"[INFO] Alíquota '{aliquota_limpa}' inserida.")
                except Exception as e_aliq:
                    print(f"[WARN] Campo de Alíquota não encontrado nesta etapa: {e_aliq}")
                    print("[INFO] O site pode ter removido este campo. Continuando...")

                # Passo 9: Clicar em "Próximo"
                print("[INFO] Passo 9: Clicando em Próximo...")
                page.locator("button:has-text('Próximo')").scroll_into_view_if_needed()
                page.click("button:has-text('Próximo')")

                # Passo 10: Clicar em "Concluir"
                print("[INFO] Passo 10: Clicando em Concluir...")
                page.wait_for_selector("button:has-text('Concluir')", timeout=15000)
                page.locator("button:has-text('Concluir')").scroll_into_view_if_needed()
                page.click("button:has-text('Concluir')")

                # Passo 11: Clicar em "Não" (visualização)
                print("[INFO] Passo 11: Finalizando confirmação...")
                page.wait_for_load_state("networkidle")
                
                # Captura de número da nota antes de fechar o modal
                try:
                    texto_completo = page.inner_text("body")
                    match = re.search(r"n[uú]mero:\s*(\d+)", texto_completo, re.IGNORECASE)
                    numero_nota = match.group(1) if match else "Desconhecido"
                except:
                    numero_nota = "Confirmado"

                # Evidência de sucesso
                if not page.is_closed():
                    path_sucesso = os.path.join(evidence_dir, f"sucesso_Nota_{numero_nota}_{cnpj}.png")
                    page.screenshot(path=path_sucesso)

                try:
                    page.locator("button:has-text('Não'), button.btn-default:has-text('Não')").scroll_into_view_if_needed()
                    page.click("button:has-text('Não'), button.btn-default:has-text('Não')")
                    time.sleep(2)
                except:
                    print("[WARN] Não foi necessário clicar em 'Não' ou timeout.")

                # Passo 12: Reinício do Loop via Menu Lateral
                print(f"[SUCCESS] Nota para {nome_empresa} emitida com sucesso!")
                print("[INFO] Passo 12: Retornando ao menu para próxima emissão...")
                
                # Update progress
                cnpj_concluidos.append(cnpj)
                with open(checkpoint_path, "w", encoding="utf-8") as f:
                    json.dump(cnpj_concluidos, f)
                
                relatorio_notas.append({
                    "empresa": nome_empresa,
                    "valor": valor,
                    "numero": numero_nota
                })
                
                # NAVIGATION SEQUENCE FROM IMAGES
                try:
                    # 1. Click "Serviços Prestados" in left sidebar
                    # Based on image: Orange text, likely inside a sidebar container
                    sidebar_selector = "xpath=//div[contains(@class, 'sidebar')]//span[contains(text(), 'Serviços Prestados')] | xpath=//*[contains(@class, 'menu-item') and contains(., 'Serviços Prestados')]"
                    print("[INFO] Clicando em 'Serviços Prestados' no menu lateral...")
                    
                    # Try a simpler text-based locator first if xpath is too complex
                    target_sidebar = page.get_by_text("Serviços Prestados", exact=False).first
                    target_sidebar.scroll_into_view_if_needed()
                    target_sidebar.click()
                    time.sleep(1)
                    
                    # 2. Click "Emitir NFS-e" card/link
                    # Based on image: A card with text "Emitir NFS-e"
                    print("[INFO] Clicando em 'Emitir NFS-e' para iniciar novo formulário...")
                    emitir_card = page.get_by_text("Emitir NFS-e", exact=True).first
                    emitir_card.scroll_into_view_if_needed()
                    emitir_card.click()
                    
                    page.wait_for_load_state("networkidle")
                    print("[INFO] Painel de emissão carregado. Próximo cliente em breve...")
                    time.sleep(2)
                except Exception as e_nav:
                    print(f"[WARN] Falha na navegação pelo menu: {str(e_nav)}. Tentando recarregar rota...")
                    page.goto("https://itu.giss.com.br/portal/home#/operacao/servicos-prestados/emitir-nfse")
                    page.wait_for_load_state("networkidle")

            except Exception as e_inner:
                print(f"[ERROR] Erro no cliente {nome_empresa}: {str(e_inner)}")
                if SALVAR_EVIDENCIA_ERRO:
                    try:
                        if not page.is_closed():
                            page.screenshot(path=os.path.join(evidence_dir, f"erro_{cnpj}.png"))
                    except:
                        pass
                
                print("[INFO] Robô 'respirando'... Aguardando 10 segundos antes de tentar recuperar e ir para o próximo.")
                time.sleep(10)
                
                # Recovery - Attempt to return to the safe panel and continue to next client
                try:
                    if not page.is_closed():
                        print("[INFO] Tentando recarregar o painel principal...")
                        page.goto("https://itu.giss.com.br/portal/home#/operacao/servicos-prestados/emitir-nfse")
                        page.wait_for_load_state("networkidle")
                        time.sleep(3)
                        print(f"[WARN] Navegação de recuperação concluída. Pulando {nome_empresa}.")
                except Exception as eval_err:
                    print(f"[CRITICAL] O navegador travou ou a sessão caiu feio: {eval_err}. Interrompendo o lote.")
                    break 
                
                # Important: Continue to next row instead of breaking the loop
                continue

        # 7. Final Report
        print("\n" + "="*60)
        print("              RELATÓRIO FINAL DE EMISSÕES")
        print("="*60)
        if relatorio_notas:
            print(f"{'EMPRESA':<40} | {'VALOR':<10} | {'NÚMERO':<10}")
            print("-" * 60)
            for nota in relatorio_notas:
                print(f"{nota['empresa'][:40]:<40} | {nota['valor']:<10} | {nota['numero']:<10}")
            
            # Save to file
            relatorio_file = os.path.join(evidence_dir, "relatorio_final.txt")
            with open(relatorio_file, "w", encoding="utf-8") as f:
                f.write("RELATÓRIO FINAL DE EMISSÕES\n")
                f.write("="*60 + "\n")
                f.write(f"{'EMPRESA':<40} | {'VALOR':<10} | {'NÚMERO':<10}\n")
                f.write("-" * 60 + "\n")
                for nota in relatorio_notas:
                    f.write(f"{nota['empresa'][:40]:<40} | {nota['valor']:<10} | {nota['numero']:<10}\n")
            print(f"\n[INFO] Relatório salvo em: {relatorio_file}")
            
            # 8. Cleanup Progress if ALL pending clients were successful
            try:
                if os.path.exists(checkpoint_path) and len(relatorio_notas) == len(df_execucao):
                    print("[INFO] Todos os clientes pendentes foram processados com sucesso. Limpando memória (progresso.json).")
                    os.remove(checkpoint_path)
                elif os.path.exists(checkpoint_path):
                    print(f"[INFO] Alguns clientes falharam. A memória foi mantida para a próxima tentativa ({len(relatorio_notas)} emitidos).")
            except: pass
        else:
            print("[INFO] Nenhuma nota foi emitida com sucesso nesta sessão.")
        print("="*60)

        input("\nPressione Enter para fechar o navegador...")
        browser.close()

if __name__ == "__main__":
    emit_nfse_batch()
