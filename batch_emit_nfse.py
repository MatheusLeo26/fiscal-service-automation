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
            
        # -- UPDATE: Handle pre-login popups (e.g., "EMPRESAS QUE EMITEM NFS-e POR WEBSERVICE")
        print("[INFO] Verificando popups pré-login...")
        try:
            # Wait up to 5 seconds for the popup 'OK' button to appear before trying to close
            try:
                page.wait_for_selector("text='OK', text='Ok'", timeout=5000)
            except:
                pass
                
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
                valor = f"{float(raw_valor):.2f}"
            else:
                valor = str(raw_valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
            
            nome_empresa = str(row[col_empresa]).strip()
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
                
                # Click field to open the Angular dropdown
                selector_nbs_input = "#ibs_nbs_value, input[name='nbsAuto'], input[placeholder*='NBS']"
                page.wait_for_selector(selector_nbs_input, timeout=15000)
                page.locator(selector_nbs_input).scroll_into_view_if_needed()
                page.click(selector_nbs_input)
                
                # Wait for the specific dropdown row to appear
                print("[INFO] Aguardando lista NBS aparecer...")
                # The user specifically mentioned the text "1.1302.21.00 Serviços de contabilidade"
                selector_nbs_option = ".angucomplete-row:has-text('1.1302.21.00'), .angucomplete-row:has-text('Serviços de contabilidade')"
                try:
                    page.wait_for_selector(selector_nbs_option, timeout=10000)
                    page.locator(selector_nbs_option).first.scroll_into_view_if_needed()
                    page.locator(selector_nbs_option).first.click()
                    print("[INFO] NBS selecionado com sucesso via clique.")
                except Exception as e_nbs:
                    print(f"[WARN] Lista NBS não expandiu. Tentando clique forçado...")
                    page.click(selector_nbs_input, force=True)
                    time.sleep(2)
                    page.locator(selector_nbs_option).first.click()

                # Passo 3: Dados do Tomador de Serviço
                print(f"[INFO] Passo 3: Pesquisando Tomador: {cnpj}...")
                selector_busca_tomador = "input#buscarTomador, input[name='cnpjTomador'], input[placeholder*='Pesquisar'], input[placeholder*='Tomador']"
                page.wait_for_selector(selector_busca_tomador, timeout=10000)
                
                # PASSO 1: CLICAR NO CAMPO DE DIGITAÇÃO E DIGITAR O CNPJ DO CLIENTE
                page.locator(selector_busca_tomador).click()
                page.fill(selector_busca_tomador, cnpj)
                
                # PASSO 2: CLICAR NO BOTÃO PESQUISAR LOGO AO LADO
                print("[INFO] Clicando no botão Pesquisar e aguardando resultado...")
                btn_pesquisar = page.locator("button:has-text('Pesquisar')").first
                btn_pesquisar.click()
                time.sleep(3) # Aguarda retorno da pesquisa (cascata Angular)
                
                # PASSO 3: DESCER A TELA PARA VER A OPÇÃO E DAR O CLIQUE OBRIGATÓRIO
                print("[INFO] Rolando a tela para visualizar a lista suspensa...")
                page.mouse.wheel(0, 300) # Rola fisicamente a tela para baixo para revelar a cascata
                time.sleep(1)
                
                try:
                    # Estratégia principal: Localizar pelo nome vindo da planilha
                    resultado_nome = page.get_by_text(nome_empresa, exact=False).first
                    resultado_nome.wait_for(state="visible", timeout=5000)
                    resultado_nome.click()
                    print("[INFO] Tomador selecionado com sucesso pelo nome da empresa.")
                except Exception:
                    # Estratégia de Palavras-Chave (Fuzzy): Se "EM" vs "DE" ou pequenas variações travarem, 
                    # tentamos um "aperto de mão" pelas primeiras palavras do nome.
                    try:
                        # Pega as primeiras 3 palavras significativas (maiores que 2 letras)
                        print(f"[INFO] Nome exato não bateu. Tentando busca por palavras-chave...")
                        keywords = [p for p in nome_empresa.split() if len(p) > 3][:3]
                        fuzzy_name = " ".join(keywords)
                        if fuzzy_name:
                            resultado_fuzzy = page.get_by_text(fuzzy_name, exact=False).first
                            resultado_fuzzy.wait_for(state="visible", timeout=5000)
                            resultado_fuzzy.click()
                            print(f"[INFO] Tomador selecionado via palavras-chave: '{fuzzy_name}'")
                        else:
                            raise Exception("Sem palavras-chave válidas.")
                    except:
                        print(f"[WARN] Busca por nome falhou. Recorrendo ao CNPJ formatado...")
                        # Fallback estratégico: O site exibe o CNPJ com pontuação na cascata (Ex: 33.726.493/0001-09)
                        cnpj_pad = cnpj.strip().zfill(14)
                        cnpj_formatado = f"{cnpj_pad[:2]}.{cnpj_pad[2:5]}.{cnpj_pad[5:8]}/{cnpj_pad[8:12]}-{cnpj_pad[12:]}"
                        
                        try:
                            # Se houver duplicidade de CNPJ (como no caso da Igreja Batista), 
                            # tentamos filtrar pelo locador que contenha parte do nome E o CNPJ
                            opcoes = page.get_by_text(cnpj_formatado, exact=False)
                            for i in range(opcoes.count()):
                                opt = opcoes.nth(i)
                                texto_opt = opt.inner_text().upper()
                                # Se alguma palavra do nome da empresa estiver nesse bloco do CNPJ, é o vencedor
                                if any(word.upper() in texto_opt for word in nome_empresa.split() if len(word) > 3):
                                    opt.click()
                                    print(f"[INFO] Tomador selecionado por CNPJ + validação de nome.")
                                    break
                            else:
                                # Se nada bater, clica no primeiro mesmo como última tentativa
                                opcoes.first.click()
                                print(f"[WARN] Tomador selecionado apenas por CNPJ (primeira opção).")
                        except:
                            print(f"[ERROR] Não foi possível selecionar o tomador. Verifique o print de erro.")
                            raise Exception("Falha definitiva ao selecionar Tomador.")
                
                time.sleep(2) # Aguarda o painel inferior de valores carregar

                # Passo 4: Valor do Serviço
                print(f"[INFO] Passo 4: Preenchendo Valor (R$ {valor})...")
                selector_valor = "input#valorServico, input[name='valorServico']"
                page.wait_for_selector(selector_valor, timeout=10000)
                page.fill(selector_valor, valor)

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

                # Passo 8: Alíquota
                print(f"[INFO] Passo 8: Preenchendo Alíquota ({aliquota})...")
                # Using ID found by subagent
                selector_aliq = "input#aliquotaValor, input[name='aliquota']"
                page.wait_for_selector(selector_aliq, timeout=10000)
                page.locator(selector_aliq).scroll_into_view_if_needed()
                
                # O site brasileiro normalmente usa vírgula para decimais na máscara
                aliquota_limpa = aliquota.replace(".", ",")
                
                # Critical: Clear field first (Ctrl+A + Backspace) as per site mask
                page.click(selector_aliq)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                time.sleep(0.5)
                
                # Fill clean value (usando type com delay simulando usuário para não bugar a máscara)
                page.locator(selector_aliq).press_sequentially(aliquota_limpa, delay=150)
                time.sleep(1)
                print(f"[INFO] Alíquota '{aliquota_limpa}' inserida.")

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
                    match = re.search(r"número:\s*(\d+)", texto_completo, re.IGNORECASE)
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
