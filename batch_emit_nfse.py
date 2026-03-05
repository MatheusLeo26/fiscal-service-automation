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

def emit_nfse_batch():
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
        print("      RETOMADA DE EXECUÇÃO")
        print("!"*40)
        resumo = input("Encontrei uma execução anterior incompleta. Deseja continuar de onde parou? (S/N): ").strip().upper()
        if resumo == 'S':
            try:
                with open(checkpoint_path, "r", encoding="utf-8") as f:
                    cnpj_concluidos = json.load(f)
                print(f"[INFO] Retomando. {len(cnpj_concluidos)} notas já foram emitidas.")
            except:
                print("[WARN] Erro ao ler progresso anterior. Iniciando do zero.")
        else:
            # If not resuming, we can delete the old checkpoint
            try: os.remove(checkpoint_path)
            except: pass
    
    # Filter pending clients for the menu and loop
    df_pendentes = df[~df['CNPJ'].isin(cnpj_concluidos)].copy().reset_index(drop=True)
    
    if df_pendentes.empty:
        print("[INFO] Não há notas pendentes para emitir. Se desejar reiniciar tudo, apague o arquivo 'progresso.json'.")
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
        browser = p.chromium.launch(headless=False) # Change to True later if preferred
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
                    print("[INFO] Navegando para o menu de emissão...")
                    page.locator("text=Serviços Prestados").click()
                    page.locator("text=Emitir NFS-e").click()
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
                    
                    # Try to select by label precisely
                    page.select_option(selector_atividade, label="17.19 / 692060100 - ATIVIDADES DE CONTABILIDADE")
                except Exception as e_step1:
                    print(f"[WARN] Falha na seleção por label: {str(e_step1)}. Tentando por índice...")
                    page.select_option("select#atividadeServico", index=3) # Fallback

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
                selector_busca_tomador = "input#buscarTomador, input[name='cnpjTomador'], input[placeholder*='Pesquisar']"
                page.wait_for_selector(selector_busca_tomador, timeout=10000)
                page.fill(selector_busca_tomador, cnpj)
                page.click("button:has-text('Pesquisar')")
                time.sleep(3)
                
                # Clicar no nome do tomador na lista de resultados
                selector_tomador_link = f"xpath=//*[contains(text(), '{nome_empresa}') or contains(text(), '{cnpj.strip()}')]"
                page.wait_for_selector(selector_tomador_link, timeout=15000)
                page.locator(selector_tomador_link).first.click()
                print("[INFO] Tomador selecionado.")

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
                
                # Critical: Clear field first (Ctrl+A + Backspace) as per site mask
                page.click(selector_aliq)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                
                # Fill clean value
                aliquota_limpa = aliquota.replace(",", ".")
                page.type(selector_aliq, aliquota_limpa)
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
                    sidebar_selector = "text=Serviços Prestados, .menu-item:has-text('Serviços Prestados')"
                    print("[INFO] Clicando em 'Serviços Prestados' no menu lateral...")
                    page.locator(sidebar_selector).scroll_into_view_if_needed()
                    page.click(sidebar_selector)
                    time.sleep(1)
                    
                    # 2. Click "Emitir NFS-e" card/link
                    emitir_selector = "text=Emitir NFS-e, .card-body:has-text('Emitir NFS-e')"
                    print("[INFO] Clicando em 'Emitir NFS-e' para iniciar novo formulário...")
                    page.locator(emitir_selector).scroll_into_view_if_needed()
                    page.click(emitir_selector)
                    
                    page.wait_for_load_state("networkidle")
                    print("[INFO] Próximo cliente em 3 segundos...")
                    time.sleep(3)
                except Exception as e_nav:
                    print(f"[WARN] Falha na navegação pelo menu: {str(e_nav)}. Tentando recarregar rota...")
                    page.goto("https://itu.giss.com.br/portal/home#/operacao/servicos-prestados/emitir-nfse")
                    page.wait_for_load_state("networkidle")

            except Exception as e_inner:
                print(f"[ERROR] Erro fatal no lote para {nome_empresa}: {str(e_inner)}")
                try:
                    if not page.is_closed():
                        page.screenshot(path=os.path.join(evidence_dir, f"erro_{cnpj}.png"))
                except:
                    pass
                # Recovery
                try:
                    if not page.is_closed():
                        page.goto("https://itu.giss.com.br/portal/home#/operacao/servicos-prestados/emitir-nfse")
                except:
                    break 

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
            
            # 8. Cleanup Progress if successful
            try:
                if os.path.exists(checkpoint_path):
                    os.remove(checkpoint_path)
            except: pass
        else:
            print("[INFO] Nenhuma nota foi emitida com sucesso nesta sessão.")
        print("="*60)

        input("\nPressione Enter para fechar o navegador...")
        browser.close()

if __name__ == "__main__":
    emit_nfse_batch()
