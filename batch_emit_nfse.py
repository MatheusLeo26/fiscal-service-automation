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
        page.fill("input[name='cpf']", cpf)
        page.click("button:has-text('Avançar')")
        page.fill("input[name='password']", password)
        page.click("button:has-text('Acessar')")
        
        # Handle "Ir para o sistema" landing page
        try:
            page.wait_for_selector("button:has-text('Ir para o sistema')", timeout=5000)
            page.click("button:has-text('Ir para o sistema')")
        except:
            pass
            
        # 4. Handle Popups
        print("[INFO] Removendo avisos iniciais...")
        # Wait for either the dashboard or a popup
        page.wait_for_load_state("networkidle")
        
        # Close popups if they exist
        popups = page.query_selector_all("button[aria-label='Close'], .modal-header .close, button:has-text('OK')")
        for popup in popups:
            try:
                popup.click()
                time.sleep(1)
            except:
                pass

        # 5. Select Company
        print("[INFO] Selecionando empresa SGR CONTABIL...")
        page.fill("input[placeholder*='Pesquisar']", "SGR CONTABIL E ASSESSORIA LTDA")
        # Wait for the table to filter and the arrow to appear
        page.wait_for_selector("i.fa-arrow-right", timeout=10000)
        page.click("i.fa-arrow-right")
        
        page.wait_for_load_state("networkidle")
        
        # 6. Batch Loop
        for index, row in df_execucao.iterrows():
            cnpj = str(row['CNPJ']).strip()
            # Clean value (handle strings and numbers)
            raw_valor = row['VALOR']
            if isinstance(raw_valor, (int, float)):
                # Use dot for backend fill, format to 2 decimal places
                valor = f"{float(raw_valor):.2f}"
            else:
                # String cleanup: R$ 3.000,00 -> 3000.00
                valor = str(raw_valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
            
            descricao = str(row['Discriminação dos Serviços']).strip()
            nome_empresa = str(row['Nome da empresa']).strip()
            
            print(f"\n[PROCESS] Emitindo nota ({index+1}/{len(df)}): {nome_empresa} ({cnpj})")
            
            try:
                # Ensure we are in a clean state (Dashboard) or navigate via top menu
                page.locator("text=Serviços Prestados").click()
                page.locator("text=Emitir NFS-e").click()
                
                # Fill form
                page.wait_for_selector("select[name='atividade']", timeout=10000)
                page.select_option("select[name='atividade']", label="ATIVIDADES DE CONTABILIDADE")
                
                # NBS select
                page.fill("input[placeholder*='NBS']", "SERVIÇOS DE CONTABILIDADE")
                time.sleep(1)
                page.keyboard.press("Enter")
                
                # Tomador
                page.fill("input[name='cnpjTomador']", cnpj)
                page.click("button:has-text('Pesquisar')")
                
                # Wait for search results and try to click the company name
                # Using a broad selector that matches the company name text
                page.wait_for_selector(f"text={nome_empresa}", timeout=15000)
                page.click(f"text={nome_empresa}")
                
                # Values and Description
                page.wait_for_selector("input[name='valorServico']", timeout=10000)
                page.fill("input[name='valorServico']", valor)
                page.fill("textarea[name='discriminacao']", descricao)
                page.select_option("select[name='pisCofins']", label="NENHUM")
                
                # Next Step
                page.click("button:has-text('Próximo')")
                
                # Alíquota Step
                page.wait_for_selector("input[name='aliquota']", timeout=10000)
                page.fill("input[name='aliquota']", aliquota)
                page.click("button:has-text('Próximo')")
                
                # Review and Concluir
                page.wait_for_selector("button:has-text('Concluir')", timeout=10000)
                page.click("button:has-text('Concluir')")
                
                # Handle "Deseja visualizar a nota?" popup
                print("[INFO] Nota enviada. Aguardando confirmação...")
                try:
                    # The button usually says 'Não' for visualization
                    # Take success screenshot
                    # Capture Note Number first to use in screenshot name
                    page.wait_for_load_state("networkidle")
                    try:
                        texto_confirmacao = page.inner_text("body")
                        match = re.search(r"número:\s*(\d+)", texto_confirmacao, re.IGNORECASE)
                        numero_nota = match.group(1) if match else "Desconhecido"
                    except:
                        numero_nota = "Nao_Capturado"

                    # Clean company name for safe filename
                    nome_seguro = re.sub(r'[\\/*?:"<>|]', "", nome_empresa).replace(" ", "_")
                    nome_screenshot = f"sucesso_Nota_{numero_nota}_{nome_seguro}.png"
                    
                    page.screenshot(path=os.path.join(evidence_dir, nome_screenshot))
                    
                    # (Texto já capturado acima)

                    relatorio_notas.append({
                        "empresa": nome_empresa,
                        "valor": valor,
                        "numero": numero_nota
                    })

                    page.click("button:has-text('Não')")
                    print(f"[SUCCESS] Nota {numero_nota} emitida para {nome_empresa}")
                    
                    # 6.1 Save Progress
                    cnpj_concluidos.append(cnpj)
                    with open(checkpoint_path, "w", encoding="utf-8") as f:
                        json.dump(cnpj_concluidos, f)
                    
                    # Safety wait before next loop
                    time.sleep(2)
                except Exception as e:
                    print(f"[WARN] Erro ao confirmar sucesso ou clicar em 'Não': {str(e)}")
                    # Take screenshot to verify state
                    page.screenshot(path=os.path.join(evidence_dir, f"verificar_{cnpj}.png"))
                
            except Exception as e:
                print(f"[ERROR] Falha ao processar {nome_empresa}: {str(e)}")
                page.screenshot(path=os.path.join(evidence_dir, f"erro_{cnpj}.png"))
                # Try to force navigation back to home menu to resume with next
                page.goto("https://itu.giss.com.br/portal/home#/dashboard")
                time.sleep(3)
            
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
