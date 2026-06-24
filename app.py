import sys
import threading
import time
import webbrowser
import pandas as pd
from unittest.mock import patch
from flask import Flask, render_template, request, jsonify
import glob
import os

# Garantir segurança cibernética: remover qualquer variável com prefixo NEXT_PUBLIC_
# para evitar exposição acidental no lado do cliente
for key in list(os.environ.keys()):
    if key.upper().startswith("NEXT_PUBLIC_"):
        del os.environ[key]

# Add the current directory to the path so we can import the batch script
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the main automation function
from batch_emit_nfse import emit_nfse_batch

app = Flask(__name__)

# Configurações de cookies de sessão seguros
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE='Lax'
)


# Global state to track automation status
automation_status = {
    "is_running": False,
    "is_paused": False,
    "message": "Aguardando início...",
    "error": None
}

# Threading event for pause/resume control
pause_event = threading.Event()
pause_event.set() # Set to "Running" by default

# Store the active dataframe globally so it can be updated during pause
active_df = None

# Advanced Python mock to intercept the 'input()' calls inside the batch script
# without needing to rewrite the existing file.
class MockInput:
    def __init__(self, aliquota):
        self.aliquota = aliquota
        self.call_count = 0

    def __call__(self, prompt=""):
        self.call_count += 1
        print(f"[GUI-Intercept] Automação pediu input: '{prompt}'")
        
        # O script original pede:
        # 1. Recuperar Histórico (S/N)
        # 2. Alíquota
        # 3. Alterar dados pendentes (S/N)
        
        if "CONTINUAR de onde parou" in prompt:
            return "N"  # Padrão: Não recuperar ou Tratar via GUI depois
            
        if "alíquota" in prompt.lower():
            return self.aliquota
            
        if "alterar valores" in prompt.lower():
            return "N"  # Padrão: Não fazemos substituições via terminal agora
        
        if "enter" in prompt.lower() or "pressione" in prompt.lower():
            return "" # Evita travar no final esperando Enter no terminal
            
        return ""

def run_automation_thread(aliquota, df_customizado, headless=False):
    """Runs the Playwright automation in a separate thread so it doesn't block Flask"""
    global automation_status, active_df
    active_df = df_customizado
    automation_status["is_running"] = True
    automation_status["is_paused"] = False
    automation_status["message"] = "Automação iniciada. Não feche o navegador!"
    automation_status["error"] = None
    
    # Keep original builtins
    import builtins
    original_input = builtins.input
    original_print = builtins.print
    
    # Smart Interceptor for Print: This is our "Pause Hook"
    def mocked_print(*args, **kwargs):
        # If we are paused, wait here before logging/stepping further
        if not pause_event.is_set():
            automation_status["is_paused"] = True
            automation_status["message"] = "⏸️ ROBÔ PAUSADO. Aguardando sua retomada..."
        
        pause_event.wait() # Blocking call if event is cleared
        
        automation_status["is_paused"] = False
        original_print(*args, **kwargs)

    try:
        # Override functions globally for the thread duration
        builtins.input = MockInput(aliquota)
        builtins.print = mocked_print
        
        # Intercepta a leitura do Excel do pandas para que o robô use os dados editados na UI
        with patch('pandas.read_excel') as mock_read:
            mock_read.return_value = active_df
            
            # Call the exact same function that START_AUTOMATION.bat calls
            emit_nfse_batch(headless=headless)
        
        automation_status["message"] = "✅ Emissão finalizada com sucesso! Verifique a pasta 'evidencias'."
        
    except Exception as e:
        automation_status["error"] = str(e)
        automation_status["message"] = f"❌ Erro na execução: {str(e)}"
    finally:
        # Restore original builtins
        builtins.input = original_input
        builtins.print = original_print
        automation_status["is_running"] = False
        automation_status["is_paused"] = False
        pause_event.set() # Ensure doesn't stay blocked


@app.before_request
def block_sensitive_files():
    path = request.path.lower()
    # Bloqueia arquivos de código-fonte, mapas de código (.map) e configurações sensíveis
    blocked_extensions = ['.map', '.tsx', '.ts', '.jsx', '.env', '.py', '.yml', '.yaml', '.git']
    if any(path.endswith(ext) for ext in blocked_extensions) or '/.git' in path or 'node_modules' in path:
        return jsonify({"error": "Acesso proibido. Arquivo protegido."}), 403

@app.after_request
def add_security_headers(response):
    # Content Security Policy (CSP) estrito
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data: https://via.placeholder.com; "
        "connect-src 'self';"
    )
    # Proteção contra Clickjacking
    response.headers['X-Frame-Options'] = 'DENY'
    # Evitar farejamento de tipo MIME (MIME-sniffing)
    response.headers['X-Content-Type-Options'] = 'nosniff'
    # Controle de Referência (Referrer Policy)
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    
    # Restrição rígida de CORS (permitir apenas o próprio localhost)
    origin = request.headers.get('Origin')
    if origin and (origin.startswith('http://localhost:') or origin.startswith('http://127.0.0.1:')):
        response.headers['Access-Control-Allow-Origin'] = origin
        response.headers['Access-Control-Allow-Credentials'] = 'true'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        
    return response

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/status', methods=['GET'])
def get_status():
    global automation_status
    return jsonify(automation_status)

@app.route('/api/report', methods=['GET'])
def get_report():
    try:
        # Pega a pasta evidencias e encontra a mais recente
        current_dir = os.path.dirname(os.path.abspath(__file__))
        evidencias_dir = os.path.join(current_dir, "evidencias", "execucao_*")
        
        pastas = sorted(glob.glob(evidencias_dir), reverse=True)
        if not pastas:
            return jsonify({"success": False, "message": "Nenhuma pasta de evidências encontrada."})
            
        relatorio_path = os.path.join(pastas[0], "relatorio_final.txt")
        
        if os.path.exists(relatorio_path):
            with open(relatorio_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return jsonify({"success": True, "report": content})
        else:
            return jsonify({"success": False, "message": "Relatório final não foi gerado."})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/data', methods=['GET'])
def get_data():
    """Reads the Excel file and returns it as JSON for the review panel"""
    try:
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "clientes.xlsx")
        if not os.path.exists(excel_path):
            return jsonify({"success": False, "message": "Planilha 'clientes.xlsx' não encontrada."})
            
        df = pd.read_excel(excel_path)
        # Convert to list of dicts for the frontend
        data = df.to_dict(orient='records')
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/start', methods=['POST'])
def start_automation():
    global automation_status, active_df
    
    if automation_status["is_running"]:
        return jsonify({"success": False, "message": "Automação já está rodando!"})
        
    data = request.json
    aliquota = data.get('aliquota', '2.24')
    clientes_editados = data.get('clientes', [])
    headless = data.get('headless', False)
    
    if not clientes_editados:
        return jsonify({"success": False, "message": "Nenhum cliente enviado para processamento."})
        
    # Reset event to running state
    pause_event.set()
    
    # Converte os dados editados da UI de volta para um DataFrame do pandas
    df_customizado = pd.DataFrame(clientes_editados)
    
    # Começa a thread em background para não travar o site
    thread = threading.Thread(target=run_automation_thread, args=(aliquota, df_customizado, headless))
    thread.daemon = True
    thread.start()
    
    return jsonify({"success": True, "message": "Processo iniciado!"})

@app.route('/api/pause', methods=['POST'])
def pause_automation():
    global automation_status
    if not automation_status["is_running"]:
        return jsonify({"success": False, "message": "Automação não está rodando."})
    
    pause_event.clear() # Robot will block next time it calls print()
    return jsonify({"success": True, "message": "Sinal de pausa enviado."})

@app.route('/api/resume', methods=['POST'])
def resume_automation():
    global automation_status, active_df
    if not automation_status["is_running"]:
        return jsonify({"success": False, "message": "Automação não está rodando."})
    
    # Se o usuário enviou atualizações (adendos) durante a pausa
    data = request.json
    if data and "clientes" in data:
        # Atualizamos o Dataframe ATIVO em tempo real
        # O robô, ao continuar o loop, pegará os novos valores
        new_df = pd.DataFrame(data["clientes"])
        for col in new_df.columns:
            active_df[col] = new_df[col]
        print("[GUI-Update] Dados atualizados durante a pausa.")

    pause_event.set() # Unblock robot
    return jsonify({"success": True, "message": "Retomando execução..."})

def open_browser():
    # Wait a tiny bit for the server to spin up
    time.sleep(1.5)
    url = "http://localhost:5000/"
    
    import subprocess
    import os
    
    # Tenta abrir o Opera forçando nova janela (específico para o PC atual)
    opera_paths = [
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Opera', 'launcher.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Opera GX', 'launcher.exe')
    ]
    
    for path in opera_paths:
        if os.path.exists(path):
            subprocess.Popen([path, '--new-window', url])
            return

    try:
        # Verifica se chrome está no PATH
        subprocess.run("where chrome", shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        os.system(f'start chrome --new-window {url}')
        return
    except Exception:
        pass

    try:
        # Se não tiver Chrome, tenta o Edge
        subprocess.run("where msedge", shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        os.system(f'start msedge --new-window {url}')
        return
    except Exception:
        pass

    # Tenta via cmd genérico como última esperança para o Opera
    os.system(f'start opera --new-window {url}')

if __name__ == '__main__':
    # Auto-open the UI when the user clicks the script
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(port=5000, debug=False)
