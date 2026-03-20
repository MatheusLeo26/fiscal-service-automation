import sys
import threading
import time
import webbrowser
import pandas as pd
from unittest.mock import patch
from flask import Flask, render_template, request, jsonify

# Add the current directory to the path so we can import the batch script
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the main automation function
from batch_emit_nfse import emit_nfse_batch

app = Flask(__name__)

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
        
        return ""

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


@app.route('/')
def index():
    return render_template('index.html')

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
    url = "http://127.0.0.1:5000/"
    # Forçar abertura em NOVA JANELA no Windows (Chrome/Edge/Default)
    # O comando 'start' com --new-window é mais garantido que webbrowser.open_new
    try:
        # Tenta Chrome primeiro que é o padrão do robô
        os.system(f'start chrome --new-window {url}')
    except:
        # Fallback para o comportamento padrão do sistema
        webbrowser.open_new(url)

if __name__ == '__main__':
    # Auto-open the UI when the user clicks the script
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(port=5000, debug=False)
