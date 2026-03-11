import sys
import threading
import time
import webbrowser
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
    "message": "Aguardando início...",
    "error": None
}

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

def run_automation_thread(aliquota):
    """Runs the Playwright automation in a separate thread so it doesn't block Flask"""
    global automation_status
    automation_status["is_running"] = True
    automation_status["message"] = "Automação iniciada. Não feche o navegador!"
    automation_status["error"] = None
    
    # Keep original builtins
    import builtins
    original_input = builtins.input
    
    try:
        # Override the input() function globally for the thread duration
        builtins.input = MockInput(aliquota)
        
        # Call the exact same function that START_AUTOMATION.bat calls
        emit_nfse_batch()
        
        automation_status["message"] = "Emissão finalizada com sucesso! Verifique a pasta 'evidencias'."
        
    except Exception as e:
        automation_status["error"] = str(e)
        automation_status["message"] = f"Erro na execução: {str(e)}"
    finally:
        # Restore original builtins
        builtins.input = original_input
        automation_status["is_running"] = False


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/status', methods=['GET'])
def get_status():
    return jsonify(automation_status)

@app.route('/api/start', methods=['POST'])
def start_automation():
    global automation_status
    
    if automation_status["is_running"]:
        return jsonify({"success": False, "message": "Automação já está rodando!"})
        
    data = request.json
    aliquota = data.get('aliquota', '2.24')
    
    # Começa a thread em background para não travar o site
    thread = threading.Thread(target=run_automation_thread, args=(aliquota,))
    thread.daemon = True
    thread.start()
    
    return jsonify({"success": True, "message": "Processo iniciado!"})

def open_browser():
    # Wait a tiny bit for the server to spin up
    time.sleep(1.5)
    webbrowser.open_new("http://127.0.0.1:5000/")

if __name__ == '__main__':
    # Auto-open the UI when the user clicks the script
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(port=5000, debug=False)
