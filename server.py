import os
import subprocess
import sys
from flask import Flask, render_template, request, Response, stream_with_context, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import json
import threading
import time

app = Flask(__name__)

# Configurações
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

# Sistema de Heartbeat
ACTIVE_SESSIONS = {}  # {session_id: {timestamp, last_heartbeat, user_agent}}
HEARTBEAT_TIMEOUT = 4  # 4 segundos (2s intervalo + 2s de margem)
SESSIONS_EVER_EXISTED = False  # Flag para rastrear se houve alguma sessão

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_and_shutdown_if_needed():
    """Verifica se deve desligar o servidor"""
    global SESSIONS_EVER_EXISTED
    
    now = datetime.now()
    
    # Limpa sessões expiradas
    for session_id, session_data in list(ACTIVE_SESSIONS.items()):
        if (now - session_data['last_heartbeat']).total_seconds() > HEARTBEAT_TIMEOUT:
            del ACTIVE_SESSIONS[session_id]
            print(f"[{now.strftime('%H:%M:%S')}] ❌ Sessão expirada/fechada: {session_id}")
    
    # Log de estado atual
    if len(ACTIVE_SESSIONS) > 0:
        print(f"[{now.strftime('%H:%M:%S')}] 💓 Heartbeat OK - {len(ACTIVE_SESSIONS)} sessões ativas.")
    
    # Se já houve sessões registradas e agora não há nenhuma, encerra o servidor
    if SESSIONS_EVER_EXISTED and len(ACTIVE_SESSIONS) == 0:
        print(f"\n[{now.strftime('%H:%M:%S')}] 🛑 Sem sessões ativas (Navegador fechado) - Encerrando servidor em 1s...\n")
        # Usar threading para não bloquear a resposta
        threading.Thread(target=lambda: (time.sleep(1), os.kill(os.getpid(), 15)), daemon=True).start()

def inactivity_monitor():
    """Thread em segundo plano que monitora inatividade sem depender de novos heartbeats"""
    print("Monitor de inatividade iniciado (Verificação a cada 2s).")
    while True:
        try:
            check_and_shutdown_if_needed()
        except Exception as e:
            print(f"Erro no monitor: {e}")
        time.sleep(2) # Verifica a cada 2 segundos agora

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/heartbeat', methods=['POST'])
def heartbeat():
    """Recebe e processa heartbeat da página aberta"""
    global SESSIONS_EVER_EXISTED
    
    try:
        data = request.get_json()
        session_id = data.get('session_id')
        
        if not session_id:
            return jsonify({'error': 'session_id não fornecido'}), 400
        
        # Registra ou atualiza a sessão
        ACTIVE_SESSIONS[session_id] = {
            'timestamp': datetime.now().isoformat(),
            'last_heartbeat': datetime.now(),
            'user_agent': request.headers.get('User-Agent', 'Unknown')
        }
        
        SESSIONS_EVER_EXISTED = True
        
        # Verifica se deve encerrar (se há sessões expiradas)
        check_and_shutdown_if_needed()
        
        return jsonify({
            'status': 'ok',
            'server_time': datetime.now().isoformat(),
            'active_sessions': len(ACTIVE_SESSIONS)
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/heartbeat/logout', methods=['POST'])
def heartbeat_logout():
    """Remove a sessão imediatamente (chamado no unload da página)"""
    try:
        data = request.json
        session_id = data.get('session_id')
        if session_id in ACTIVE_SESSIONS:
            del ACTIVE_SESSIONS[session_id]
            print(f"Sessão encerrada via logout: {session_id}")
        
        # Verifica se deve encerrar agora
        check_and_shutdown_if_needed()
        return jsonify({'status': 'ok'}), 200
    except:
        return jsonify({'status': 'error'}), 500

@app.route('/sessions/status', methods=['GET'])
def sessions_status():
    """Retorna o status das sessões ativas"""
    # Remove sessões expiradas
    now = datetime.now()
    expired_sessions = []
    
    for session_id, session_data in list(ACTIVE_SESSIONS.items()):
        last_beat = session_data['last_heartbeat']
        if (now - last_beat).total_seconds() > HEARTBEAT_TIMEOUT:
            expired_sessions.append(session_id)
            del ACTIVE_SESSIONS[session_id]
    
    return jsonify({
        'active_sessions': len(ACTIVE_SESSIONS),
        'sessions': {
            session_id: {
                'timestamp': data['timestamp'],
                'seconds_since_heartbeat': (now - data['last_heartbeat']).total_seconds()
            }
            for session_id, data in ACTIVE_SESSIONS.items()
        },
        'expired_sessions': len(expired_sessions),
        'server_time': datetime.now().isoformat()
    }), 200

@app.route('/sessions/active-count', methods=['GET'])
def active_count():
    """Retorna apenas a quantidade de sessões ativas"""
    # Remove sessões expiradas
    now = datetime.now()
    for session_id, session_data in list(ACTIVE_SESSIONS.items()):
        if (now - session_data['last_heartbeat']).total_seconds() > HEARTBEAT_TIMEOUT:
            del ACTIVE_SESSIONS[session_id]
    
    return jsonify({
        'active_count': len(ACTIVE_SESSIONS),
        'server_time': datetime.now().isoformat()
    }), 200

@app.route('/run-script', methods=['POST'])
def run_script():
    file = request.files.get('file')
    sheet_name = request.form.get('sheet', 'PESSOAS')

    if not file or not allowed_file(file.filename):
        return Response("Erro: Arquivo inválido ou não enviado.", status=400)

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Caminho absoluto para o script PowerShell
    script_path = os.path.abspath("Populate-SharePointList.ps1")
    
    # Comando para executar o PowerShell
    # -ExecutionPolicy Bypass é crucial para evitar bloqueios
    cmd = [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-File", script_path,
        "-ExcelPath", file_path,
        "-SheetName", sheet_name
    ]

    def generate():
        yield f"Arquivo recebido: {filename}\n"
        yield f"Aba selecionada: {sheet_name}\n"
        yield f"Executando script PowerShell...\n{'-'*30}\n"

        try:
            # Executa o processo e captura stdout/stderr em tempo real
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, # Redireciona stderr para stdout
                text=True,
                bufsize=1, # Line buffered
                universal_newlines=True
            )

            # Lê linha a linha e envia para o navegador
            for line in process.stdout:
                yield line

            process.wait()
            
            if process.returncode == 0:
                yield f"\n{'-'*30}\n[SUCESSO] Processo finalizado com código 0.\n"
            else:
                yield f"\n{'-'*30}\n[ERRO] Processo finalizado com código {process.returncode}.\n"

        except Exception as e:
            yield f"\n[ERRO DE EXECUÇÃO]: {str(e)}\n"
        finally:
            # Limpeza (opcional: remover arquivo após uso)
            # if os.path.exists(file_path):
            #     os.remove(file_path)
            pass

    return Response(stream_with_context(generate()), mimetype='text/plain')

if __name__ == '__main__':
    # Inicia monitor de inatividade em background
    threading.Thread(target=inactivity_monitor, daemon=True).start()
    
    # Roda o servidor acessível localmente
    print("Servidor rodando em http://localhost:5000")
    app.run(debug=False, host='0.0.0.0', port=5000)
