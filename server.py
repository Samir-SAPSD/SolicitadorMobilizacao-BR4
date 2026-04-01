import os
import subprocess
import sys
from flask import Flask, render_template, request, Response, stream_with_context, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import json
import re
import threading
import time

app = Flask(__name__)
app.json.ensure_ascii = False

# Configurações
UPLOAD_FOLDER = 'templates'
ALLOWED_EXTENSIONS = {'xlsx'}

# Sistema de Heartbeat
ACTIVE_SESSIONS = {}  # {session_id: {timestamp, last_heartbeat, user_agent}}
HEARTBEAT_TIMEOUT = 25  # 25 segundos (2s intervalo + 2s de margem)
SESSIONS_EVER_EXISTED = False  # Flag para rastrear se houve alguma sessão
ACTIVE_JOBS = 0
JOBS_LOCK = threading.Lock()

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
TEMPLATE_FILENAME = 'ModeloSolicitacaoMob.xlsx'
TEMPLATE_STATUS_FILE = os.path.join('templates', 'template_update_status.json')


def start_job():
    global ACTIVE_JOBS
    with JOBS_LOCK:
        ACTIVE_JOBS += 1


def end_job():
    global ACTIVE_JOBS
    with JOBS_LOCK:
        if ACTIVE_JOBS > 0:
            ACTIVE_JOBS -= 1


def build_powershell_command(script_path, file_path, sheet_name):
    """Executa scripts PowerShell com stdout em UTF-8 para preservar acentuação."""
    utf8_preamble = (
        "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [Console]::OutputEncoding; "
        "chcp 65001 > $null; "
    )
    escaped_script_path = script_path.replace("'", "''")
    escaped_file_path = file_path.replace("'", "''")
    escaped_sheet_name = sheet_name.replace("'", "''")

    return [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command",
        (
            f"{utf8_preamble}& '{escaped_script_path}' "
            f"-ExcelPath '{escaped_file_path}' -SheetName '{escaped_sheet_name}'"
        )
    ]


def build_powershell_template_update_command(script_path, template_path):
    """Executa script PowerShell de atualização do template com stdout em UTF-8."""
    utf8_preamble = (
        "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [Console]::OutputEncoding; "
        "chcp 65001 > $null; "
    )
    escaped_script_path = script_path.replace("'", "''")
    escaped_template_path = template_path.replace("'", "''")

    return [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command",
        (
            f"{utf8_preamble}& '{escaped_script_path}' "
            f"-TemplatePath '{escaped_template_path}'"
        )
    ]


def get_template_status():
    if not os.path.exists(TEMPLATE_STATUS_FILE):
        return {'last_updated': None}

    try:
        with open(TEMPLATE_STATUS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return {'last_updated': data.get('last_updated')}
    except Exception:
        return {'last_updated': None}


def save_template_status(iso_datetime):
    payload = {'last_updated': iso_datetime}
    with open(TEMPLATE_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def decode_powershell_output(raw_output):
    """Decodifica stdout do Windows PowerShell 5.1 preservando acentuação."""
    utf8_text = raw_output.decode('utf-8', errors='replace')
    cp1252_text = raw_output.decode('cp1252', errors='replace')
    repaired_cp1252 = repair_mojibake(cp1252_text)

    utf8_score = score_decoded_text(utf8_text)
    repaired_cp1252_score = score_decoded_text(repaired_cp1252)

    if repaired_cp1252_score < utf8_score:
        return repaired_cp1252
    return utf8_text


def repair_mojibake(text):
    """Corrige trechos UTF-8 lidos como cp1252 sem afetar texto já correto."""
    previous_text = text
    for _ in range(3):
        repaired_text = repair_mojibake_once(previous_text)
        if score_decoded_text(repaired_text) >= score_decoded_text(previous_text):
            return previous_text
        previous_text = repaired_text
    return previous_text


def repair_mojibake_once(text):
    parts = re.split(r'(\s+)', text)
    repaired_parts = []
    for part in parts:
        if not part or part.isspace() or not has_mojibake_markers(part):
            repaired_parts.append(part)
            continue

        try:
            repaired_candidate = part.encode('cp1252', errors='strict').decode('utf-8', errors='strict')
        except (UnicodeEncodeError, UnicodeDecodeError):
            repaired_parts.append(part)
            continue

        if score_decoded_text(repaired_candidate) <= score_decoded_text(part):
            repaired_parts.append(repaired_candidate)
        else:
            repaired_parts.append(part)

    return ''.join(repaired_parts)


def has_mojibake_markers(text):
    return any(marker in text for marker in ('Ã', 'Â', 'â'))


def score_decoded_text(text):
    replacement_penalty = text.count('�') * 10
    mojibake_penalty = sum(text.count(marker) for marker in ('Ã', 'Â', 'â')) * 6
    return replacement_penalty + mojibake_penalty

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
            print(f"[{now.strftime('%H:%M:%S')}] Sessão expirada/fechada: {session_id}")
    
    # Log de estado atual
    if len(ACTIVE_SESSIONS) > 0:
        print(f"[{now.strftime('%H:%M:%S')}] Heartbeat OK - {len(ACTIVE_SESSIONS)} sessões ativas.")
    
    # Não encerra o servidor enquanto houver jobs ativos (validação/upload em andamento)
    if ACTIVE_JOBS > 0:
        return

    # Se já houve sessões registradas e agora não há nenhuma, encerra o servidor
    if SESSIONS_EVER_EXISTED and len(ACTIVE_SESSIONS) == 0:
        print(f"\n[{now.strftime('%H:%M:%S')}] Sem sessões ativas (Navegador fechado) - Encerrando servidor em 1s...\n")
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

@app.route('/download-template')
def download_template():
    return send_from_directory(app.config['UPLOAD_FOLDER'], TEMPLATE_FILENAME, as_attachment=True)


@app.route('/template-update-status', methods=['GET'])
def template_update_status():
    return jsonify(get_template_status()), 200


@app.route('/update-template', methods=['POST'])
def update_template():
    start_job()
    try:
        template_path = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], TEMPLATE_FILENAME))
        if not os.path.exists(template_path):
            return jsonify({'status': 'error', 'message': 'Template não encontrado.'}), 404

        script_path = os.path.abspath('Update-ExcelTemplateChoices.ps1')
        if not os.path.exists(script_path):
            return jsonify({'status': 'error', 'message': 'Script de atualização não encontrado.'}), 500

        cmd = build_powershell_template_update_command(script_path, template_path)

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=False
        )
        raw_output = process.communicate()[0]
        process.wait()
        full_output = decode_powershell_output(raw_output)

        if process.returncode != 0:
            return jsonify({
                'status': 'error',
                'message': 'Falha ao atualizar template.',
                'log': full_output
            }), 500

        updated_at = datetime.now().isoformat()
        save_template_status(updated_at)

        return jsonify({
            'status': 'success',
            'message': 'Template atualizado com sucesso.',
            'last_updated': updated_at,
            'log': full_output
        }), 200
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Erro ao atualizar template: {str(e)}'
        }), 500
    finally:
        end_job()

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

@app.route('/validate', methods=['POST'])
def validate():
    """Fase 1: Executa validação separada antes do upload"""
    start_job()
    try:
        file = request.files.get('file')
        sheet_name = request.form.get('sheet', 'PESSOAS')
        sheet_name = (sheet_name or 'PESSOAS').strip() or 'PESSOAS'

        if not file or not allowed_file(file.filename):
            return jsonify({'status': 'error', 'errors': ['Arquivo inválido ou não enviado.']}), 400

        filename = secure_filename(file.filename)
        if not filename:
            return jsonify({'status': 'error', 'errors': ['Nome de arquivo inválido.']}), 400

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        script_path = os.path.abspath("Validate-ExcelData.ps1")
        cmd = build_powershell_command(script_path, file_path, sheet_name)

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=False
        )
        raw_output = process.communicate()[0]
        process.wait()
        full_output = decode_powershell_output(raw_output)

        # Extrair JSON de validação da saída
        start_marker = '---VALIDATION_JSON_START---'
        end_marker = '---VALIDATION_JSON_END---'
        
        if start_marker in full_output and end_marker in full_output:
            start_idx = full_output.index(start_marker) + len(start_marker)
            end_idx = full_output.index(end_marker)
            json_str = full_output[start_idx:end_idx].strip()
            result = json.loads(json_str)
            
            # Adicionar log da validação para debug
            result['log'] = full_output
            
            # Incluir filename para a fase 2 reutilizar
            result['filename'] = filename
            return jsonify(result), 200
        else:
            return jsonify({
                'status': 'error',
                'errors': ['Não foi possível obter resultado da validação.'],
                'log': full_output,
                'filename': filename
            }), 200

    except Exception as e:
        return jsonify({
            'status': 'error',
            'errors': [f'Erro ao executar validação: {str(e)}']
        }), 500
    finally:
        end_job()


@app.route('/run-script', methods=['POST'])
def run_script():
    """Fase 2: Upload para SharePoint (usa o arquivo já salvo pela validação)"""
    data = request.get_json()
    if not data:
        return Response("Erro: Dados não enviados.", status=400)

    filename = data.get('filename', '')
    sheet_name = data.get('sheet', 'PESSOAS')

    if not filename or not allowed_file(filename):
        return Response("Erro: Arquivo inválido.", status=400)

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    if not os.path.exists(file_path):
        return Response("Erro: Arquivo não encontrado. Execute a validação primeiro.", status=400)

    # Caminho absoluto para o script PowerShell
    script_path = os.path.abspath("Populate-SharePointList.ps1")
    
    # Comando para executar o PowerShell
    # -ExecutionPolicy Bypass é crucial para evitar bloqueios
    cmd = build_powershell_command(script_path, file_path, sheet_name)

    def generate():
        start_job()
        yield f"Arquivo recebido: {filename}\n"
        yield f"Aba selecionada: {sheet_name}\n"
        yield f"Executando script PowerShell...\n{'-'*30}\n"

        try:
            # Executa o processo e captura stdout/stderr em tempo real
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, # Redireciona stderr para stdout
                bufsize=1, # Line buffered
                universal_newlines=False
            )

            detected_error_in_output = False
            error_markers = [
                "FALHA CRÍTICA",
                "UPLOAD CANCELADO",
                "--- RESULT: ERROR ---",
                "Write-Error",
                "Erro ao adicionar item",
            ]

            # Lê linha a linha e envia para o navegador
            while True:
                chunk = process.stdout.readline()
                if not chunk:
                    break
                decoded = decode_powershell_output(chunk)
                if any(marker in decoded for marker in error_markers):
                    detected_error_in_output = True
                yield decoded

            process.wait()
            if process.returncode == 0 and not detected_error_in_output:
                yield f"\n{'-'*30}\n[SUCESSO] Processo finalizado com código 0.\n"
            else:
                yield f"\n{'-'*30}\n[ERRO] Processo finalizado com código {process.returncode}.\n"

        except Exception as e:
            yield f"\n[ERRO DE EXECUÇÃO]: {str(e)}\n"
        finally:
            # Limpeza (opcional: remover arquivo após uso)
            # if os.path.exists(file_path):
            #     os.remove(file_path)
            end_job()

    return Response(stream_with_context(generate()), content_type='text/plain; charset=utf-8')

if __name__ == '__main__':
    # Inicia monitor de inatividade em background
    threading.Thread(target=inactivity_monitor, daemon=True).start()
    
    # Roda o servidor acessível localmente
    print("Servidor rodando em http://localhost:5000")
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
