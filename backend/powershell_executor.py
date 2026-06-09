"""PowerShell command building and process execution utilities."""

import os
import subprocess


UTF8_PREAMBLE = (
    '[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); '
    '[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); '
    '$OutputEncoding = [Console]::OutputEncoding; '
    'chcp 65001 > $null; '
)


def build_powershell_command(script_path: str, file_path: str, sheet_name: str) -> list:
    """Executa scripts PowerShell com stdout em UTF-8 para preservar acentuação."""
    escaped_script_path = script_path.replace("'", "''")
    escaped_file_path = file_path.replace("'", "''")
    escaped_sheet_name = sheet_name.replace("'", "''")

    return [
        'powershell.exe',
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-Command',
        (
            f"{UTF8_PREAMBLE}& '{escaped_script_path}' "
            f"-ExcelPath '{escaped_file_path}' -SheetName '{escaped_sheet_name}'"
        )
    ]


def build_powershell_populate_command(script_path: str, file_path: str) -> list:
    """Executa Populate-SharePointList.ps1 sem o parâmetro -SheetName (processamento unificado)."""
    escaped_script_path = script_path.replace("'", "''")
    escaped_file_path = file_path.replace("'", "''")

    return [
        'powershell.exe',
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-Command',
        (
            f"{UTF8_PREAMBLE}& '{escaped_script_path}' "
            f"-ExcelPath '{escaped_file_path}'"
        )
    ]


def build_powershell_template_update_command(script_path: str, template_path: str) -> list:
    """Executa script PowerShell de atualização do template com stdout em UTF-8."""
    escaped_script_path = script_path.replace("'", "''")
    escaped_template_path = template_path.replace("'", "''")

    return [
        'powershell.exe',
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-Command',
        (
            f"{UTF8_PREAMBLE}& '{escaped_script_path}' "
            f"-TemplatePath '{escaped_template_path}'"
        )
    ]


def popen_hidden_kwargs() -> dict:
    """Oculta janelas de terminal ao executar subprocessos no Windows."""
    if os.name != 'nt':
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return {
        'startupinfo': startupinfo,
        'creationflags': subprocess.CREATE_NO_WINDOW
    }
