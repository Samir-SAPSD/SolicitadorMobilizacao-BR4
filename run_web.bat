@echo off
setlocal enabledelayedexpansion

echo --- Localizando Python no Sistema ---
set "PY_EXE="
set "PYW_EXE="

:: 1. BUSCA PRIORITARIA: Caminho local do usuario (AppData)
for %%v in (315 314 313 312 311 310) do (
    if not defined PY_EXE (
        set "BASE_DIR=%LOCALAPPDATA%\Programs\Python\Python%%v"
        if exist "!BASE_DIR!\python.exe" (
            set "PY_EXE=!BASE_DIR!\python.exe"
            set "PYW_EXE=!BASE_DIR!\pythonw.exe"
            echo Detectado Python %%v no perfil do usuario.
        )
    )
)

:: 2. BUSCA SECUNDARIA: Comandos globais (PATH)
if not defined PY_EXE (
    where python >nul 2>nul
    if %ERRORLEVEL% equ 0 (
        :: Se encontrou, pega o caminho completo para evitar erro 'pythonw' nao encontrado
        for /f "delims=" %%i in ('where python') do set "PY_EXE=%%i"
        
        where pythonw >nul 2>nul
        if %ERRORLEVEL% equ 0 (
            for /f "delims=" %%i in ('where pythonw') do set "PYW_EXE=%%i"
        ) else (
            set "PYW_EXE=!PY_EXE:python.exe=pythonw.exe!"
        )
        echo Detectado Python no PATH global.
    )
)

:: 3. BUSCA TERCIARIA: Program Files
if not defined PY_EXE (
    for %%v in (315 314 313 312 311 310) do (
        if not defined PY_EXE (
            set "BASE_DIR=C:\Program Files\Python%%v"
            if exist "!BASE_DIR!\python.exe" (
                set "PY_EXE=!BASE_DIR!\python.exe"
                set "PYW_EXE=!BASE_DIR!\pythonw.exe"
                echo Detectado Python %%v em Program Files.
            )
        )
    )
)

if not defined PY_EXE (
    echo.
    echo [ERRO] Nao foi possivel encontrar o Python instalado.
    echo Por favor, instale o Python em https://www.python.org/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

:: Garantir que o pythonw existe de fato
if not exist "!PYW_EXE!" (
    echo [AVISO] 'pythonw.exe' nao encontrado. Usando 'python.exe' padrao.
    set "PYW_EXE=!PY_EXE!"
)

echo Usando Executavel: "!PY_EXE!"
echo Verificando dependencias...
"!PY_EXE!" -m pip install -r requirements.txt --quiet

echo --- Iniciando Servidor (Modo Silencioso) ---
echo Abrindo navegador...

:: Inicia o servidor de forma invisivel usando pythonw com caminho completo
start "" /b "!PYW_EXE!" server.py

:: Abre o navegador
start http://localhost:5000

echo.
echo Tudo pronto! O servidor fechara automaticamente ao fechar o navegador.
timeout /t 5 >nul
exit
