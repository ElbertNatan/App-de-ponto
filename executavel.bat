@echo off
setlocal

REM Descobre o Python: usa 'py' (launcher oficial) se existir, senao 'python'.
where py >nul 2>&1 && (set "PY=py") || (set "PY=python")

REM Garante as dependencias na primeira execucao (ou apos formatar a maquina).
%PY% -c "import openpyxl, PIL" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias necessarias ^(openpyxl, Pillow^)...
    %PY% -m pip install --upgrade openpyxl Pillow
    if errorlevel 1 (
        echo.
        echo Nao foi possivel instalar as dependencias.
        echo Verifique se o Python esta instalado ^(https://www.python.org/downloads/^).
        pause
        exit /b 1
    )
)

REM Sobe o app sem janela de console.
start "" %PY%w "%~dp0src\folha_ponto.py"
exit /b 0
