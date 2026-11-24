@echo off
REM Sobe o app no Windows. Na primeira vez cria o ambiente (.venv) e instala as dependências.
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Criando ambiente virtual...
    py -3.13 -m venv .venv 2>nul || py -3 -m venv .venv || python -m venv .venv
    if errorlevel 1 (
        echo Nao encontrei o Python. Instale o Python 3.11 a 3.14 em python.org e tente de novo.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)

".venv\Scripts\python.exe" -m streamlit run app.py --browser.gatherUsageStats false
pause
