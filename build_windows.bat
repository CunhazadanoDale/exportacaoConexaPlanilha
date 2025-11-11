@echo off
setlocal ENABLEDELAYEDEXPANSION
REM ===== Athena Office - Build Windows (PyInstaller) =====

REM 1) Ambiente virtual
python -m venv .venv
call .venv\Scripts\activate

REM 2) Dependências
python -m pip install --upgrade pip wheel setuptools
pip install -r requirements.txt
pip install pyinstaller

REM 3) Build
pyinstaller --name "AthenaDashboard" --onefile app_dashboard_fixed.py ^
  --hidden-import=customtkinter ^
  --hidden-import=PIL ^
  --hidden-import=PIL._tkinter_finder

echo.
echo Build concluido. Arquivo em: dist\AthenaDashboard.exe
pause
