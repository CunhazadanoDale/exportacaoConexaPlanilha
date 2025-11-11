#!/usr/bin/env bash
set -euo pipefail

# ===== Athena Office - Build Linux (PyInstaller) =====
# 1) Ambiente virtual
python3 -m venv .venv
source .venv/bin/activate

# 2) Dependências
python -m pip install --upgrade pip wheel setuptools
pip install -r requirements.txt
pip install pyinstaller

# 3) Build
pyinstaller --name "AthenaDashboard" --onefile app_dashboard_fixed.py \
  --hidden-import=customtkinter \
  --hidden-import=PIL \
  --hidden-import=PIL._tkinter_finder

echo "Build concluido. Binario em: dist/AthenaDashboard"
