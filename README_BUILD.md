# Empacotar `app_dashboard.py` como App Desktop (sem mudar o código)

## Windows
1) Tenha Python 3.9+ no PATH.
2) Coloque `app_dashboard.py`, `requirements.txt` e `build_windows.bat` na mesma pasta.
3) Dê duplo clique em `build_windows.bat`.
4) Saída: `dist/AthenaDashboard.exe`.

## Linux
```bash
chmod +x build_linux.sh
./build_linux.sh
```
Saída: `dist/AthenaDashboard`.

Notas:
- Mantém console (útil para logs/instalação em runtime). Se quiser ocultar, adicione `--windowed` ao comando PyInstaller.
- Hidden imports incluídos: `customtkinter`, `PIL`, `PIL._tkinter_finder`.
- Dependências: pandas, openpyxl, xlrd, customtkinter, pillow.
