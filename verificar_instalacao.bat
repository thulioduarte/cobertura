@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python verificar_instalacao.py
pause
