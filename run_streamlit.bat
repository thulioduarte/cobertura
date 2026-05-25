@echo off
cd /d "%~dp0"
echo.
echo Iniciando Estudo de Cobertura Streamlit...
echo Pasta atual: %CD%
echo.
python -m streamlit run app.py
pause
