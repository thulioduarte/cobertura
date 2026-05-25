@echo off
cd /d "%~dp0"
echo.
echo Instalando dependencias do Estudo de Cobertura...
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo Instalacao finalizada. Agora execute run_streamlit.bat
pause
