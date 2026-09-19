@echo off
REM Lanzador portable sin .exe: abre el dashboard con la carpeta data/ local.
REM Uso: doble clic, o: lanzar_dashboard.bat "C:\Contratos\Excels"
setlocal
cd /d "%~dp0"
if not "%~1"=="" set "DASHBOARD_DATA_DIR=%~1"
if not exist data mkdir data
python -m pip install -q -r requirements.txt
python -m streamlit run app.py
pause
