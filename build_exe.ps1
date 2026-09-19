# Genera un único .exe del dashboard con PyInstaller.
# Uso:
#   powershell -ExecutionPolicy Bypass -File build_exe.ps1
# El .exe queda en dist\DashboardAlfresco.exe
# Uso del .exe:
#   .\DashboardAlfresco.exe                                   # eliges carpeta en la app
#   .\DashboardAlfresco.exe "C:\Contratos\Excels"             # carpeta inicial por argumento
#   $env:DASHBOARD_DATA_DIR="C:\Contratos\Excels"; .\DashboardAlfresco.exe

$ErrorActionPreference = "Stop"

$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $dir

Write-Host "== 1/3 Verificando Python y dependencias =="
python --version
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

Write-Host "== 2/3 Cerrando .exe anterior y limpiando builds =="
# El .exe anterior suele seguir corriendo en segundo plano (--windowed no muestra
# consola) y Windows bloquea el archivo. Hay que cerrarlo antes de recompilar.
Get-Process -Name DashboardAlfresco -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue build, dist
Remove-Item -Force -ErrorAction SilentlyContinue DashboardAlfresco.spec
if (Test-Path -LiteralPath "dist\DashboardAlfresco.exe") {
  throw "dist\DashboardAlfresco.exe sigue bloqueado. Ciérralo desde el Administrador de tareas (DashboardAlfresco.exe) o reinicia el PC y vuelve a ejecutar este script."
}

Write-Host "== 3/3 Construyendo .exe único (tarda varios minutos, ~300-800 MB) =="
python -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name DashboardAlfresco `
  --add-data "app.py;." `
  --add-data ".streamlit;./.streamlit" `
  --collect-all streamlit `
  --collect-all plotly `
  --collect-all pandas `
  --collect-submodules openpyxl `
  lanzar_dashboard.py

Write-Host ""
Write-Host "Listo: dist\DashboardAlfresco.exe"
Write-Host "Llévalo a cualquier PC Windows junto con tus Excel (no necesita Python)."
