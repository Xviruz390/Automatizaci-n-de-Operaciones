@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    start "Automatizacion de reportes" ".venv\Scripts\python.exe" "interfaz_reporte.py"
) else (
    start "Automatizacion de reportes" py -3 "interfaz_reporte.py"
)
endlocal
