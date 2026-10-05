@echo off
cd /d "%~dp0"
python "generador_carga_masiva_nd.py"
if errorlevel 1 (
  echo.
  echo No se pudo iniciar la aplicacion. Verifique que Python 3 este instalado.
  pause
)
