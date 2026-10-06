@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Generando Conversor de Cobranzas ODE
echo ==========================================
echo.

if not exist ".venv_build\Scripts\python.exe" (
    echo Creando entorno limpio de compilacion...
    python -m venv .venv_build
    if errorlevel 1 goto error
)

echo Instalando dependencias necesarias...
".venv_build\Scripts\python.exe" -m pip install --disable-pip-version-check pandas==2.2.3 openpyxl==3.1.5 xlrd==2.0.1 xlwt==1.3.0 pyinstaller==6.21.0
if errorlevel 1 goto error

echo Limpiando compilaciones anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo Generando ejecutable...
".venv_build\Scripts\python.exe" -m PyInstaller --clean --noconfirm --onefile --windowed --name Conversor_ODE app.py
if errorlevel 1 goto error

echo.
echo ==========================================
echo EJECUTABLE GENERADO CORRECTAMENTE
echo Ubicacion: %CD%\dist\Conversor_ODE.exe
echo ==========================================
echo.
pause
exit /b 0

:error
echo.
echo ==========================================
echo ERROR: No se pudo generar el ejecutable
echo Revisa los mensajes mostrados arriba.
echo ==========================================
pause
exit /b 1
