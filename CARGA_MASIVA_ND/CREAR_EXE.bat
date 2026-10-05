@echo off
cd /d "%~dp0"
echo ================================================
echo Colegio de Notarios de Lima
echo Generador de Carga Masiva ND - Compilacion EXE
echo ================================================
echo.
python -m pip install --upgrade pyinstaller
python -m pip install --upgrade tkinterdnd2
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "CNL_Carga_Masiva_No_Domiciliados" ^
  --icon "cnl_nd.ico" ^
  --add-data "logo_cnl.png;." ^
  --add-data "cnl_nd.ico;." ^
  --collect-all tkinterdnd2 ^
  "generador_carga_masiva_nd.py"
echo.
echo Listo. El ejecutable esta en:
echo dist\CNL_Carga_Masiva_No_Domiciliados.exe
pause
