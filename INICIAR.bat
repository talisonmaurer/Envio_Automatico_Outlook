@echo off
cd /d "%~dp0"
py -3 -c "import win32com.client" 2>nul
if errorlevel 1 (
 echo Instalando pywin32...
 py -3 -m pip install pywin32
 if errorlevel 1 (
  echo Erro ao instalar pywin32.
  pause
  exit /b 1
 )
)
py -3 envio_empresas.py
if errorlevel 1 pause
