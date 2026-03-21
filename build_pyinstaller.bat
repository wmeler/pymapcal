@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%"

if "%PYTHON_BIN%"=="" set "PYTHON_BIN=.\venv\Scripts\python.exe"

"%PYTHON_BIN%" -m PyInstaller --noconfirm --clean pymapcal.spec
if errorlevel 1 exit /b %errorlevel%

echo.
echo Build finished:
echo   %ROOT_DIR%dist\pymapcal
