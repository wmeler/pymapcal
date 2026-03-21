@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%"

if "%PYTHON_BIN%"=="" set "PYTHON_BIN=.\venv\Scripts\python.exe"
if "%BUILD_SUPPORT_DIR%"=="" set "BUILD_SUPPORT_DIR=%ROOT_DIR%build_support"
if "%IMGKAP_SRC_DIR%"=="" set "IMGKAP_SRC_DIR=%USERPROFILE%\imgkap"
if "%IMGKAP_OUT%"=="" set "IMGKAP_OUT=%BUILD_SUPPORT_DIR%\imgkap.exe"

if not exist "%BUILD_SUPPORT_DIR%" mkdir "%BUILD_SUPPORT_DIR%"

if exist "%IMGKAP_SRC_DIR%\imgkap.c" (
  echo Building bundled imgkap from: %IMGKAP_SRC_DIR%
  gcc "%IMGKAP_SRC_DIR%\imgkap.c" -O3 -s -lm -lfreeimage -o "%IMGKAP_OUT%"
  if errorlevel 1 exit /b %errorlevel%
  if not "%FREEIMAGE_DLL%"=="" copy /Y "%FREEIMAGE_DLL%" "%BUILD_SUPPORT_DIR%\FreeImage.dll" >nul
)

"%PYTHON_BIN%" -m PyInstaller --noconfirm --clean pymapcal.spec
if errorlevel 1 exit /b %errorlevel%

echo.
echo Build finished:
echo   %ROOT_DIR%dist\pymapcal
