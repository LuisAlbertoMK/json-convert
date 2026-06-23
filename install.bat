@echo off
REM install.bat — Instala dependencias para extract_browser.py + extract_aa.py
REM NO requiere permisos de administrador (usa --user para pip)
REM
REM Uso:
REM   install.bat                         # instalación normal
REM   install.bat --proxy http://proxy:8080  # detrás de proxy corporativo
REM
REM Requisitos:
REM   - Python 3.11+  (verificar con: python --version)
REM   - Google Chrome (recomendado para URLs de produccion)

setlocal enabledelayedexpansion

set PROXY=

:parse_args
if "%~1"=="--proxy" (
    set PROXY=%~2
    shift
    shift
    goto :parse_args
)

echo ========================================
echo Instalando dependencias...
echo ========================================

echo.
echo 1. Verificando Python...

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado.
    echo.
    echo Solucion: Instala Python desde https://www.python.org/downloads/
    echo Marca "Add Python to PATH" durante la instalacion.
    echo Version requerida: 3.11 o superior.
    echo.
    pause
    exit /b 1
)

REM Check version >= 3.11
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PY_VER=%%i
for /f "tokens=1,2 delims=." %%a in ("%PY_VER%") do (
    set PY_MAJOR=%%a
    set PY_MINOR=%%b
)

if %PY_MAJOR% lss 3 (
    echo [ERROR] Python 3 requerido. Version encontrada: %PY_VER%
    pause
    exit /b 1
)
if %PY_MAJOR%==3 if %PY_MINOR% lss 11 (
    echo [ERROR] Python 3.11+ requerido. Version encontrada: %PY_VER%
    pause
    exit /b 1
)

echo    Python %PY_VER% OK

echo.
echo 2. Instalando openpyxl + playwright...

set PIP_ARGS=--user --quiet --no-warn-script-location
if not "%PROXY%"=="" (
    set PIP_ARGS=%PIP_ARGS% --proxy %PROXY%
    echo    Proxy: %PROXY%
)

python -m pip install %PIP_ARGS% openpyxl playwright
if %errorlevel% neq 0 (
    echo [ERROR] Fallo la instalacion de paquetes.
    echo.
    echo Posibles causas:
    echo   - Sin conexion a internet
    echo   - Proxy corporativo: usA install.bat --proxy http://proxy:puerto
    echo   - Restriccion de seguridad: intentA sin --user
    echo.
    echo Para diagnosticar, ejecuta:
    echo   python -m pip install --user openpyxl --verbose
    echo.
    pause
    exit /b 1
)
echo    Paquetes OK

echo.
echo 2b. Instalando el paquete json-convert como editable...
python -m pip install --user --quiet --no-warn-script-location -e .
if !errorlevel! neq 0 (
    echo [ERROR] Fallo la instalacion del paquete json-convert.
    echo.
    echo Posibles causas:
    echo   - pip desactualizado: python -m pip install --upgrade pip
    echo   - Permisos: intentA sin --user o como administrador
    echo.
    pause
    exit /b 1
)
echo    json-convert OK

echo.
echo 3. Instalando Chromium para Playwright...

set PLAYWRIGHT_BROWSERS_PATH=%LOCALAPPDATA%\ms-playwright
python -m playwright install chromium
if %errorlevel% neq 0 (
    echo [ERROR] Fallo la instalacion de Chromium.
    echo.
    echo Posibles causas:
    echo   - Sin conexion a internet
    echo   - Proxy corporativo
    echo   - Antivirus bloqueando la descarga
    echo.
    echo Para diagnosticar:
    echo   python -m playwright install chromium --verbose
    echo.
    pause
    exit /b 1
)
echo    Chromium OK

echo.
echo 4. Verificando Google Chrome (recomendado)...
where chrome.exe >nul 2>&1
if %errorlevel% equ 0 (
    echo    Chrome OK
) else (
    echo    [AVISO] No se encontro Google Chrome.
    echo.
    echo    Las URLs de PRODUCCION requieren Chrome real para
    echo    evitar el bloqueo de Akamai WAF. Sin Chrome solo
    echo    se podran auditar URLs de PREVIEW.
    echo.
    echo    Descarga: https://www.google.com/chrome/
)

echo.
echo ========================================
echo Instalacion completa.
echo ========================================
echo.
echo Ya podes ejecutar:
echo   python menu.py                  # menu interactivo
echo   python extract_browser.py --urls urls.json
echo   python extract_aa.py
echo.
pause
