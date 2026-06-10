@echo off
REM install.bat — Instala dependencias para extract_browser.py + extract_aa.py
REM NO requiere permisos de administrador (usa --user para pip)

echo ========================================
echo Instalando dependencias...
echo ========================================

echo 1. Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no encontrado. Instalalo desde https://www.python.org/downloads/
    echo Marca "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)
echo    Python OK

echo 2. Instalando openpyxl + playwright (--user, no requiere admin)...
python -m pip install --user openpyxl playwright --quiet
if %errorlevel% neq 0 (
    echo ERROR: Fallo la instalacion de paquetes.
    pause
    exit /b 1
)
echo    Paquetes OK

echo 3. Instalando Chromium para Playwright (en carpeta de usuario)...
python -m playwright install chromium
if %errorlevel% neq 0 (
    echo ERROR: Fallo la instalacion de Chromium.
    pause
    exit /b 1
)
echo    Chromium OK

echo.
echo ========================================
echo Instalacion completa. Ya podes ejecutar:
echo   python extract_browser.py
echo   python extract_aa.py
echo ========================================
pause
