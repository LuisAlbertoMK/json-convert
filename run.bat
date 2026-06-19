@echo off
REM run.bat - Ejecuta run.ps1 (analisis + limpieza del proyecto)
REM
REM Uso:
REM   run.bat                    analisis + limpieza
REM   run.bat --SkipTests        solo limpieza
REM   run.bat --SkipCleanup      solo analisis

setlocal enabledelayedexpansion

set "SCRIPT_DIR=%%~dp0"

set "PWSH=powershell"
where pwsh >nul 2>&1
if !errorlevel! equ 0 set "PWSH=pwsh"

set "PS_ARGS="
:parse
if "%%~1"=="" goto :exec
set "PS_ARGS=%%PS_ARGS%% -%%~1"
shift
goto :parse

:exec
echo ========================================
echo   run.ps1 - Analisis + Limpieza
echo ========================================
echo.

%%PWSH%% -ExecutionPolicy Bypass -NoProfile -File "%%SCRIPT_DIR%%scripts\run.ps1" %%PS_ARGS%%

if %%errorlevel%% neq 0 (
    echo.
    echo [ERROR] El script detecto problemas. Revisa el reporte arriba.
    pause
    exit /b %%errorlevel%%
)

echo.
echo Proceso completado. Presiona cualquier tecla para salir...
pause >nul
