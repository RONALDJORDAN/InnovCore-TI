@echo off
setlocal EnableDelayedExpansion
title InnovCore TI - Gestao de Ativos & Inventario [v2.0]
chcp 65001 >nul

set "PENDRIVE_DIR=%~dp0"
set "CORE_DIR=%~dp0core"
set "SCRIPT_FILE=%CORE_DIR%\inventario.ps1"

cd /d "%PENDRIVE_DIR%"

if not exist "%SCRIPT_FILE%" (
    echo.
    echo ======================================================================
    echo  [!] ERRO: Motor principal nao encontrado!
    echo      Esperado em: %SCRIPT_FILE%
    echo ======================================================================
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { $env:INNOVCORE_USB_DIR = '%PENDRIVE_DIR%'; & '%SCRIPT_FILE%' }"

endlocal
exit /b %ERRORLEVEL%
