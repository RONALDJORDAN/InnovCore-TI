@echo off
title InnovTech - Inventario de TI
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0inventario.ps1"
exit /b %ERRORLEVEL%
