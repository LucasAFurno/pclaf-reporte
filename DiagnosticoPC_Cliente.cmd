@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
mode con: cols=100 lines=30 >nul 2>&1
title Diagnostico PC PCLAF - Cliente
color 0C

net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cls
echo.
echo ============================================================
echo                     P C L A F
echo            DIAGNOSTICO TECNICO DE EQUIPOS
echo ============================================================
echo.
echo [ 10%% ] Preparando entorno...
ping 127.0.0.1 -n 2 >nul

if not exist "%~dp0DiagnosticoPC.ps1" (
    echo ERROR: No se encuentra el archivo DiagnosticoPC.ps1
    echo.
    pause
    exit /b 1
)

echo [ 25%% ] Lanzando modo cliente...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DiagnosticoPC.ps1" -Modo cliente -Tecnico "Lucas / PCLAF" -MesesMantenimiento 6

echo.
echo [100%% ] Diagnostico finalizado.
echo ============================================================
echo FIN DEL DIAGNOSTICO
echo ============================================================
pause
endlocal
