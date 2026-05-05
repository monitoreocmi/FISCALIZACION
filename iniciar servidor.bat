@echo off
title Servidor Luxor - Sistema de Fiscalizacion
color 0A

:: 1. Verificacion de librerias necesarias
echo [1/3] Verificando componentes...
pip install flask pandas openpyxl --quiet

:: 2. Limpieza de consola para orden
cls
echo ======================================================
echo           SISTEMA DE FISCALIZACION LUXOR 2026
echo ======================================================
echo.
echo  [*] Iniciando servidor local...
echo  [*] El panel estara disponible en: http://127.0.0.1:5000
echo.
echo  [!] IMPORTANTE: No cierres esta ventana mientras 
echo      estes usando el sistema.
echo.
echo ======================================================
echo.

:: 3. Ejecucion del script de Python
:: Cambia "SERVIDOR.py" por el nombre real de tu archivo .py
python "SERVIDOR.py"

:: Si el script falla o se cierra, la ventana no se cerrara sola
if %errorlevel% neq 0 (
    echo.
    echo [!] EL SERVIDOR SE HA DETENIDO POR UN ERROR.
    pause
)