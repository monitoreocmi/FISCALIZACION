@echo off
setlocal enabledelayedexpansion
title Instalador de Dependencias - Sistema LUXOR

echo ============================================================
echo   INSTALADOR DE LIBRERIAS MULTI-METODO (PANDAS/FLASK)
echo ============================================================
echo.

set LIBRERIAS=flask pandas openpyxl jinja2 xlrd pyinstaller

:: --- NIVEL 1: Intento estándar con PIP ---
echo [NIVEL 1] Intentando instalacion estandar...
pip install %LIBRERIAS%
if %ERRORLEVEL% EQU 0 goto :EXITO

:: --- NIVEL 2: Intento a traves del modulo de Python ---
echo.
echo [NIVEL 2] El comando 'pip' fallo. Intentando via modulo 'python -m'...
python -m pip install %LIBRERIAS%
if %ERRORLEVEL% EQU 0 goto :EXITO

:: --- NIVEL 3: Intento con modo Usuario (por falta de permisos) ---
echo.
echo [NIVEL 3] Fallo de permisos detectado. Intentando instalacion en modo --user...
python -m pip install --user %LIBRERIAS%
if %ERRORLEVEL% EQU 0 goto :EXITO

:: --- NIVEL 4: Actualizacion forzada de PIP y reintento ---
echo.
echo [NIVEL 4] Reintentando tras actualizar herramientas base...
python -m pip install --upgrade pip setuptools wheel
python -m pip install %LIBRERIAS%
if %ERRORLEVEL% EQU 0 goto :EXITO

:FALLO
echo.
echo xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
echo   ERROR: No se pudieron instalar las librerias.
echo   Verifica tu conexion a internet o si Python esta instalado.
echo xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
pause
exit

:EXITO
echo.
echo ============================================================
echo   OK: ¡Todas las librerias se instalaron correctamente!
echo ============================================================
pause
exit