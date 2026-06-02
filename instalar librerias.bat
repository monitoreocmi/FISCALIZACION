@echo off
echo Creando entorno virtual 'venv'...
if not exist venv python -m venv venv

echo Activando entorno virtual...
call venv\Scripts\activate

echo Forzando reinstalacion de librerias...
python -m pip install --upgrade pip
python -m pip install --force-reinstall --no-cache-dir flask flask-cors flask-socketio pandas openpyxl requests beautifulsoup4 watchdog

echo.
echo ========================================================
echo Instalacion forzada completada.
echo Todas las librerias han sido reinstaladas desde cero.
echo ========================================================
pause