@echo off
:: Cambia esta ruta por la carpeta donde están tus scripts
cd /d "C:\Users\Admin\Desktop\FISCALIZACION"

echo === Iniciando Actualizacion de Datos Luxor ===

:: Ejecutar los 7 scripts en orden
echo Ejecutando script 1...
python 1.py
echo Ejecutando script 2...
python 2.py
echo Ejecutando script 3...
python 3.py
echo Ejecutando script 4...
python 4.py
echo Ejecutando script 5...
python 5.py
echo Ejecutando script 6...
python 6.py
echo Ejecutando script 7...
python 7.py

echo === Scripts finalizados. Subiendo a GitHub... ===

:: Comandos de Git para actualizar la web
git add .
git commit -m "Auto-update Luxor: %date% %time%"
git push origin main

echo === Proceso completado con exito ===
:: Opcional: timeout para ver si hubo errores antes de que se cierre la ventana
timeout /t 10