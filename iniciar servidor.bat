@echo off
:: Cambia al directorio donde están tus scripts
cd /d "C:\Users\Admin\Desktop\FISCALIZACION"

:: El comando 'start /min' abre la consola pero minimizada en la barra de tareas
:: Redirigimos la salida a logs por si ocurre un error (como el que viste antes)
start "Servidor Principal" /min python "SERVIDOR.py" > log_servidor.txt 2>&1
start "Servidor del Panel" /min python "servidor_panel.py" > log_panel.txt 2>&1

exit