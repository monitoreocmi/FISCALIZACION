import os
import sys
import subprocess
from datetime import datetime

# --- CONFIGURACIÓN ---
RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))

# Forzamos a Python a usar UTF-8 para evitar errores de codificación en Windows
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

# Lista de tus scripts de Fiscalización Luxor
SCRIPTS_PANEL = [
    "COLORES.py",
    "1.py",
    "2.py",
    "3.py",
    "4.py",
    "5.py",
    "6.py",
    "7.py"
]

def ejecutar_panel():
    print(f"\n--- INICIO DE ACTUALIZACIÓN PANEL LUXOR [{datetime.now().strftime('%H:%M:%S')}] ---")
    print(f"Directorio de trabajo: {RUTA_RAIZ}\n")
    
    for script in SCRIPTS_PANEL:
        ruta_script = os.path.join(RUTA_RAIZ, script)
        
        if os.path.exists(ruta_script):
            print(f">> Ejecutando: {script}...", flush=True)
            try:
                # CAMBIO CLAVE: Quitamos capture_output=True para evitar el bloqueo (Deadlock)
                # Al usar stdout=None, el script imprime directo en la pantalla y no llena el buffer.
                resultado = subprocess.run(
    [sys.executable, ruta_script], 
    stdout=None, 
    stderr=None,
    text=True,
    encoding='utf-8', 
    errors='replace',
    env=os.environ,
    cwd=RUTA_RAIZ  # <--- Esto asegura que el script busque sus archivos en la carpeta correcta
)
                
                if resultado.returncode == 0:
                    print(f"✅ {script} FINALIZADO CON ÉXITO\n")
                else:
                    print(f"❌ {script} TERMINÓ CON ERRORES (Código: {resultado.returncode})\n")
                    
            except Exception as e:
                print(f"💥 FALLO CRÍTICO AL LANZAR {script}: {e}\n")
        else:
            print(f"⚠️ SALTADO: '{script}' no se encuentra en la carpeta.\n")

    print(f"--- FIN DEL PROCESO [{datetime.now().strftime('%H:%M:%S')}] ---")
    print("El Panel Centralizado ya debería tener los datos actualizados.")

if __name__ == "__main__":
    ejecutar_panel()