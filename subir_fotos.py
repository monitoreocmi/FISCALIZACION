import time
import requests
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- CONFIGURACIÓN ---
CARPETA_VIGILADA = r"C:\Users\Admin\Desktop\FISCALIZACION\fotos_incidencias"
URL_API_FLASK = "http://127.0.0.1:5000/api/subir_foto"

class ManejadorEventos(FileSystemEventHandler):
    def on_created(self, event):
        # Verificamos si lo que se creó es un archivo (la foto)
        if not event.is_directory:
            print(f"Nueva foto detectada en: {event.src_path}")
            
            # Pequeña espera para asegurar que iVMS terminó de escribir la foto
            time.sleep(1.5)
            
            try:
                with open(event.src_path, 'rb') as f:
                    # Enviamos el archivo al servidor Flask
                    files = {'foto': f}
                    response = requests.post(URL_API_FLASK, files=files)
                    
                    if response.status_code == 200:
                        print(">>> Foto enviada correctamente al servidor.")
                    else:
                        print(f">>> Error al enviar: {response.text}")
            except Exception as e:
                print(f">>> No se pudo procesar el archivo: {e}")

if __name__ == "__main__":
    print(f"Vigilando carpetas y subcarpetas en: {CARPETA_VIGILADA}")
    event_handler = ManejadorEventos()
    observer = Observer()
    
    # RECURSIVE=TRUE es la clave para que vigile las subcarpetas de fechas
    observer.schedule(event_handler, CARPETA_VIGILADA, recursive=True)
    
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()