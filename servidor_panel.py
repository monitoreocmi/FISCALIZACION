from flask import Flask, send_from_directory, render_template_string
import os

app = Flask(__name__)

# Configuración de rutas: Usamos la carpeta donde se encuentra este archivo
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@app.route('/')
def index():
    """Sirve el archivo index.html principal."""
    if os.path.exists(os.path.join(BASE_DIR, 'index.html')):
        return send_from_directory(BASE_DIR, 'index.html')
    else:
        return "Error: No se encontró el archivo index.html en la raíz del servidor.", 404

@app.route('/<path:filename>')
def serve_files(filename):
    """
    Sirve cualquier otro archivo (HTML, CSS, JS, Imágenes) de forma dinámica.
    Esto permite navegar a otros HTML del panel automáticamente.
    """
    return send_from_directory(BASE_DIR, filename)

if __name__ == '__main__':
    # Puerto 80 para acceso directo vía IP (ej: http://192.168.1.XX)
    # Cambia el puerto a 8080 o 5001 si el 80 está ocupado por otro servicio.
    print(f"Servidor del Panel iniciado en la ruta: {BASE_DIR}")
    app.run(host='0.0.0.0', port=80, debug=True)
