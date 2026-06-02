import os
from PIL import Image

def convertir_a_jpeg():
    # Obtiene la ruta donde está guardado el script
    ruta_raiz = os.path.dirname(os.path.abspath(__file__))
    
    print(f"--- Iniciando conversión en: {ruta_raiz} ---")
    
    extensiones_soportadas = ('.png', '.webp', '.jfif', '.bmp', '.tiff')
    convertidos = 0
    errores = 0

    # os.walk recorre carpetas y subcarpetas
    for root, dirs, files in os.walk(ruta_raiz):
        for nombre_archivo in files:
            # Ignorar si ya es jpeg o jpg
            if nombre_archivo.lower().endswith(extensiones_soportadas):
                ruta_completa = os.path.join(root, nombre_archivo)
                nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
                nueva_ruta = os.path.join(root, f"{nombre_sin_ext}.jpeg")

                try:
                    # Abrir imagen y convertir a RGB (necesario para JPEG si hay transparencias)
                    with Image.open(ruta_completa) as img:
                        rgb_img = img.convert('RGB')
                        rgb_img.save(nueva_ruta, 'JPEG', quality=95)
                    
                    # Eliminar el archivo original para que solo quede el .jpeg
                    os.remove(ruta_completa)
                    
                    print(f"✅ Convertido: {nombre_archivo} -> {nombre_sin_ext}.jpeg")
                    convertidos += 1
                except Exception as e:
                    print(f"❌ Error convirtiendo {nombre_archivo}: {e}")
                    errores += 1

    print("\n" + "="*30)
    print(f"PROCESO FINALIZADO")
    print(f"Convertidos: {convertidos}")
    print(f"Errores: {errores}")
    print("="*30)

if __name__ == "__main__":
    convertir_a_jpeg()