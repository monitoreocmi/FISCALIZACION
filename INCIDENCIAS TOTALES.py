import os
import json
import re
import sys

# Forzar UTF-8 para evitar problemas con tildes
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def generar_json_desde_solo_mes():
    try:
        print("\n" + "="*50)
        print(">>> EXTRACTOR DE INCIDENCIAS (VIA SOLO_MES.HTML) <<<")
        print("="*50)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        datos_json = {}

        # 1. Buscar en las carpetas de los meses
        for carpeta_mes in os.listdir(ruta_raiz):
            mes_key = carpeta_mes.upper()
            if mes_key in MESES_ES:
                ruta_mes = os.path.join(ruta_raiz, carpeta_mes)
                print(f"\n📂 Procesando mes: {mes_key}")

                # 2. Buscar sucursales
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    # Apuntamos al nuevo archivo simplificado
                    archivo_fuente = os.path.join(p_suc, "solo_mes.html")
                    
                    if os.path.exists(archivo_fuente):
                        with open(archivo_fuente, "r", encoding="utf-8") as f:
                            html = f.read()
                            
                            # Extraer el número de incidencias (clase total-incidencias)
                            # Buscamos el número que está dentro de las etiquetas <td>
                            match = re.search(r"total-incidencias'>(.*?)</td>", html, re.IGNORECASE)
                            
                            if match:
                                valor = int(match.group(1).strip())
                                nombre_clave = f"{suc.strip()} ({mes_key})"
                                datos_json[nombre_clave] = valor
                                print(f"   ✅ {suc}: {valor}")
                    else:
                        # Si no existe solo_mes, intentamos buscar en el reporte viejo como respaldo
                        pass

        # 3. Ordenar de mayor a menor (Ranking)
        ranking_ordenado = dict(sorted(datos_json.items(), key=lambda x: x[1], reverse=True))

        # 4. Guardar el JSON final
        ruta_salida = os.path.join(ruta_raiz, "incidencias_totales.json")
        with open(ruta_salida, "w", encoding="utf-8") as jf:
            json.dump(ranking_ordenado, jf, ensure_ascii=False, indent=4)

        print("\n" + "="*50)
        print(f"✅ ÉXITO: Archivo creado en: {ruta_salida}")
        print(f"Total registros: {len(ranking_ordenado)}")
        print("="*50)
        input("Presione ENTER para salir...")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        input("Presione ENTER para cerrar...")

if __name__ == "__main__":
    generar_json_desde_solo_mes()