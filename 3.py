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
        print(">>> EXTRACTOR DE INCIDENCIAS (VERSION TABLA) <<<")
        print("="*50)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        datos_json = {}

        for carpeta_mes in os.listdir(ruta_raiz):
            mes_key = carpeta_mes.upper()
            if mes_key in MESES_ES:
                ruta_mes = os.path.join(ruta_raiz, carpeta_mes)
                
                if os.path.isdir(ruta_mes):
                    print(f"\n📂 Procesando mes: {mes_key}")

                    for suc in os.listdir(ruta_mes):
                        p_suc = os.path.join(ruta_mes, suc)
                        if os.path.isdir(p_suc):
                            archivo_fuente = os.path.join(p_suc, "solo_mes.html")
                            
                            if os.path.exists(archivo_fuente):
                                with open(archivo_fuente, "r", encoding="utf-8") as f:
                                    html = f.read()
                                    
                                    # Busca todas las cifras dentro de etiquetas <td>
                                    valores = re.findall(r"<td>(.*?)</td>", html, re.IGNORECASE)
                                    
                                    # Según tu estructura: [0]=Mes, [1]=Total Incidencias, [2]=Calificación
                                    if len(valores) >= 2:
                                        try:
                                            # Limpiamos el valor por si hay espacios o etiquetas extra
                                            valor_str = re.sub(r'<.*?>', '', valores[1]).strip()
                                            valor = int(valor_str)
                                            
                                            nombre_clave = f"{suc.strip()} ({mes_key})"
                                            datos_json[nombre_clave] = valor
                                            print(f"   ✅ {suc}: {valor}")
                                        except ValueError:
                                            continue

        # Ordenar de mayor a menor
        ranking_ordenado = dict(sorted(datos_json.items(), key=lambda x: x[1], reverse=True))

        # Guardar el JSON final
        ruta_salida = os.path.join(ruta_raiz, "incidencias_totales.json")
        with open(ruta_salida, "w", encoding="utf-8") as jf:
            json.dump(ranking_ordenado, jf, ensure_ascii=False, indent=4)

        print("\n" + "="*50)
        print(f"✅ ÉXITO: {len(ranking_ordenado)} registros encontrados.")
        print("="*50)

    except Exception as e:
        print(f"\n❌ ERROR: {e}")

if __name__ == "__main__":
    generar_json_desde_solo_mes()