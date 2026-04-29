import os
import json
import re
import sys

# Forzado de UTF-8 para evitar errores de caracteres
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def clasificar_estatus_unico():
    try:
        print(">>> GENERANDO: sucursales_status.json <<<")
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        
        datos_estatus = {
            "aprobadas": [],
            "reprobadas": []
        }

        # 1. Buscar en las carpetas de meses
        for carpeta in os.listdir(ruta_raiz):
            if carpeta.upper() in MESES_ES:
                mes_llave = carpeta.upper()
                ruta_mes = os.path.join(ruta_raiz, carpeta)

                # 2. Buscar en las sucursales de ese mes
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    fuente = os.path.join(p_suc, "solo_mes.html")
                    
                    if os.path.exists(fuente):
                        with open(fuente, "r", encoding="utf-8") as f:
                            html = f.read()
                            
                            # Extraer el porcentaje del archivo solo_mes.html
                            match = re.search(r"calificacion-final'>(.*?)%</td>", html, re.IGNORECASE)
                            
                            if match:
                                nota = int(match.group(1).strip())
                                nombre_clave = f"{suc.strip()} ({mes_llave})"
                                item = {"n": nombre_clave, "v": f"{nota}%", "num": nota}
                                
                                # Clasificación (Umbral 75%)
                                if nota >= 75:
                                    datos_estatus["aprobadas"].append(item)
                                else:
                                    datos_estatus["reprobadas"].append(item)

        # 3. Ordenar por nota de mayor a menor
        for lista in ["aprobadas", "reprobadas"]:
            datos_estatus[lista].sort(key=lambda x: x["num"], reverse=True)
            # Eliminar la clave temporal para dejar el JSON limpio
            for i in datos_estatus[lista]:
                i.pop("num", None)

        # 4. Guardar archivo final
        ruta_final = os.path.join(ruta_raiz, "sucursales_status.json")
        with open(ruta_final, "w", encoding="utf-8") as jf:
            json.dump(datos_estatus, jf, ensure_ascii=False, indent=4)

        print(f"\n✅ PROCESO COMPLETADO")
        print(f"Se clasificaron {len(datos_estatus['aprobadas'])} aprobadas y {len(datos_estatus['reprobadas'])} reprobadas.")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
    
    input("\nPresiona ENTER para salir...")

if __name__ == "__main__":
    clasificar_estatus_unico()