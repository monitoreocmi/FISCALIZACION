import os
import json
import re
import sys

# Forzar UTF-8 para evitar errores con tildes o símbolos
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto_sucursal(texto):
    if not texto: return 0.0
    limpio = re.sub(r'<.*?>', '', texto)
    limpio = limpio.replace('$', '').replace(' ', '').strip()
    if not limpio: return 0.0
    
    if ',' in limpio and '.' in limpio:
        if limpio.rfind(',') > limpio.rfind('.'):
            limpio = limpio.replace('.', '').replace(',', '.')
        else:
            limpio = limpio.replace(',', '')
    elif ',' in limpio:
        limpio = limpio.replace(',', '.')
    elif '.' in limpio:
        partes = limpio.split('.')
        if len(partes[-1]) != 2:
             limpio = limpio.replace('.', '')
    try:
        return float(limpio)
    except:
        return 0.0

def generar_json_sucursales_ranking():
    try:
        print("\n" + "="*50)
        print(">>> GENERANDO: TOTALES_SUCURSALES_COBROS.json <<<")
        print("="*50)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        lista_ranking = []

        # 1. Buscar en carpetas de meses
        for carpeta in os.listdir(ruta_raiz):
            # --- CAMBIO CLAVE AQUÍ ---
            # Buscamos si algún mes de la lista está DENTRO del nombre de la carpeta
            nombre_carpeta_up = carpeta.upper()
            mes_detectado = next((m for m in MESES_ES if m in nombre_carpeta_up), None)
            
            if mes_detectado:
                mes_actual = mes_detectado
                ruta_mes = os.path.join(ruta_raiz, carpeta)

                if not os.path.isdir(ruta_mes): continue
                
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    if not os.path.isdir(p_suc): continue
                    
                    archivo_html = os.path.join(p_suc, "cobros_detalles.html")
                    
                    if os.path.exists(archivo_html):
                        with open(archivo_html, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        
                        matches = re.findall(r"class='monto'>(.*?)</div>", html_content)
                        
                        if len(matches) >= 2:
                            val_c = limpiar_monto_sucursal(matches[0])
                            val_r = limpiar_monto_sucursal(matches[1])
                            
                            lista_ranking.append({
                                "sucursal": f"{suc.strip()} ({mes_actual})",
                                "COBRADO": val_c,
                                "PERDIDA_PATRIMONIO": val_r,
                                "TOTAL_GESTIONADO": round(val_c + val_r, 2)
                            })
                            print(f"   📊 Procesado: {suc} ({mes_actual}) -> Cobrado: {val_c}")

        # Ordenar y Guardar
        lista_ranking.sort(key=lambda x: x["COBRADO"], reverse=True)
        ruta_final = os.path.join(ruta_raiz, "TOTALES_SUCURSALES_COBROS.json")
        with open(ruta_final, "w", encoding="utf-8") as jf:
            json.dump(lista_ranking, jf, indent=4, ensure_ascii=False)

        print("\n" + "="*50)
        print(f"✅ ÉXITO: Archivo generado con {len(lista_ranking)} sucursales.")
        print("="*50)

    except Exception as e:
        print(f"\n❌ ERROR: {e}")

if __name__ == "__main__":
    generar_json_sucursales_ranking()