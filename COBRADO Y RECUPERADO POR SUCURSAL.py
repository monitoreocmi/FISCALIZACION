import os
import json
import re
import sys

# Forzar UTF-8 para evitar errores con tildes o símbolos
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto_sucursal(texto):
    """Limpia formatos como '$1.250,50' para que sean números reales"""
    if not texto: return 0.0
    # Quitar $, espacios y dejar solo números, puntos y comas
    limpio = re.sub(r'[^\d,.]', '', texto)
    if not limpio: return 0.0
    
    # Estandarizar decimales (coma -> punto)
    if ',' in limpio and '.' in limpio:
        if limpio.rfind(',') > limpio.rfind('.'): # Formato 1.234,56
            limpio = limpio.replace('.', '').replace(',', '.')
        else: # Formato 1,234.56
            limpio = limpio.replace(',', '')
    elif ',' in limpio:
        limpio = limpio.replace(',', '.')
        
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
            if carpeta.upper() in MESES_ES:
                mes_actual = carpeta.upper()
                ruta_mes = os.path.join(ruta_raiz, carpeta)

                # 2. Buscar en cada sucursal
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    # Archivo generado por tu script de Fiscalización v3.3
                    archivo_html = os.path.join(p_suc, "cobros_detalles.html")
                    
                    if os.path.exists(archivo_html):
                        with open(archivo_html, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        
                        # Extraer montos de las tarjetas <div class='monto'>
                        # matches[0] = Cobrado | matches[1] = Recuperado
                        matches = re.findall(r"<div class='monto'>(.*?)</div>", html_content)
                        
                        if len(matches) >= 2:
                            val_c = limpiar_monto_sucursal(matches[0])
                            val_r = limpiar_monto_sucursal(matches[1])
                            
                            # Formato de objeto que espera el Dashboard
                            lista_ranking.append({
                                "sucursal": f"{suc.strip()} ({mes_actual})",
                                "COBRADO": val_c,
                                "PERDIDA_PATRIMONIO": val_r,
                                "TOTAL_GESTIONADO": round(val_c + val_r, 2)
                            })
                            print(f"   📊 Procesado: {suc} ({mes_actual})")

        # 3. Ordenar: Las que más cobraron van primero
        lista_ranking.sort(key=lambda x: x["COBRADO"], reverse=True)

        # 4. Guardar con el nombre exactO que busca el HTML
        ruta_final = os.path.join(ruta_raiz, "TOTALES_SUCURSALES_COBROS.json")
        with open(ruta_final, "w", encoding="utf-8") as jf:
            json.dump(lista_ranking, jf, indent=4, ensure_ascii=False)

        print("\n" + "="*50)
        print(f"✅ ÉXITO: Archivo generado con {len(lista_ranking)} sucursales.")
        print(f"Ruta: {ruta_final}")
        print("="*50)
        input("Presiona ENTER para salir...")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        input("Presiona ENTER para cerrar...")

if __name__ == "__main__":
    generar_json_sucursales_ranking()