import os
import json
import re
import sys

# Forzar UTF-8 para evitar errores con tildes o símbolos
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto_sucursal(texto):
    """Limpia formatos como '$1.250,50' o '$1,250.50' para que sean floats válidos"""
    if not texto: return 0.0
    
    # 1. Quitar símbolo de dólar, espacios y cualquier etiqueta HTML residual
    limpio = re.sub(r'<.*?>', '', texto) # Elimina HTML si lo hubiera
    limpio = limpio.replace('$', '').replace(' ', '').strip()
    
    if not limpio: return 0.0
    
    # 2. Manejo robusto de separadores:
    if ',' in limpio and '.' in limpio:
        if limpio.rfind(',') > limpio.rfind('.'): # Caso estándar latino: 1.250,50
            limpio = limpio.replace('.', '').replace(',', '.')
        else: # Caso anglo: 1,250.50
            limpio = limpio.replace(',', '')
    elif ',' in limpio:
        limpio = limpio.replace(',', '.')
    elif '.' in limpio:
        partes = limpio.split('.')
        if len(partes[-1]) != 2: # No parece decimal (ej: 1.250)
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
            if carpeta.upper() in MESES_ES:
                mes_actual = carpeta.upper()
                ruta_mes = os.path.join(ruta_raiz, carpeta)

                # 2. Buscar en cada sucursal
                if not os.path.isdir(ruta_mes): continue
                
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    if not os.path.isdir(p_suc): continue
                    
                    archivo_html = os.path.join(p_suc, "cobros_detalles.html")
                    
                    if os.path.exists(archivo_html):
                        with open(archivo_html, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        
                        # Expresión regular mejorada para capturar el contenido de la clase monto
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

        # 3. Ordenar: Las que más cobraron van primero
        lista_ranking.sort(key=lambda x: x["COBRADO"], reverse=True)

        # 4. Guardar archivo final
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