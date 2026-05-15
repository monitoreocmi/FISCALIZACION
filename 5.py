import os
import json
import re
import sys

# =================================================================
# ID: SINCRONIZADOR DE RANKING Y GESTIÓN DE COBROS (LUXOR)
# FUNCIÓN: Consolidar totales de sucursales para el Ranking Global
# =================================================================

# Forzar UTF-8 para evitar errores de codificación en Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto_sucursal(texto):
    """Limpia etiquetas HTML y convierte strings de moneda a float"""
    if not texto: return 0.0
    # Elimina tags HTML, símbolos de dólar, espacios y comas de miles
    limpio = re.sub(r'<.*?>', '', texto)
    limpio = limpio.replace('$', '').replace(' ', '').replace(',', '').strip()
    try:
        return float(limpio)
    except:
        return 0.0

def ejecutar_sincronizacion():
    try:
        print("\n" + "="*60)
        print(">>> LUXOR: GENERANDO RANKING GLOBAL DE RECAUDACION <<<")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(__file__))
        lista_ranking = []

        # Escaneo de carpetas de meses en la raíz
        if not os.path.exists(ruta_raiz):
            print("Error: No se encuentra la ruta raiz.")
            return

        for carpeta in os.listdir(ruta_raiz):
            nombre_carpeta_up = carpeta.upper()
            mes_detectado = next((m for m in MESES_ES if m in nombre_carpeta_up), None)
            
            ruta_mes = os.path.join(ruta_raiz, carpeta)
            if mes_detectado and os.path.isdir(ruta_mes):
                
                # Escaneo de cada sucursal dentro del mes
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    if not os.path.isdir(p_suc): continue
                    
                    # Buscamos el reporte de detalles generado previamente
                    archivo_html = os.path.join(p_suc, "cobros_detalles.html")
                    
                    if os.path.exists(archivo_html):
                        with open(archivo_html, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        
                        # Extraemos los montos usando la clase 'monto' definida en tus HTMLs
                        matches = re.findall(r"class='monto'>(.*?)</div>", html_content)
                        
                        if matches:
                            val_c = limpiar_monto_sucursal(matches[0]) if len(matches) >= 1 else 0.0
                            val_r = limpiar_monto_sucursal(matches[1]) if len(matches) >= 2 else 0.0
                            val_e = limpiar_monto_sucursal(matches[2]) if len(matches) >= 3 else 0.0
                            
                            lista_ranking.append({
                                "sucursal": f"{suc.strip()}",
                                "mes": mes_detectado,
                                "COBRADO": val_c,
                                "PERDIDA_PATRIMONIO": val_r,
                                "EXCEDENTE": val_e,
                                "TOTAL_GESTIONADO": round(val_c + val_r + val_e, 2)
                            })
                            print(f"OK: Sincronizado {suc} ({mes_detectado}) -> ${val_c:,.2f}")

        # Ordenar el ranking de mayor a menor según lo COBRADO
        lista_ranking.sort(key=lambda x: x["COBRADO"], reverse=True)
        
        # Guardar el JSON consolidado para el Dashboard
        ruta_final = os.path.join(ruta_raiz, "TOTALES_SUCURSALES_COBROS.json")
        with open(ruta_final, "w", encoding="utf-8") as jf:
            json.dump(lista_ranking, jf, indent=4, ensure_ascii=False)

        print("\n" + "="*60)
        print(f"EXITO: {len(lista_ranking)} registros consolidados en el Ranking.")
        print("="*60)

    except Exception as e:
        print(f"\nERROR EN SINCRONIZADOR: {e}")

if __name__ == "__main__":
    ejecutar_sincronizacion()