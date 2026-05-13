import os
import json
import re
import sys
import time
import threading

# =================================================================
# ID: SINCRONIZADOR DE RANKING Y GESTIÓN DE COBROS (LUXOR)
# FUNCIÓN: Consolidar totales de sucursales para el Ranking Global
# =================================================================

# Forzar UTF-8 para evitar errores de codificación en consola
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto_sucursal(texto):
    """Limpia etiquetas HTML y convierte strings de moneda a float"""
    if not texto: return 0.0
    limpio = re.sub(r'<.*?>', '', texto) # Elimina cualquier tag residual
    limpio = limpio.replace('$', '').replace(' ', '').replace(',', '').strip()
    try:
        return float(limpio)
    except:
        return 0.0

def generar_json_sucursales_ranking():
    try:
        print("\n" + "="*50)
        print(">>> SINCRONIZANDO RANKING CON DATOS DE GESTIÓN <<<")
        print("="*50)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        lista_ranking = []

        # Escaneo de carpetas de meses
        for carpeta in os.listdir(ruta_raiz):
            nombre_carpeta_up = carpeta.upper()
            mes_detectado = next((m for m in MESES_ES if m in nombre_carpeta_up), None)
            
            if mes_detectado:
                mes_actual = mes_detectado
                ruta_mes = os.path.join(ruta_raiz, carpeta)

                if not os.path.isdir(ruta_mes): continue
                
                # Escaneo de sucursales dentro del mes
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    if not os.path.isdir(p_suc): continue
                    
                    # El origen de datos es el archivo generado por el script de cobros
                    archivo_html = os.path.join(p_suc, "cobros_detalles.html")
                    
                    if os.path.exists(archivo_html):
                        with open(archivo_html, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        
                        # Captura de los 3 valores principales: Cobrado, Recuperado, Excedente
                        matches = re.findall(r"class='monto'>(.*?)</div>", html_content)
                        
                        if matches:
                            val_c = limpiar_monto_sucursal(matches[0]) if len(matches) >= 1 else 0.0
                            val_r = limpiar_monto_sucursal(matches[1]) if len(matches) >= 2 else 0.0
                            val_e = limpiar_monto_sucursal(matches[2]) if len(matches) >= 3 else 0.0
                            
                            lista_ranking.append({
                                "sucursal": f"{suc.strip()} ({mes_actual})",
                                "COBRADO": val_c,
                                "PERDIDA_PATRIMONIO": val_r,
                                "EXCEDENTE": val_e,
                                "TOTAL_GESTIONADO": round(val_c + val_r + val_e, 2)
                            })
                            print(f"   📊 Datos extraídos: {suc} -> ${val_c:,.2f}")

        # Ordenar el ranking de mayor a menor recaudación (COBRADO)
        lista_ranking.sort(key=lambda x: x["COBRADO"], reverse=True)
        
        # Guardar resultado para que el Dashboard lo consuma
        ruta_final = os.path.join(ruta_raiz, "TOTALES_SUCURSALES_COBROS.json")
        with open(ruta_final, "w", encoding="utf-8") as jf:
            json.dump(lista_ranking, jf, indent=4, ensure_ascii=False)

        print("\n" + "="*50)
        print(f"✅ ÉXITO: {len(lista_ranking)} sucursales sincronizadas al Ranking.")
        print("="*50)

    except Exception as e:
        print(f"\n❌ ERROR DE PROCESAMIENTO: {e}")

    # Sistema de salida controlada
    print("\n" + "-"*30)
    print("Presiona ENTER para salir (o espera 10 segundos)...")
    
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__":
    generar_json_sucursales_ranking()