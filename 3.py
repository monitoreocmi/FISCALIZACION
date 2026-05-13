import os
import sys
import json
import re
import threading
import time

# =================================================================
# ID: PROCESADOR DE DATOS Y GENERADOR DE ÍNDICE MENSUAL (LUXOR)
# FUNCIÓN: Escaneo de HTMLs y creación de index LOCAL por mes
# =================================================================

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
RUTA_LOGO_PANEL = "RECURSOS/logo.png"

def sistema_luxor_v3():
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("        SISTEMA LUXOR: PROCESADOR MENSUAL V3.0")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        
        dict_status = {"aprobadas": [], "reprobadas": []}
        dict_incidencias = {}
        
        print(f"🔍 Analizando carpetas de meses en: {ruta_raiz}")
        carpetas_meses = [d for d in os.listdir(ruta_raiz) if d.upper() in MESES_ES]

        for carpeta_mes in carpetas_meses:
            mes_key = carpeta_mes.upper()
            ruta_mes = os.path.join(ruta_raiz, carpeta_mes)
            
            if os.path.isdir(ruta_mes):
                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    if os.path.isdir(p_suc):
                        archivo_fuente = os.path.join(p_suc, "solo_mes.html")
                        if os.path.exists(archivo_fuente):
                            with open(archivo_fuente, "r", encoding="utf-8") as f:
                                html = f.read()
                                valores = re.findall(r"<td[^>]*>(.*?)</td>", html, re.IGNORECASE | re.DOTALL)
                                
                                if len(valores) >= 3:
                                    inc_raw = re.sub(r'<.*?>', '', valores[1]).strip()
                                    num_incidencias = int(inc_raw) if inc_raw.isdigit() else 0
                                    
                                    eval_raw = re.sub(r'<.*?>', '', valores[2]).replace('%', '').strip()
                                    try: calificacion = float(eval_raw)
                                    except: calificacion = 0.0
                                    
                                    nombre_clave = f"{suc.strip()} ({mes_key})"
                                    dato_sucursal = {"n": nombre_clave, "v": int(calificacion)}

                                    if calificacion >= 75:
                                        dict_status["aprobadas"].append(dato_sucursal)
                                    else:
                                        dict_status["reprobadas"].append(dato_sucursal)
                                    
                                    dict_incidencias[nombre_clave] = num_incidencias

                sucs_fisc = sorted([s for s in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, s)) and s.upper() != "CENTRAL"])
                
                l_aprob = [i for i in dict_status["aprobadas"] if f"({mes_key})" in i['n']]
                l_reprob = [i for i in dict_status["reprobadas"] if f"({mes_key})" in i['n']]
                l_tot = sorted([{'n': k, 'v': v} for k, v in dict_incidencias.items() if f"({mes_key})" in k], key=lambda x: x['v'], reverse=True)

                def limpiar_n(t): return str(t).split("(")[0].strip()
                def gen_rows(lista, css=""):
                    res = ""
                    for i in lista:
                        name = limpiar_n(i['n'])
                        val = i.get('v', 0)
                        res += f"<div class='audit-row {css}'><span>{name}</span><b>{val}</b></div>"
                    return res or '<div class="audit-row">Sin registros</div>'

                html_local = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><style>
                    :root {{ --azul: #0844a4; --amarillo: #F9D908; --verde: #27ae60; --rojo: #ed1c24; --fondo: #f4f7f6; }}
                    body {{ font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 20px; }}
                    .header {{ background: white; padding: 15px; border-bottom: 4px solid var(--amarillo); display: flex; justify-content: space-between; align-items: center; border-radius: 8px; }}
                    h1 {{ color: var(--azul); margin: 0; font-size: 20px; }}
                    .grid-local {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 10px; margin: 20px 0; }}
                    .card {{ background: var(--azul); color: white; padding: 15px; text-decoration: none; text-align: center; border-radius: 6px; font-weight: bold; font-size: 13px; }}
                    .audit-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; }}
                    .audit-card {{ background: white; padding: 15px; border-radius: 8px; border-top: 4px solid var(--azul); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
                    .audit-row {{ display: flex; justify-content: space-between; padding: 8px; border-bottom: 1px solid #eee; font-size: 12px; }}
                    .status-ok {{ color: var(--verde); }} .status-fail {{ color: var(--rojo); }}
                </style></head><body>
                    <div class="header"><h1>RESUMEN MENSUAL: {mes_key}</h1><a href="../../index.html" style="color:var(--azul); font-weight:bold;">← Volver al Panel Global</a></div>
                    <div class="grid-local">{''.join([f'<a href="{s}/reporte.html" class="card">{s}</a>' for s in sucs_fisc])}</div>
                    <div class="audit-grid">
                        <div class="audit-card"><h3>INCIDENCIAS</h3>{gen_rows(l_tot)}</div>
                        <div class="audit-card"><h3>APROBADAS (>=75%)</h3>{gen_rows(l_aprob, "status-ok")}</div>
                        <div class="audit-card"><h3>REPROBADAS (<75%)</h3>{gen_rows(l_reprob, "status-fail")}</div>
                    </div>
                </body></html>"""

                with open(os.path.join(ruta_mes, "index.html"), "w", encoding="utf-8") as f:
                    f.write(html_local)
                print(f"    📂 Index local generado en: {carpeta_mes}/index.html")

        with open(os.path.join(ruta_raiz, "sucursales_status.json"), "w", encoding="utf-8") as f:
            json.dump(dict_status, f, ensure_ascii=False, indent=4)
        
        ranking_inc = dict(sorted(dict_incidencias.items(), key=lambda x: x[1], reverse=True))
        with open(os.path.join(ruta_raiz, "incidencias_totales.json"), "w", encoding="utf-8") as f:
            json.dump(ranking_inc, f, ensure_ascii=False, indent=4)
        
        print("\n" + "="*50)
        print("✅ PROCESO COMPLETADO")
        print("   - JSONs globales actualizados en raíz.")
        print("   - Los index.html se guardaron DENTRO de cada mes.")
        print("="*50)

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO: {e}")

    # Lógica de cierre automático en 10 segundos o por teclado
    print("\nPresiona ENTER para salir o espera 10 segundos...")
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__":
    sistema_luxor_v3()