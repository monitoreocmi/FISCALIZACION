import os
import sys
import json
import re
import threading
import time

# Forzar UTF-8 para evitar problemas con tildes
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
RUTA_LOGO_PANEL = "RECURSOS/logo.png"

def sistema_luxor_v3():
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("        SISTEMA LUXOR: GENERADOR Y PANEL V3.0")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        
        # --- FASE 1: ESCANEO Y GENERACIÓN DE JSON ---
        # Estructura requerida para que el panel lea aprobadas/reprobadas
        dict_status = {"aprobadas": [], "reprobadas": []}
        dict_incidencias = {}
        
        print(f"🔍 Escaneando directorios en: {ruta_raiz}")
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
                                    # Incidencias (Columna 2)
                                    inc_raw = re.sub(r'<.*?>', '', valores[1]).strip()
                                    num_incidencias = int(inc_raw) if inc_raw.isdigit() else 0
                                    
                                    # Evaluación (Columna 3)
                                    eval_raw = re.sub(r'<.*?>', '', valores[2]).replace('%', '').strip()
                                    try: calificacion = float(eval_raw)
                                    except: calificacion = 0.0
                                    
                                    nombre_clave = f"{suc.strip()} ({mes_key})"
                                    dato_sucursal = {"n": nombre_clave, "v": int(calificacion)}

                                    # Lógica de estatus según la foto
                                    if calificacion >= 75:
                                        dict_status["aprobadas"].append(dato_sucursal)
                                    else:
                                        dict_status["reprobadas"].append(dato_sucursal)
                                    
                                    dict_incidencias[nombre_clave] = num_incidencias

        # Guardar JSONs fundamentales
        with open(os.path.join(ruta_raiz, "sucursales_status.json"), "w", encoding="utf-8") as f:
            json.dump(dict_status, f, ensure_ascii=False, indent=4)
        
        ranking_inc = dict(sorted(dict_incidencias.items(), key=lambda x: x[1], reverse=True))
        with open(os.path.join(ruta_raiz, "incidencias_totales.json"), "w", encoding="utf-8") as f:
            json.dump(ranking_inc, f, ensure_ascii=False, indent=4)
        
        print("✅ Archivos JSON actualizados.")

        # --- FASE 2: GENERACIÓN DEL PANEL (INDEX.HTML) ---
        def cargar_json(nombre):
            p = os.path.join(ruta_raiz, nombre)
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    try: return json.load(f)
                    except: return {}
            return {}

        data_totales = cargar_json("incidencias_totales.json")
        data_status = cargar_json("sucursales_status.json")
        data_graves = cargar_json("incidencias_graves.json")
        data_cobros_glob = cargar_json("TOTALES_GLOBALES_COBROS.json")
        data_suc_cobros_raw = cargar_json("TOTALES_SUCURSALES_COBROS.json")

        meses_disponibles = sorted([m for m in os.listdir(ruta_raiz) if m.upper() in MESES_ES and os.path.isdir(os.path.join(ruta_raiz, m))], 
                                   key=lambda x: MESES_ES.index(x.upper()))

        if not meses_disponibles:
            print("\n❌ Error: No se encontraron carpetas de meses.")
            return

        html_meses_data = ""
        opciones_dropdown = ""

        for m in meses_disponibles:
            m_key = m.upper()
            opciones_dropdown += f'<option value="mes-{m}" {"selected" if m == meses_disponibles[-1] else ""}>{m}</option>'
            ruta_mes = os.path.join(ruta_raiz, m)
            sucs_fisc = sorted([s for s in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, s)) and s.upper() != "CENTRAL"])

            # Filtrar datos por mes para las columnas del panel
            l_aprob = [i for i in data_status.get("aprobadas", []) if f"({m_key})" in i['n']]
            l_reprob = [i for i in data_status.get("reprobadas", []) if f"({m_key})" in i['n']]
            l_tot = sorted([{'n': k, 'v': v} for k, v in data_totales.items() if f"({m_key})" in k], key=lambda x: x['v'], reverse=True)
            l_grv = sorted([i for i in (data_graves if isinstance(data_graves, list) else []) if f"({m_key})" in i.get('n','')], key=lambda x: x.get('v',0), reverse=True)

            def limpiar_n(t): return str(t).split("(")[0].strip()

            def gen_rows(lista, css="", key_val="v", es_dinero=False):
                res = ""
                for i in lista:
                    name = limpiar_n(i['n'])
                    val = f"${i[key_val]:,.2f}" if es_dinero else i[key_val]
                    res += f"<div class='audit-row {css}'><span>{name}</span><b>{val}</b></div>"
                return res or '<div class="audit-row">Sin datos</div>'

            # Bloque de Cobros Globales por mes
            c_glob = data_cobros_glob.get(m_key, {})
            tc = c_glob.get('TOTAL_COBRADO', 0)
            tp = c_glob.get('TOTAL_PERDIDA_PATRIMONIO', 0)
            te = c_glob.get('TOTAL_EXCEDENTE', 0)

            html_meses_data += f"""
            <div id="mes-{m}" class="mes-container">
                <div id="incs-{m}" class="tab-content active">
                    <h2 class="sub-title">REPORTES FISCALIZACIÓN - {m_key}</h2>
                    <div class="blue-box"><div class="grid">{''.join([f'<a href="{m}/{s}/reporte.html" class="card">{s}</a>' for s in sucs_fisc])}</div></div>
                    <div class="audit-grid-full">
                        <div class="audit-card"><h3>INCIDENCIAS TOTALES</h3><div class="scroll-area">{gen_rows(l_tot, "row-blue")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES APROBADAS</h3><div class="scroll-area">{gen_rows(l_aprob, "status-ok")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES REPROBADAS</h3><div class="scroll-area">{gen_rows(l_reprob, "status-fail")}</div></div>
                        <div class="audit-card" style="border-top-color:var(--rojo)"><h3>INCIDENCIAS GRAVES</h3><div class="scroll-area">{gen_rows(l_grv, "status-fail")}</div></div>
                    </div>
                </div>
            </div>"""

        # Estilos y JS del Panel
        html_final = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>
            :root {{ --azul: #0844a4; --amarillo: #F9D908; --verde: #27ae60; --rojo: #ed1c24; --fondo: #f4f7f6; }}
            body {{ font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; }}
            .header {{ background: white; height: 80px; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; border-bottom: 4px solid var(--amarillo); }}
            .logo {{ height: 50px; }}
            h1 {{ color: var(--azul); font-size: 18px; text-transform: uppercase; font-weight: 900; }}
            .controls {{ display: flex; justify-content: center; gap: 10px; padding: 15px; background: #eee; border-bottom: 1px solid #ddd; }}
            .mes-container {{ display: none; padding: 20px; max-width: 1400px; margin: auto; }}
            .active {{ display: block !important; }}
            .sub-title {{ background: var(--azul); color: white; padding: 12px; border-radius: 6px; font-size: 14px; border-left: 6px solid var(--amarillo); margin-bottom: 15px; }}
            .audit-grid-full {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 15px; }}
            .audit-card {{ background: white; border-radius: 8px; padding: 15px; border-top: 4px solid var(--azul); box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
            .audit-row {{ display: flex; justify-content: space-between; padding: 8px; border-bottom: 1px solid #eee; font-size: 11px; font-weight: bold; }}
            .row-blue {{ background: #e3f2fd; color: #0d47a1; }}
            .status-ok {{ color: var(--verde); background: #e8f5e9; }}
            .status-fail {{ color: var(--rojo); background: #ffebee; }}
            .blue-box {{ background: white; padding: 15px; border-radius: 10px; border-top: 4px solid var(--azul); margin-bottom: 15px; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 10px; }}
            .card {{ background: var(--azul); color: white; padding: 15px 10px; text-decoration: none; text-align: center; border-radius: 6px; font-weight: 900; font-size: 12px; transition: 0.2s; }}
            .card:hover {{ background: var(--amarillo); color: var(--azul); transform: translateY(-2px); }}
            .scroll-area {{ max-height: 400px; overflow-y: auto; }}
            select {{ padding: 10px; border-radius: 5px; font-weight: bold; border: 2px solid var(--azul); }}
        </style></head><body>
            <header class="header"><img src="{RUTA_LOGO_PANEL}" class="logo"><h1>FISCALIZACIÓN LUXOR</h1><img src="{RUTA_LOGO_PANEL}" class="logo"></header>
            <div class="controls">
                <select id="mes-selector" onchange="cambiarMes()">{opciones_dropdown}</select>
            </div>
            <main>{html_meses_data}</main>
            <script>
                function cambiarMes() {{
                    document.querySelectorAll('.mes-container').forEach(e => e.classList.remove('active'));
                    let mes = document.getElementById('mes-selector').value;
                    if(document.getElementById(mes)) document.getElementById(mes).classList.add('active');
                }}
                window.onload = cambiarMes;
            </script>
        </body></html>"""

        with open(os.path.join(ruta_raiz, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_final)
        
        print("\n" + "="*50)
        print("✅ ÉXITO: JSON e Index.html generados correctamente.")
        print("="*50)

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO: {e}")

    print("\nPresiona ENTER para salir...")
    input()

if __name__ == "__main__":
    sistema_luxor_v3()