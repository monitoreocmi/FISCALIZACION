import os
import sys
import json
import threading
import time

# Forzar UTF-8 para evitar errores de caracteres
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
RUTA_LOGO_PANEL = "RECURSOS/logo.png" 

def generar_panel_luxor_centralizado():
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("        SISTEMA LUXOR - PANEL CENTRALIZADO V3.0")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        print(f"📂 Directorio: {ruta_raiz}")

        def cargar_json(nombre):
            p = os.path.join(ruta_raiz, nombre)
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                        print(f"   ✅ {nombre} cargado.")
                        return data
                    except:
                        print(f"   ⚠️ Error en formato de {nombre}.")
                        return {} if "totales" in nombre or "status" in nombre else []
            print(f"   ❌ {nombre} no encontrado.")
            return {} if "totales" in nombre or "status" in nombre else []

        # Carga de datos
        data_totales = cargar_json("incidencias_totales.json")
        data_status = cargar_json("sucursales_status.json")
        data_graves = cargar_json("incidencias_graves.json")
        data_cobros_glob = cargar_json("TOTALES_GLOBALES_COBROS.json")
        data_suc_cobros_raw = cargar_json("TOTALES_SUCURSALES_COBROS.json")

        meses_carpetas = sorted([m for m in os.listdir(ruta_raiz) if m.upper() in MESES_ES and os.path.isdir(os.path.join(ruta_raiz, m))], 
                                key=lambda x: MESES_ES.index(x.upper()))

        if not meses_carpetas:
            print("\n❌ Error: No hay carpetas de meses.")
            return

        print(f"\n📊 Procesando {len(meses_carpetas)} meses...")
        html_meses_data = ""
        opciones_dropdown = ""

        # Lógica de Cobros
        cobros_db = []
        c_raw = data_suc_cobros_raw if isinstance(data_suc_cobros_raw, (dict, list)) else []
        if isinstance(c_raw, dict):
            for k, v in c_raw.items():
                if isinstance(v, dict):
                    cobros_db.append({"sucursal": k, "c": v.get("COBRADO", 0), "p": v.get("PERDIDA_PATRIMONIO", 0), "e": v.get("EXCEDENTE", 0)})
        else:
            for item in c_raw:
                if isinstance(item, dict):
                    cobros_db.append({"sucursal": item.get("sucursal", ""), "c": item.get("COBRADO", 0), "p": item.get("PERDIDA_PATRIMONIO", 0), "e": item.get("EXCEDENTE", 0)})

        for m in meses_carpetas:
            m_key = m.upper()
            print(f"   > Generando HTML para: {m_key}")
            opciones_dropdown += f'<option value="mes-{m}" {"selected" if m == meses_carpetas[-1] else ""}>{m}</option>'
            ruta_mes = os.path.join(ruta_raiz, m)
            sucursales_fisicas = sorted([s for s in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, s)) and s.upper() != "CENTRAL"])

            def limpiar(t): return str(t).split("(")[0].strip().upper()

            def filtrar_c(lista, es_dict=False):
                if es_dict and isinstance(lista, dict):
                    return {k: v for k, v in lista.items() if limpiar(k) != "CENTRAL"}
                if isinstance(lista, list):
                    return [i for i in lista if limpiar(i.get('n', '')) != "CENTRAL"]
                return {} if es_dict else []

            # Filtrado de Incidencias
            dict_filtrado = filtrar_c(data_totales, True)
            list_totales = sorted([{'n': limpiar(k), 'v': v} for k, v in dict_filtrado.items() if f"({m_key})" in str(k).upper()], key=lambda x: x['v'], reverse=True)
            list_graves = sorted([{'n': limpiar(i['n']), 'v': i['v']} for i in filtrar_c(data_graves) if isinstance(i, dict) and f"({m_key})" in str(i.get('n','')).upper()], key=lambda x: x['v'], reverse=True)
            
            st_aprob = data_status.get("aprobadas", []) if isinstance(data_status, dict) else []
            list_aprob = [{'n': limpiar(i['n']), 'v': i['v']} for i in filtrar_c(st_aprob) if isinstance(i, dict) and f"({m_key})" in str(i.get('n','')).upper()]
            st_reprob = data_status.get("reprobadas", []) if isinstance(data_status, dict) else []
            list_reprob = [{'n': limpiar(i['n']), 'v': i['v']} for i in filtrar_c(st_reprob) if isinstance(i, dict) and f"({m_key})" in str(i.get('n','')).upper()]

            # Datos Central
            c_key = f"CENTRAL ({m_key})"
            list_totales_c = [{'n': 'CENTRAL', 'v': data_totales.get(c_key, 0) if isinstance(data_totales, dict) else 0}]
            grv_c_val = next((g.get('v', 0) for g in data_graves if isinstance(g, dict) and str(g.get('n','')).upper() == c_key), 0) if isinstance(data_graves, list) else 0
            list_graves_c = [{'n': 'CENTRAL', 'v': grv_c_val}]

            n_aprobadas = [i['n'] for i in list_aprob]
            n_reprobadas = [i['n'] for i in list_reprob]
            rank_fisc, rank_cobs, rank_fisc_c, rank_cobs_c = [], [], [], []

            for s in sorted([x for x in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, x))]):
                s_key = f"{s.strip().upper()} ({m_key})"
                inc_val = data_totales.get(s_key, 0) if isinstance(data_totales, dict) else 0
                grv_val = next((g.get('v', 0) for g in data_graves if isinstance(g, dict) and str(g.get('n','')).upper() == s_key), 0) if isinstance(data_graves, list) else 0
                c_val, p_val, e_val = 0, 0, 0
                for cb in cobros_db:
                    if str(cb.get('sucursal','')).upper() == s_key: 
                        c_val, p_val, e_val = cb.get('c',0), cb.get('p',0), cb.get('e', 0)
                        break
                item_f, item_c = {'n': s, 'v': inc_val + (grv_val * 10)}, {'n': s, 'c': c_val, 'p': p_val, 'e': e_val, 't': c_val + p_val + e_val}
                if s.upper() == "CENTRAL":
                    rank_fisc_c.append(item_f); rank_cobs_c.append(item_c)
                else:
                    rank_fisc.append(item_f); rank_cobs.append(item_c)

            def gen_rows_simple(lista, css=""):
                html = ""
                for i in lista: html += f"<div class='audit-row {css}'><span>{i['n']}</span></div>"
                return html or '<div class="audit-row">Sin datos</div>'

            def gen_rows(lista, key="v", css="", money=False):
                html = ""
                for i in lista:
                    val = f"${i[key]:,.2f}" if money else i[key]
                    html += f"<div class='audit-row {css}'><span>{i['n']}</span><b>{val}</b></div>"
                return html or '<div class="audit-row">Sin datos</div>'

            c_glob = data_cobros_glob.get(m_key, {}) if isinstance(data_cobros_glob, dict) else {}
            tc, tp, te = c_glob.get('TOTAL_COBRADO', 0), c_glob.get('TOTAL_PERDIDA_PATRIMONIO', 0), c_glob.get('TOTAL_EXCEDENTE', 0)
            
            def crear_bloque(is_c):
                p, sucs = ("c-", ["CENTRAL"]) if is_c else ("", sucursales_fisicas)
                l_tot, l_grv = (list_totales_c, list_graves_c) if is_c else (list_totales, list_graves)
                r_cobs, r_fisc = (rank_cobs_c, rank_fisc_c) if is_c else (rank_cobs, rank_fisc)
                
                bloque_global = f'<div class="global-cobros-box" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;"><div class="global-item" style="background: white; padding: 12px; border-radius: 8px; text-align: center; border-bottom: 4px solid var(--verde);"><span>TOTAL COBRADO</span><br><b style="font-size: 18px; display: block;">${tc:,.2f}</b></div><div class="global-item" style="background: white; padding: 12px; border-radius: 8px; text-align: center; border-bottom: 4px solid var(--rojo);"><span>PÉRDIDA MITIGADA</span><br><b style="font-size: 18px; display: block;">${tp:,.2f}</b></div><div class="global-item" style="background: white; padding: 12px; border-radius: 8px; text-align: center; border-bottom: 4px solid var(--azul);"><span>TOTAL MENSUAL</span><br><b style="font-size: 18px; display: block;">${(tc + tp + te):,.2f}</b></div></div>'

                return f"""
                <div id="{p}mes-{m}" class="mes-container {'central-mode' if is_c else ''}">
                    <div id="{p}incs-{m}" class="tab-content active">
                        <h2 class="sub-title">REPORTES {"CENTRAL" if is_c else "INCIDENCIAS"} - {m_key}</h2>
                        <div class="blue-box"><div class="grid">{''.join([f'<a href="{m}/{s}/reporte.html" class="card card-inc">{s}</a>' for s in sucs])}</div></div>
                        <div class="audit-grid-full">
                            <div class="audit-card"><h3>INCIDENCIAS TOTALES</h3><div class="scroll-area">{gen_rows(l_tot, css="row-blue")}</div></div>
                            <div class="audit-card"><h3>SUCURSALES APROBADAS</h3><div class="scroll-area">{gen_rows(list_aprob if not is_c else [], "v", "status-ok")}</div></div>
                            <div class="audit-card"><h3>SUCURSALES REPROBADAS</h3><div class="scroll-area">{gen_rows(list_reprob if not is_c else [], "v", "status-fail")}</div></div>
                            <div class="audit-card" style="border-top-color:var(--rojo)"><h3>MAYOR INCIDENCIAS GRAVES</h3><div class="scroll-area">{gen_rows(l_grv, css="status-fail")}</div></div>
                        </div>
                    </div>
                    <div id="{p}cobs-{m}" class="tab-content">
                        <h2 class="sub-title">REPORTES COBROS {"CENTRAL" if is_c else ""} - {m_key}</h2>
                        <div class="blue-box"><div class="grid">{''.join([f'<a href="{m}/{s}/cobros_detalles.html" class="card">{s}</a>' for s in sucs])}</div></div>
                        {bloque_global}
                        <div class="audit-grid-full">
                            <div class="audit-card"><h3>CANTIDAD COBRADA</h3><div class="scroll-area">{gen_rows(sorted(r_cobs, key=lambda x: x['c'], reverse=True), "c", money=True)}</div></div>
                            <div class="audit-card"><h3>PÉRDIDA MITIGADA</h3><div class="scroll-area">{gen_rows(sorted(r_cobs, key=lambda x: x['p'], reverse=True), "p", money=True)}</div></div>
                            <div class="audit-card" style="border-top-color:var(--azul)"><h3>EXCEDENTES</h3><div class="scroll-area">{gen_rows(sorted(r_cobs, key=lambda x: x['e'], reverse=True), "e", money=True)}</div></div>
                            <div class="audit-card" style="border-top-color:var(--amarillo)"><h3>TOTAL GLOBAL</h3><div class="scroll-area">{gen_rows(sorted(r_cobs, key=lambda x: x['t'], reverse=True), "t", money=True)}</div></div>
                        </div>
                    </div>
                    <div id="{p}honor-{m}" class="tab-content">
                        <h2 class="sub-title" style="background:var(--verde); border-left-color:white;">EXCELENTE RENDIMIENTO - {m_key}</h2>
                        <div class="audit-grid-full grid-half">
                            <div class="audit-card podio-high"><h3>MEJOR FISCALIZACIÓN</h3>{gen_rows_simple(sorted([x for x in r_fisc if x['n'].upper() in n_aprobadas or is_c], key=lambda x: x['v'])[:3], "status-ok-green")}</div>
                            <div class="audit-card podio-high"><h3>MEJOR COBROS</h3>{gen_rows_simple(sorted([x for x in r_cobs if x['n'].upper() in n_aprobadas or is_c], key=lambda x: x['t'])[:3], "status-ok-green")}</div>
                        </div>
                    </div>
                    <div id="{p}peores-{m}" class="tab-content">
                        <h2 class="sub-title" style="background:var(--rojo); border-left-color:black;">RENDIMIENTO CRÍTICO - {m_key}</h2>
                        <div class="audit-grid-full grid-half">
                            <div class="audit-card podio-low"><h3>PEOR FISCALIZACIÓN</h3>{gen_rows_simple(sorted([x for x in r_fisc if x['n'].upper() in n_reprobadas or is_c], key=lambda x: x['v'], reverse=True)[:3], "status-fail-red")}</div>
                            <div class="audit-card podio-low"><h3>PEOR COBROS</h3>{gen_rows_simple(sorted([x for x in r_cobs if x['n'].upper() in n_reprobadas or is_c], key=lambda x: x['t'], reverse=True)[:3], "status-fail-red")}</div>
                        </div>
                    </div>
                </div>"""

            html_meses_data += crear_bloque(False) + crear_bloque(True)

        # HTML e Interfaz de Usuario
        html_final = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>
            :root {{ --azul: #0844a4; --amarillo: #F9D908; --verde: #27ae60; --rojo: #ed1c24; --fondo: #f4f7f6; }}
            body {{ font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; padding: 0; }}
            .header-container {{ background: white; height: 80px; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; border-bottom: 4px solid var(--amarillo); }}
            .logo-panel {{ height: 50px; max-width: 90px; object-fit: contain; cursor: pointer; }}
            h1 {{ color: var(--azul); text-transform: uppercase; font-weight: 900; font-size: 16px; text-align: center; flex-grow: 1; margin: 0 10px; }}
            .controls-bar {{ display: flex; justify-content: center; gap: 8px; margin: 15px 0; flex-wrap: wrap; padding: 0 10px; }}
            .selector-wrapper {{ background: var(--azul); padding: 8px 15px; border-radius: 5px; border: 2px solid var(--amarillo); }}
            #mes-selector {{ background: transparent; color: white; border: none; font-weight: 900; font-size: 14px; cursor: pointer; outline: none; }}
            .tab-btn {{ padding: 10px 15px; border: none; background: #ddd; color: #666; font-weight: 900; border-radius: 5px; cursor: pointer; font-size: 11px; }}
            .tab-btn.active {{ background: var(--azul); color: white; box-shadow: 0 3px 0 var(--amarillo); }}
            #btn-hon.active {{ background: var(--verde) !important; }}
            #btn-peo.active {{ background: var(--rojo) !important; }}
            .main-content {{ padding: 0 15px 20px; max-width: 1400px; margin: 0 auto; }}
            .mes-container, .tab-content {{ display: none; }}
            .active {{ display: block !important; }}
            .sub-title {{ background: var(--azul); color: white; padding: 10px; border-radius: 6px; font-size: 12px; border-left: 6px solid var(--amarillo); margin-bottom: 15px; }}
            .audit-grid-full {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 15px; }}
            .grid-half {{ grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); }}
            .audit-card {{ background: white; border-radius: 8px; padding: 12px; border-top: 4px solid var(--azul); box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
            .audit-row {{ display: flex; justify-content: space-between; padding: 8px; border-bottom: 1px solid #eee; font-size: 11px; font-weight: bold; border-left: 4px solid transparent; }}
            .row-blue {{ background: #e3f2fd; color: #0d47a1; border-left-color: var(--azul); }}
            .status-ok {{ color: var(--verde); background: #e8f5e9; border-left-color: var(--verde); }}
            .status-fail {{ color: var(--rojo); background: #ffebee; border-left-color: var(--rojo); }}
            .status-ok-green {{ color: white !important; background: var(--verde) !important; font-size: 14px; margin-bottom: 5px; }}
            .status-fail-red {{ color: white !important; background: var(--rojo) !important; font-size: 14px; margin-bottom: 5px; }}
            .blue-box {{ background: white; padding: 20px; border-radius: 10px; border-top: 4px solid var(--azul); margin-bottom: 15px; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }}
            .card {{ padding: 18px 10px; text-align: center; border-radius: 8px; text-decoration: none; font-weight: 900; font-size: 13px; color: white; background: var(--azul); transition: 0.2s; min-width: 140px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .card:hover {{ background: var(--amarillo); color: var(--azul); transform: translateY(-2px); }}
            .scroll-area {{ max-height: 350px; overflow-y: auto; }}
        </style></head><body>
            <header class="header-container"><img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()"><h1>FISCALIZACIÓN LUXOR</h1><img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()"></header>
            <div class="controls-bar">
                <div class="selector-wrapper"><select id="mes-selector" onchange="cambiarMes()">{opciones_dropdown}</select></div>
                <button class="tab-btn active" id="btn-inc" onclick="showGlobalTab('incs')">INCIDENCIAS</button>
                <button class="tab-btn" id="btn-cob" onclick="showGlobalTab('cobs')">COBROS Y MITIGACION</button>
                <button class="tab-btn" id="btn-hon" onclick="showGlobalTab('honor')">MEJOR RENDIMIENTO</button>
                <button class="tab-btn" id="btn-peo" onclick="showGlobalTab('peores')">RENDIMIENTO CRÍTICO</button>
            </div>
            <main class="main-content">{html_meses_data}</main>
            <script>
                let isCentral = (localStorage.getItem('luxor_isCentral') === 'true');
                
                function toggleCentral() {{ 
                    isCentral = !isCentral; 
                    localStorage.setItem('luxor_isCentral', isCentral);
                    document.querySelector('h1').innerText = isCentral ? "SISTEMA LUXOR - CENTRAL" : "FISCALIZACIÓN LUXOR"; 
                    cambiarMes(); 
                }}

                function cambiarMes() {{ 
                    document.querySelectorAll(".mes-container").forEach(e => e.classList.remove('active')); 
                    let mes = document.getElementById("mes-selector").value; 
                    localStorage.setItem('luxor_mesActivo', mes);
                    let id = (isCentral ? "c-" : "") + mes; 
                    if(document.getElementById(id)) document.getElementById(id).classList.add('active'); 
                    actualizarPestañaInterna(); 
                }}

                function showGlobalTab(t) {{ 
                    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active')); 
                    localStorage.setItem('luxor_tabActiva', t);
                    let btnId = 'btn-' + t.substring(0,3); 
                    if(document.getElementById(btnId)) document.getElementById(btnId).classList.add('active'); 
                    
                    let mesActivo = document.querySelector('.mes-container.active'); 
                    if(mesActivo) {{ 
                        mesActivo.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active')); 
                        let id = (isCentral ? "c-" : "") + t + "-" + document.getElementById("mes-selector").value.replace('mes-', ''); 
                        if(document.getElementById(id)) document.getElementById(id).classList.add('active'); 
                    }} 
                }}

                function actualizarPestañaInterna() {{ 
                    let t = localStorage.getItem('luxor_tabActiva') || 'incs';
                    showGlobalTab(t); 
                }}

                window.onload = function() {{ 
                    if(isCentral) {{ document.querySelector('h1').innerText = "SISTEMA LUXOR - CENTRAL"; }}
                    let mesGuardado = localStorage.getItem('luxor_mesActivo');
                    if(mesGuardado && document.getElementById("mes-selector").querySelector('option[value="'+mesGuardado+'"]')) {{
                        document.getElementById("mes-selector").value = mesGuardado;
                    }}
                    cambiarMes(); 
                }};
            </script></body></html>"""
        
        with open(os.path.join(ruta_raiz, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_final)
        print("\n✅ PANEL ACTUALIZADO CORRECTAMENTE.")

    except Exception as e: 
        print(f"\n❌ ERROR CRÍTICO: {e}")

    print("\nPresiona ENTER para salir o espera 10 segundos...")
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__": 
    generar_panel_luxor_centralizado()
