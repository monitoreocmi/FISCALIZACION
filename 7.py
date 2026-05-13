import os
import sys
import json
import threading
import time

# =================================================================
# ID: PANEL CENTRALIZADO LUXOR V3.0 (EL INTEGRADOR)
# FUNCIÓN: Generar el index.html final con dashboard interactivo
# =================================================================

# Forzar UTF-8 para evitar errores de caracteres en Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
RUTA_LOGO_PANEL = "RECURSOS/logo.png" 

def generar_panel_luxor_centralizado():
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("         SISTEMA LUXOR - PANEL CENTRALIZADO V3.0")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        print(f"📂 Directorio: {ruta_raiz}")

        def cargar_json(nombre):
            p = os.path.join(ruta_raiz, nombre)
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                        print(f"   ✅ {nombre} cargado correctamente.")
                        return data
                    except:
                        print(f"   ⚠️ Error en formato de {nombre}. Se usará lista vacía.")
                        return [] if "sucursales" in nombre or "graves" in nombre else {}
            print(f"   ❌ {nombre} no encontrado en la ruta.")
            return [] if "sucursales" in nombre or "graves" in nombre else {}

        # Carga de todos los componentes del ecosistema
        data_totales = cargar_json("incidencias_totales.json")
        data_status = cargar_json("sucursales_status.json")
        data_graves = cargar_json("incidencias_graves.json")
        data_cobros_glob = cargar_json("TOTALES_GLOBALES_COBROS.json")
        data_suc_cobros_raw = cargar_json("TOTALES_SUCURSALES_COBROS.json")

        # Filtrado de carpetas por mes
        meses_carpetas = sorted([m for m in os.listdir(ruta_raiz) if m.upper() in MESES_ES and os.path.isdir(os.path.join(ruta_raiz, m))], 
                                key=lambda x: MESES_ES.index(x.upper()))

        if not meses_carpetas:
            print("\n❌ Error: No se encontraron carpetas de meses válidas."); return

        print(f"\n📊 Procesando datos para {len(meses_carpetas)} meses...")
        html_meses_data = ""
        opciones_dropdown = ""

        # Normalización de datos de cobros
        cobros_db = []
        c_raw = data_suc_cobros_raw if isinstance(data_suc_cobros_raw, (dict, list)) else []
        if isinstance(c_raw, dict):
            for k, v in c_raw.items():
                cobros_db.append({"sucursal": k, "c": v.get("COBRADO", 0), "p": v.get("PERDIDA_PATRIMONIO", 0)})
        else:
            for item in c_raw:
                cobros_db.append({"sucursal": item.get("sucursal", ""), "c": item.get("COBRADO", 0), "p": item.get("PERDIDA_PATRIMONIO", 0)})

        for m in meses_carpetas:
            m_key = m.upper()
            print(f"   > Generando vista para: {m_key}")
            opciones_dropdown += f'<option value="mes-{m}" {"selected" if m == meses_carpetas[-1] else ""}>{m}</option>'
            ruta_mes = os.path.join(ruta_raiz, m)
            sucursales_fisicas = sorted([s for s in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, s))])

            def limpiar(t): return str(t).split("(")[0].strip().upper()

            # Filtrado por mes para las listas de ranking
            list_totales = sorted([{'n': limpiar(k), 'v': v} for k, v in data_totales.items() if f"({m_key})" in str(k).upper()], key=lambda x: x['v'], reverse=True)
            list_graves = sorted([{'n': limpiar(i['n']), 'v': i['v']} for i in data_graves if f"({m_key})" in str(i['n']).upper()], key=lambda x: x['v'], reverse=True)
            list_aprob = [{'n': limpiar(i['n']), 'v': i['v']} for i in data_status.get("aprobadas", []) if f"({m_key})" in str(i['n']).upper()]
            list_reprob = [{'n': limpiar(i['n']), 'v': i['v']} for i in data_status.get("reprobadas", []) if f"({m_key})" in str(i['n']).upper()]

            # Cálculo de Rankings
            rank_fisc, rank_cobs = [], []
            for s in sucursales_fisicas:
                s_key = f"{s.strip().upper()} ({m_key})"
                inc_val = data_totales.get(s_key, 0)
                grv_val = next((g.get('v', 0) for g in data_graves if str(g.get('n','')).upper() == s_key), 0)
                c_val, p_val = 0, 0
                for cb in cobros_db:
                    if str(cb.get('sucursal','')).upper() == s_key: c_val, p_val = cb.get('c',0), cb.get('p',0); break
                rank_fisc.append({'n': s, 'v': inc_val + (grv_val * 10)}) # Peso extra a graves
                rank_cobs.append({'n': s, 'c': c_val, 'p': p_val, 't': c_val + p_val})

            # Selección de mejores/peores
            top_f = sorted([x for x in rank_fisc if x['n'].upper() in [i['n'] for i in list_aprob]], key=lambda x: x['v'])[:3]
            top_c = sorted([x for x in rank_cobs if x['n'].upper() in [i['n'] for i in list_aprob]], key=lambda x: x['t'])[:3]
            bad_f = sorted([x for x in rank_fisc if x['n'].upper() in [i['n'] for i in list_reprob]], key=lambda x: x['v'], reverse=True)[:3]
            bad_c = sorted([x for x in rank_cobs if x['n'].upper() in [i['n'] for i in list_reprob]], key=lambda x: x['t'], reverse=True)[:3]

            def gen_rows_simple(lista, css=""):
                html = ""
                for i in lista:
                    tipo = "sucursal-central" if "CENTRAL" in i['n'].upper() else "sucursal-comun"
                    html += f"<div class='audit-row {css} {tipo}'><span>{i['n']}</span></div>"
                return html or '<div class="audit-row">Sin datos</div>'

            def gen_rows(lista, key="v", css="", money=False):
                html = ""
                for i in lista:
                    val = f"${i[key]:,.2f}" if money else i[key]
                    tipo = "sucursal-central" if "CENTRAL" in i['n'].upper() else "sucursal-comun"
                    html += f"<div class='audit-row {css} {tipo}'><span>{i['n']}</span><b>{val}</b></div>"
                return html or '<div class="audit-row">Sin datos</div>'

            # Estética de montos globales
            c_glob = data_cobros_glob.get(m_key, {})
            tc, tp = c_glob.get('TOTAL_COBRADO', 0), c_glob.get('TOTAL_PERDIDA_PATRIMONIO', 0)
            color_label = c_glob.get("COLOR_COBRADO", "NEGRO")
            estilo_monto = f"color: {'var(--verde)' if color_label == 'VERDE' else 'var(--rojo)' if color_label == 'ROJO' else '#333'};"

            botones_suc = "".join([f'<a href="{m}/{s}/reporte.html" class="card card-inc {"sucursal-central" if s.upper()=="CENTRAL" else "sucursal-comun"}">{s}</a>' for s in sucursales_fisicas])
            botones_cob = "".join([f'<a href="{m}/{s}/cobros_detalles.html" class="card {"sucursal-central" if s.upper()=="CENTRAL" else "sucursal-comun"}">{s}</a>' for s in sucursales_fisicas])

            # Bloque HTML para cada mes
            html_meses_data += f"""
            <div id="mes-{m}" class="mes-container">
                <div id="incs-{m}" class="tab-content active">
                    <h2 class="sub-title">REPORTES DE INCIDENCIAS - {m_key}</h2>
                    <div class="blue-box"><div class="grid">{botones_suc}</div></div>
                    <div class="audit-grid-full">
                        <div class="audit-card"><h3>INCIDENCIAS TOTALES</h3><div class="scroll-area">{gen_rows(list_totales, css="row-blue")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES APROBADAS</h3><div class="scroll-area">{gen_rows(list_aprob, "v", "status-ok")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES REPROBADAS</h3><div class="scroll-area">{gen_rows(list_reprob, "v", "status-fail")}</div></div>
                        <div class="audit-card" style="border-top-color:var(--rojo)"><h3>MAYOR INCIDENCIAS GRAVES</h3><div class="scroll-area">{gen_rows(list_graves, css="status-fail")}</div></div>
                    </div>
                </div>
                <div id="cobs-{m}" class="tab-content">
                    <h2 class="sub-title">REPORTES DE COBROS Y RECUPERACIÓN - {m_key}</h2>
                    <div class="blue-box"><div class="grid">{botones_cob}</div></div>
                    <div class="global-cobros-box">
                        <div class="global-item"><span>TOTAL COBRADO</span><br><b style="{estilo_monto}">${tc:,.2f}</b></div>
                        <div class="global-item"><span>PÉRDIDA MITIGADA</span><br><b>${tp:,.2f}</b></div>
                        <div class="global-item"><span>TOTAL MENSUAL</span><br><b>${(tc + tp):,.2f}</b></div>
                    </div>
                    <div class="audit-grid-full">
                        <div class="audit-card"><h3>CANTIDAD COBRADA</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['c'], reverse=True), "c", money=True)}</div></div>
                        <div class="audit-card"><h3>PÉRDIDA MITIGADA</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['p'], reverse=True), "p", money=True)}</div></div>
                        <div class="audit-card"><h3>TOTAL GLOBAL</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['t'], reverse=True), "t", money=True)}</div></div>
                    </div>
                </div>
                <div id="honor-{m}" class="tab-content">
                    <h2 class="sub-title" style="background:var(--verde); border-left-color:white;">EXCELENTE RENDIMIENTO - {m_key}</h2>
                    <div class="audit-grid-full grid-half">
                        <div class="audit-card podio-high"><h3>MEJORES SUCURSALES EN FISCALIZACIÓN</h3>{gen_rows_simple(top_f, "status-ok-green")}</div>
                        <div class="audit-card podio-high"><h3>MEJORES SUCURSALES EN COBROS Y MITIGACION</h3>{gen_rows_simple(top_c, "status-ok-green")}</div>
                    </div>
                </div>
                <div id="peores-{m}" class="tab-content">
                    <h2 class="sub-title" style="background:var(--rojo); border-left-color:black;">RENDIMIENTO CRÍTICO - {m_key}</h2>
                    <div class="audit-grid-full grid-half">
                        <div class="audit-card podio-low"><h3>PEORES SUCURSALES EN FISCALIZACIÓN</h3>{gen_rows_simple(bad_f, "status-fail-red")}</div>
                        <div class="audit-card podio-low"><h3>PEORES SUCURSALES EN COBROS Y MITIGACION</h3>{gen_rows_simple(bad_c, "status-fail-red")}</div>
                    </div>
                </div>
            </div>"""

        # Construcción del HTML final con estilos CSS y JS embebido
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
            .main-content {{ padding: 0 15px 20px; max-width: 1400px; margin: 0 auto; }}
            .mes-container, .tab-content {{ display: none; }}
            .active {{ display: block !important; }}
            .sub-title {{ background: var(--azul); color: white; padding: 10px; border-radius: 6px; font-size: 12px; border-left: 6px solid var(--amarillo); margin-bottom: 15px; }}
            .audit-grid-full {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 15px; }}
            .audit-card {{ background: white; border-radius: 8px; padding: 12px; border-top: 4px solid var(--azul); box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
            .audit-row {{ display: flex; justify-content: space-between; padding: 8px; border-bottom: 1px solid #eee; font-size: 11px; font-weight: bold; border-left: 4px solid transparent; }}
            
            /* LÓGICA DE CENTRAL */
            .sucursal-central {{ display: none !important; }}
            .modo-central .sucursal-central {{ display: flex !important; }}
            a.sucursal-central.card {{ display: none !important; }}
            .modo-central a.sucursal-central.card {{ display: block !important; }}

            .row-blue {{ background: #e3f2fd; color: #0d47a1; border-left-color: var(--azul); }}
            .status-ok {{ color: var(--verde); background: #e8f5e9; border-left-color: var(--verde); }}
            .status-fail {{ color: var(--rojo); background: #ffebee; border-left-color: var(--rojo); }}
            .status-ok-green {{ color: white !important; background: var(--verde) !important; font-size: 14px; margin-bottom: 5px; }}
            .status-fail-red {{ color: white !important; background: var(--rojo) !important; font-size: 14px; margin-bottom: 5px; }}
            .blue-box {{ background: white; padding: 15px; border-radius: 10px; border-top: 4px solid var(--azul); margin-bottom: 15px; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 8px; }}
            .card {{ padding: 10px; text-align: center; border-radius: 6px; text-decoration: none; font-weight: 900; font-size: 10px; color: white; background: var(--azul); transition: 0.2s; }}
            .card:hover {{ background: var(--amarillo); color: var(--azul); }}
            .global-cobros-box {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px; }}
            .global-item {{ background: white; padding: 12px; border-radius: 8px; text-align: center; border-bottom: 4px solid var(--verde); }}
            .global-item b {{ font-size: 18px; display: block; }}
            .scroll-area {{ max-height: 350px; overflow-y: auto; }}
        </style></head><body>
            <header class="header-container">
                <img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()">
                <h1>FISCALIZACIÓN LUXOR</h1>
                <img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()">
            </header>
            <div class="controls-bar">
                <div class="selector-wrapper"><select id="mes-selector" onchange="cambiarMes()">{opciones_dropdown}</select></div>
                <button class="tab-btn active" id="btn-inc" onclick="showGlobalTab('incs')">INCIDENCIAS</button>
                <button class="tab-btn" id="btn-cob" onclick="showGlobalTab('cobs')">COBROS Y MITIGACION</button>
                <button class="tab-btn" id="btn-hon" onclick="showGlobalTab('honor')">MEJOR RENDIMIENTO</button>
                <button class="tab-btn" id="btn-peo" onclick="showGlobalTab('peores')">RENDIMIENTO CRÍTICO</button>
            </div>
            <main class="main-content" id="main-panel">{html_meses_data}</main>
            <script>
                function toggleCentral() {{
                    document.body.classList.toggle('modo-central');
                }}
                
                function cambiarMes() {{
                    document.querySelectorAll(".mes-container").forEach(e => e.classList.remove('active'));
                    let sel = document.getElementById("mes-selector").value;
                    if(sel) document.getElementById(sel).classList.add('active');
                    actualizarPestañaInterna();
                }}

                function showGlobalTab(t) {{
                    localStorage.setItem('pestañaActiva', t);
                    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                    let btnId = 'btn-' + t.substring(0,3);
                    let btn = document.getElementById(btnId);
                    if(btn) btn.classList.add('active');
                    
                    let mes = document.querySelector('.mes-container.active');
                    if(mes) {{
                        mes.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                        let targetId = t + "-" + mes.id.replace('mes-', '');
                        let target = document.getElementById(targetId);
                        if(target) target.classList.add('active');
                    }}
                }}

                function actualizarPestañaInterna() {{
                    let t = localStorage.getItem('pestañaActiva') || 'incs';
                    showGlobalTab(t);
                }}

                window.onload = function() {{
                    cambiarMes();
                }};
            </script></body></html>"""
        
        with open(os.path.join(ruta_raiz, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_final)
        print("\n✅ PANEL ACTUALIZADO CORRECTAMENTE.")
    except Exception as e: print(f"\n❌ ERROR CRÍTICO: {e}")

    # Temporizador de cierre
    print("\nPresiona ENTER para salir o espera 10 segundos...")
    timer = threading.Timer(10.0, lambda: os._exit(0)); timer.start()
    try: input()
    finally: timer.cancel()

if __name__ == "__main__": generar_panel_luxor_centralizado()