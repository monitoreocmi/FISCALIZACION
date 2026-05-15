import os
import sys
import json

# =================================================================
# ID: PANEL CENTRALIZADO LUXOR V3.5 - ENLACES EN MONTOS TOTALES
# =================================================================

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
RUTA_LOGO_PANEL = "RECURSOS/logo.png" 

def generar_panel_luxor_centralizado():
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("         SISTEMA LUXOR - PANEL CENTRALIZADO V3.5")
        print("="*60)
        
        ruta_raiz = os.path.dirname(os.path.abspath(__file__))

        def cargar_json(nombre):
            p = os.path.join(ruta_raiz, nombre)
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                        print(f"   ✅ {nombre} cargado.")
                        return data
                    except:
                        print(f"   ⚠️ Error de formato en {nombre}.")
                        return [] if "sucursales" in nombre or "graves" in nombre else {}
            else:
                default = [] if "sucursales" in nombre or "graves" in nombre else {}
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(default, f)
                return default

        data_totales = cargar_json("incidencias_totales.json")
        data_status = cargar_json("sucursales_status.json")
        data_graves_raw = cargar_json("incidencias_graves.json")
        data_cobros_glob = cargar_json("TOTALES_GLOBALES_COBROS.json")
        data_suc_cobros_raw = cargar_json("TOTALES_SUCURSALES_COBROS.json")

        data_graves = []
        if isinstance(data_graves_raw, dict):
            for mes_k, items in data_graves_raw.items():
                if isinstance(items, list):
                    for item in items:
                        n = item.get('sucursal') or item.get('n') or "S/N"
                        v = item.get('total') or item.get('v') or 0
                        data_graves.append({'n': f"{n.upper()} ({mes_k.upper()})", 'v': v})
        else:
            for item in data_graves_raw:
                data_graves.append({'n': item.get('sucursal', item.get('n', 'S/N')), 'v': item.get('total', item.get('v', 0))})

        meses_carpetas = sorted([m for m in os.listdir(ruta_raiz) if m.upper() in MESES_ES and os.path.isdir(os.path.join(ruta_raiz, m))], 
                                key=lambda x: MESES_ES.index(x.upper()))

        if not meses_carpetas:
            print("\n❌ Error: No hay carpetas de meses."); input(); return

        html_meses_data = ""
        opciones_dropdown = ""
        
        cobros_db = []
        c_raw = data_suc_cobros_raw if isinstance(data_suc_cobros_raw, (dict, list)) else []
        if isinstance(c_raw, dict):
            for k, v in c_raw.items():
                cobros_db.append({"sucursal": k, "c": v.get("COBRADO", 0), "p": v.get("PERDIDA_PATRIMONIO", 0), "e": v.get("EXCEDENTE", 0)})
        else:
            for item in c_raw:
                cobros_db.append({"sucursal": item.get("sucursal", ""), "c": item.get("COBRADO", 0), "p": item.get("PERDIDA_PATRIMONIO", 0), "e": item.get("EXCEDENTE", 0)})

        for m in meses_carpetas:
            m_key = m.upper()
            opciones_dropdown += f'<option value="mes-{m}" {"selected" if m == meses_carpetas[-1] else ""}>{m}</option>'
            
            def limpiar(t): return str(t).split("(")[0].strip().upper()

            list_totales = sorted([{'n': limpiar(k), 'v': v} for k, v in data_totales.items() if f"({m_key})" in str(k).upper()], key=lambda x: x['v'], reverse=True)
            list_graves = sorted([{'n': limpiar(i['n']), 'v': i['v']} for i in data_graves if f"({m_key})" in str(i['n']).upper()], key=lambda x: x['v'], reverse=True)
            list_aprob = [{'n': limpiar(i.get('sucursal', i.get('n', ''))), 'v': i.get('total', i.get('v', 0))} for i in data_status.get("aprobadas", []) if f"({m_key})" in str(i.get('sucursal', i.get('n', ''))).upper()]
            list_reprob = [{'n': limpiar(i.get('sucursal', i.get('n', ''))), 'v': i.get('total', i.get('v', 0))} for i in data_status.get("reprobadas", []) if f"({m_key})" in str(i.get('sucursal', i.get('n', ''))).upper()]

            sucursales_mes = sorted(list(set([str(k).split("(")[0].strip() for k in data_totales.keys() if f"({m_key})" in str(k).upper()])))

            rank_fisc, rank_cobs = [], []
            total_excedentes_mes = 0
            for s in sucursales_mes:
                s_key = f"{s.upper()} ({m_key})"
                inc_val = data_totales.get(s_key, 0)
                grv_val = next((g.get('v', 0) for g in data_graves if str(g.get('n','')).upper() == s_key), 0)
                c_val, p_val, e_val = 0, 0, 0
                for cb in cobros_db:
                    if str(cb.get('sucursal','')).upper() == s_key: 
                        c_val, p_val, e_val = cb.get('c',0), cb.get('p',0), cb.get('e',0); break
                total_excedentes_mes += e_val
                rank_fisc.append({'n': s, 'v': inc_val + (grv_val * 10)})
                rank_cobs.append({'n': s, 'c': c_val, 'p': p_val, 't': c_val + p_val})

            top_f = sorted([x for x in rank_fisc if x['n'].upper() in [i['n'] for i in list_aprob]], key=lambda x: x['v'])[:3]
            top_c = sorted([x for x in rank_cobs if x['n'].upper() in [i['n'] for i in list_aprob]], key=lambda x: x['t'], reverse=True)[:3]
            bad_f = sorted([x for x in rank_fisc if x['n'].upper() in [i['n'] for i in list_reprob]], key=lambda x: x['v'], reverse=True)[:3]
            bad_c = sorted([x for x in rank_cobs if x['n'].upper() in [i['n'] for i in list_reprob]], key=lambda x: x['t'])[:3]

            def gen_rows(lista, key=None, css="", money=False):
                html = ""
                for i in lista:
                    tipo = "sucursal-central" if "CENTRAL" in i['n'].upper() else "sucursal-comun"
                    val_html = ""
                    if key and key in i:
                        val = f"${i[key]:,.2f}" if money else i[key]
                        val_html = f"<b>{val}</b>"
                    html += f"<div class='audit-row {css} {tipo}'><span>{i['n']}</span>{val_html}</div>"
                return html or '<div class="audit-row">Sin registros</div>'

            c_glob = data_cobros_glob.get(m_key, {})
            tc, tp = c_glob.get('TOTAL_COBRADO', 0), c_glob.get('TOTAL_PERDIDA_PATRIMONIO', 0)
            estilo_monto = f"color: {'var(--verde)' if c_glob.get('COLOR_COBRADO') == 'VERDE' else 'var(--rojo)' if c_glob.get('COLOR_COBRADO') == 'ROJO' else '#333'};"

            botones_suc = "".join([f'<a href="{m}/{s}/reporte.html" class="card {"sucursal-central" if s.upper()=="CENTRAL" else "sucursal-comun"}">{s}</a>' for s in sucursales_mes])
            botones_cob = "".join([f'<a href="{m}/{s}/cobros_detalles.html" class="card {"sucursal-central" if s.upper()=="CENTRAL" else "sucursal-comun"}">{s}</a>' for s in sucursales_mes])

            html_meses_data += f"""
            <div id="mes-{m}" class="mes-container">
                <div id="incs-{m}" class="tab-content active">
                    <h2 class="sub-title">REPORTES DE INCIDENCIAS - {m_key}</h2>
                    <div class="blue-box"><div class="grid">{botones_suc}</div></div>
                    <div class="audit-grid-full">
                        <div class="audit-card"><h3>INCIDENCIAS TOTALES</h3><div class="scroll-area">{gen_rows(list_totales, 'v', "row-blue")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES APROBADAS</h3><div class="scroll-area">{gen_rows(list_aprob, 'v', "status-ok")}</div></div>
                        <div class="audit-card"><h3>SUCURSALES REPROBADAS</h3><div class="scroll-area">{gen_rows(list_reprob, 'v', "status-fail")}</div></div>
                        <div class="audit-card" style="border-top-color:var(--rojo)"><h3>INCIDENCIAS GRAVES</h3><div class="scroll-area">{gen_rows(list_graves, 'v', "status-fail")}</div></div>
                    </div>
                </div>
                <div id="cobs-{m}" class="tab-content">
                    <h2 class="sub-title">RECUPERACIÓN Y COBROS - {m_key}</h2>
                    <div class="blue-box"><div class="grid">{botones_cob}</div></div>
                    
                    <div class="global-cobros-box">
                        <a href="{m}/todos_cobrados.html" class="global-item-link">
                            <div class="global-item"><span>TOTAL COBRADO</span><br><b style="{estilo_monto}">${tc:,.2f}</b></div>
                        </a>
                        <a href="{m}/todos_recuperados.html" class="global-item-link">
                            <div class="global-item"><span>PÉRDIDA MITIGADA</span><br><b>${tp:,.2f}</b></div>
                        </a>
                        <a href="{m}/todos_excedentes.html" class="global-item-link">
                            <div class="global-item" style="border-bottom-color: var(--amarillo)"><span>EXCEDENTES</span><br><b>${total_excedentes_mes:,.2f}</b></div>
                        </a>
                        <div class="global-item"><span>TOTAL MENSUAL</span><br><b>${(tc + tp):,.2f}</b></div>
                    </div>

                    <div class="audit-grid-full">
                        <div class="audit-card"><h3>POR COBRADO</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['c'], reverse=True), 'c', money=True)}</div></div>
                        <div class="audit-card"><h3>POR MITIGACIÓN</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['p'], reverse=True), 'p', money=True)}</div></div>
                        <div class="audit-card"><h3>TOTAL GLOBAL</h3><div class="scroll-area">{gen_rows(sorted(rank_cobs, key=lambda x: x['t'], reverse=True), 't', money=True)}</div></div>
                    </div>
                </div>
                <div id="honor-{m}" class="tab-content">
                    <h2 class="sub-title" style="background:var(--verde); border-left-color:white;">MEJOR RENDIMIENTO - {m_key}</h2>
                    <div class="audit-grid-full" style="grid-template-columns: 1fr 1fr;">
                        <div class="audit-card" style="border-top-color:var(--verde)"><h3>TOP FISCALIZACIÓN</h3>{gen_rows(top_f, css="status-ok-green")}</div>
                        <div class="audit-card" style="border-top-color:var(--verde)"><h3>TOP COBROS</h3>{gen_rows(top_c, css="status-ok-green")}</div>
                    </div>
                </div>
                <div id="peores-{m}" class="tab-content">
                    <h2 class="sub-title" style="background:var(--rojo); border-left-color:black;">RENDIMIENTO CRÍTICO - {m_key}</h2>
                    <div class="audit-grid-full" style="grid-template-columns: 1fr 1fr;">
                        <div class="audit-card" style="border-top-color:var(--rojo)"><h3>CRÍTICOS FISCALIZACIÓN</h3>{gen_rows(bad_f, css="status-fail-red")}</div>
                        <div class="audit-card" style="border-top-color:var(--rojo)"><h3>CRÍTICOS COBROS</h3>{gen_rows(bad_c, css="status-fail-red")}</div>
                    </div>
                </div>
            </div>"""

        html_final = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><style>
            :root {{ --azul: #0844a4; --amarillo: #F9D908; --verde: #27ae60; --rojo: #ed1c24; --fondo: #f4f7f6; }}
            body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: var(--fondo); margin: 0; padding-bottom: 50px; }}
            .header-container {{ background: white; height: 80px; display: flex; align-items: center; justify-content: space-between; padding: 0 30px; border-bottom: 5px solid var(--amarillo); }}
            .logo-panel {{ height: 55px; cursor: pointer; }}
            h1 {{ color: var(--azul); font-weight: 900; font-size: 20px; }}
            .controls-bar {{ display: flex; justify-content: center; gap: 10px; margin: 20px 0; flex-wrap: wrap; }}
            #mes-selector {{ background: var(--azul); color: white; padding: 10px; border-radius: 5px; font-weight: bold; border: 2px solid var(--amarillo); }}
            .tab-btn {{ padding: 10px 20px; border: none; background: #ddd; font-weight: 900; border-radius: 5px; cursor: pointer; transition: 0.3s; }}
            .tab-btn.active {{ background: var(--azul); color: white; box-shadow: 0 4px 0 var(--amarillo); }}
            .main-content {{ padding: 0 20px; max-width: 1400px; margin: auto; }}
            .mes-container, .tab-content {{ display: none; }}
            .active {{ display: block !important; }}
            .sub-title {{ background: var(--azul); color: white; padding: 12px; border-radius: 6px; font-size: 14px; border-left: 50px solid var(--amarillo); margin-bottom: 15px; }}
            .audit-grid-full {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 15px; }}
            .audit-card {{ background: white; border-radius: 8px; padding: 15px; border-top: 5px solid var(--azul); box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .audit-row {{ display: flex; justify-content: space-between; padding: 10px; border-bottom: 1px solid #eee; font-size: 12px; font-weight: bold; }}
            .row-blue {{ background: #e3f2fd; color: #0d47a1; }}
            .status-ok {{ color: var(--verde); background: #e8f5e9; }}
            .status-fail {{ color: var(--rojo); background: #ffebee; }}
            .status-ok-green {{ background: var(--verde) !important; color: white !important; margin-bottom: 5px; border-radius: 4px; }}
            .status-fail-red {{ background: var(--rojo) !important; color: white !important; margin-bottom: 5px; border-radius: 4px; }}
            .blue-box {{ background: white; padding: 20px; border-radius: 10px; border-top: 5px solid var(--azul); margin-bottom: 15px; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(130px, 1fr)); gap: 10px; }}
            .card {{ padding: 12px; text-align: center; border-radius: 6px; text-decoration: none; font-weight: 900; font-size: 11px; color: white; background: var(--azul); }}
            .card:hover {{ background: var(--amarillo); color: var(--azul); }}
            .global-cobros-box {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }}
            .global-item {{ background: white; padding: 15px; border-radius: 8px; text-align: center; border-bottom: 5px solid var(--verde); transition: 0.3s; }}
            .global-item:hover {{ transform: translateY(-5px); box-shadow: 0 5px 15px rgba(0,0,0,0.1); }}
            .global-item b {{ font-size: 22px; display: block; margin-top: 5px; }}
            .global-item-link {{ text-decoration: none; color: inherit; }}
            .scroll-area {{ max-height: 400px; overflow-y: auto; }}
            .sucursal-central {{ display: none !important; }}
            .modo-central .sucursal-central {{ display: flex !important; }}
        </style></head><body>
            <header class="header-container">
                <img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()">
                <h1>FISCALIZACIÓN GRUPO LUXOR</h1>
                <img src="{RUTA_LOGO_PANEL}" class="logo-panel" onclick="toggleCentral()">
            </header>
            <div class="controls-bar">
                <select id="mes-selector" onchange="cambiarMes()">{opciones_dropdown}</select>
                <button class="tab-btn active" id="btn-inc" onclick="showGlobalTab('incs')">INCIDENCIAS</button>
                <button class="tab-btn" id="btn-cob" onclick="showGlobalTab('cobs')">COBROS Y PERDIDA MITIGADA</button>
                <button class="tab-btn" id="btn-hon" onclick="showGlobalTab('honor')">MEJOR RENDIMIENTO</button>
                <button class="tab-btn" id="btn-peo" onclick="showGlobalTab('peores')">RENDIMIENTO CRÍTICO</button>
                <button class="tab-btn btn-sync" onclick="sincronizarServidor()">🔄 SINCRONIZAR</button>
            </div>
            <main class="main-content" id="main-panel">{html_meses_data}</main>
            <script>
                function sincronizarServidor() {{
                    if(!confirm("¿Deseas actualizar los datos desde el servidor?")) return;
                    event.target.innerText = "⌛ CONECTANDO...";
                    fetch('http://192.168.7.7:5000/guardar', {{ method: 'POST' }})
                    .then(res => res.json()).then(() => {{ location.reload(); }})
                    .catch(() => alert("❌ Error de conexión al servidor Flask"));
                }}
                function toggleCentral() {{ document.body.classList.toggle('modo-central'); }}
                function cambiarMes() {{
                    document.querySelectorAll(".mes-container").forEach(e => e.classList.remove('active'));
                    let sel = document.getElementById("mes-selector").value;
                    if(sel) document.getElementById(sel).classList.add('active');
                    actualizarPestañaInterna();
                }}
                function showGlobalTab(t) {{
                    localStorage.setItem('pestañaActiva', t);
                    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                    const btnId = 'btn-' + t.substring(0,3);
                    if(document.getElementById(btnId)) document.getElementById(btnId).classList.add('active');
                    let mes = document.querySelector('.mes-container.active');
                    if(mes) {{
                        mes.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                        let contentId = t + "-" + mes.id.replace('mes-', '');
                        if(document.getElementById(contentId)) document.getElementById(contentId).classList.add('active');
                    }}
                }}
                function actualizarPestañaInterna() {{ showGlobalTab(localStorage.getItem('pestañaActiva') || 'incs'); }}
                window.onload = cambiarMes;
            </script></body></html>"""
        
        with open(os.path.join(ruta_raiz, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_final)
        
        print("\n✅ PANEL 'index.html' GENERADO EXITOSAMENTE.")

    except Exception as e: print(f"\n❌ ERROR CRÍTICO: {e}"); input()

if __name__ == "__main__": generar_panel_luxor_centralizado()