import pandas as pd
import os
import sys
import re
import warnings
import threading
import time

# =================================================================
# ID: SCRIPT DE GENERACIÓN DE REPORTES DE INCIDENCIAS V3.3.3
# =================================================================

warnings.filterwarnings("ignore")

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

RUTA_LOGO = "../../RECURSOS/logo.png"
ARCHIVO_SUCURSALES = "sucursales.txt"

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[^\w\s-]', '', str(nombre)).strip().replace(' ', '_')

def obtener_sucursales_txt(ruta_base):
    ruta_txt = os.path.join(ruta_base, ARCHIVO_SUCURSALES)
    if not os.path.exists(ruta_txt):
        with open(ruta_txt, "w", encoding="utf-8") as f: f.write("CENTRAL")
    with open(ruta_txt, "r", encoding="utf-8") as f:
        return [line.strip().upper() for line in f if line.strip()]

def obtener_links_fotos(sucursal, mes, nombre_archivo_n):
    CARPETA_FOTOS = "fotos_incidencias" 
    links = []
    nombre_n = str(nombre_archivo_n).strip()
    if nombre_n and nombre_n.lower() != 'nan' and nombre_n != '-':
        nombres = nombre_n.split()
        for n in nombres:
            ruta_relativa = f"../../{CARPETA_FOTOS}/{mes}/{sucursal}/{n}"
            links.append(f"<a href='#' onclick=\"abrirModal('{ruta_relativa}')\" class='link-incidencias'>📷 Ver Foto</a>")
    return "<br>" + " ".join(links) if links else ""

CSS_UNIFICADO = f"""
<style>
    body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background-color: #f8f9fa; margin: 0; padding: 0; text-align: center; }}
    .top-bar {{ height: 100px; background: white; border-bottom: 4px solid #F9D908; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; margin-bottom: 30px; }}
    .logo-ext {{ height: 60px; max-width: 100px; object-fit: contain; }}
    h1 {{ color: #0844a4; margin: 0; text-transform: uppercase; font-weight: 900; font-size: 18px; text-align: center; flex-grow: 1; padding: 0 10px; }}
    .main-container {{ background-color: white; width: 95%; margin: 0 auto 40px auto; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); display: inline-block; text-align: center; overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; border: 2px solid #333; margin: 0 auto; }}
    th {{ background-color: #0844a4; color: #F9D908; padding: 10px; font-size: 12px; text-transform: uppercase; border: 2px solid #333; }}
    td {{ border: 2px solid #333; padding: 8px; text-align: center; font-size: 11px; font-weight: bold; color: black; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    .btn-volver {{ display: inline-block; cursor: pointer; margin: 20px 5px; padding: 10px 20px; background-color: #0844a4; color: #ffffff !important; text-decoration: none; font-weight: 900; font-size: 13px; text-transform: uppercase; border-radius: 50px; border: 2px solid #F9D908; }}
    .ranking-title {{ background: #ed1c24; color: white; padding: 10px; margin-top: 20px; font-size: 14px; font-weight: 900; border: 2px solid #333; border-bottom: none; }}
    .link-incidencias {{ color: #0844a4; text-decoration: underline; font-size: 10px; font-weight: bold; cursor: pointer; margin-right: 5px; }}
    #modalFoto {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.8); }}
    .modal-content {{ margin: auto; display: block; max-width: 80%; max-height: 80%; margin-top: 50px; border: 5px solid white; }}
    .close-modal {{ position: absolute; top: 20px; right: 35px; color: white; font-size: 40px; font-weight: bold; cursor: pointer; }}
</style>
<div id="modalFoto" onclick="cerrarModal()">
    <span class="close-modal">&times;</span>
    <img class="modal-content" id="imgModal">
</div>
<script>
    function abrirModal(ruta) {{
        document.getElementById("imgModal").src = ruta;
        document.getElementById("modalFoto").style.display = "block";
    }}
    function cerrarModal() {{
        document.getElementById("modalFoto").style.display = "none";
    }}
</script>
"""

def generar_reporte_v30_final():
    try:
        print("\n" + "="*60)
        print(">>> INICIANDO PROCESAMIENTO GLOBAL DE REPORTES <<<")
        print("="*60)
        
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        sucursales_txt = obtener_sucursales_txt(ruta_base)
        
        # Leemos todos los Excel desde la raíz (ya no solo carpeta cuadros)
        archivos = [os.path.join(root, f) for root, dirs, files in os.walk(ruta_base) for f in files if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
        
        if not archivos:
            print("⚠️ No se encontraron archivos Excel."); return

        lista_df = []
        for f in archivos:
            try:
                df_t = pd.read_excel(f, engine='openpyxl')
                df_t.columns = [str(c).strip().upper() for c in df_t.columns]
                for col in df_t.columns:
                    if 'SUCURSAL' in col: df_t.rename(columns={col: 'SUCURSAL'}, inplace=True)
                    if 'FECHA' in col: df_t.rename(columns={col: 'FECHA'}, inplace=True)
                    if 'NOMBRE' in col or 'AUDITOR' in col: df_t.rename(columns={col: 'RESPONSABLE'}, inplace=True)
                if len(df_t.columns) >= 14: df_t.rename(columns={df_t.columns[13]: 'FOTO_INCIDENCIAS'}, inplace=True)
                if len(df_t.columns) >= 7: df_t['INC_LIMPIA'] = df_t.iloc[:, 6].astype(str).str.strip().str.upper().str.replace('.', '', regex=False)
                df_t['FECHA'] = pd.to_datetime(df_t['FECHA'], errors='coerce')
                lista_df.append(df_t)
            except Exception as e: print(f" [!] Error leyendo {os.path.basename(f)}: {e}")

        df_master = pd.concat(lista_df, ignore_index=True).dropna(subset=['FECHA', 'SUCURSAL'])
        df_master['SUCURSAL'] = df_master['SUCURSAL'].astype(str).str.strip().str.upper()
        df_master['PERIODO'] = df_master['FECHA'].dt.to_period('M')
        
        periodos = sorted(df_master['PERIODO'].unique())
        columnas_reporte = ['RESPONSABLE', 'PROVEEDOR', 'FECHA', 'FACTURA', 'INCIDENCIA', 'TIPO FISCALIZACIÓN', 'OBSERVACIÓN']
        grupos_defs = [
            {"tipo": "A", "porc": 10, "color": "#f4faf0", "items": ["NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO", "FALTA SELLO, FIRMA O CÉDULA", "DOCUMENTO NO LEGIBLE"]},
            {"tipo": "B", "porc": 15, "color": "#f0f7ff", "items": ["DOCUMENTACIÓN ERRÓNEA", "FISCALIZACIÓN A DESTIEMPO", "PRODUCTO O SKU DUPLICADO", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN"]},
            {"tipo": "C", "porc": 20, "color": "#fffef0", "items": ["FISCALIZACIÓN CON USUARIO NO CORRESPONDIENTE", "ERROR DE KG EN TARA", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS"]},
            {"tipo": "D", "porc": 25, "color": "#fff8f0", "items": ["NO SE IMPORTÓ DIFERENCIA AL DORSO DE LA FACTURA", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO"]},
            {"tipo": "E", "porc": 30, "color": "#fff0f0", "items": ["RECEPCIÓN SIN AUTORIZACIÓN DE CMF", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA"]}
        ]

        for p_idx, p_act in enumerate(periodos):
            n_m_act = MESES_ES[p_act.month]
            p_ant = p_act - 1
            n_m_ant = MESES_ES[p_ant.month] if p_ant.month in MESES_ES else "INICIO"
            
            print(f"\n>>> PROCESANDO MES: {n_m_act} ({p_idx + 1}/{len(periodos)})")
            df_m_act = df_master[df_master['PERIODO'] == p_act]

            for n_s in sucursales_txt:
                n_s_limpio = n_s.strip().upper()
                df_suc_act = df_m_act[df_m_act['SUCURSAL'] == n_s_limpio]
                p_f = os.path.join(ruta_base, n_m_act, n_s_limpio); os.makedirs(p_f, exist_ok=True)
                
                html_count = 0

                def generar_filas_con_fotos(dataframe):
                    cuerpo = ""
                    for _, row in dataframe.iterrows():
                        links = obtener_links_fotos(n_s_limpio, n_m_act, row.get('FOTO_INCIDENCIAS', ''))
                        cuerpo += "<tr>"
                        for col in columnas_reporte:
                            val = row.get(col, '-')
                            if col == 'FECHA' and not pd.isna(val): val = pd.to_datetime(val).strftime('%Y-%m-%d')
                            if col == 'OBSERVACIÓN': val = f"{val} {links}"
                            cuerpo += f"<td>{val}</td>"
                        cuerpo += "</tr>"
                    return cuerpo

                ranking_html = ""
                if not df_suc_act.empty and 'RESPONSABLE' in df_suc_act.columns:
                    top_inc = df_suc_act['RESPONSABLE'].value_counts().head(3)
                    for nombre, total in top_inc.items():
                        archivo_persona = f"INCIDENCIAS_{limpiar_nombre_archivo(nombre)}.html"
                        ranking_html += f"<tr><td style='text-align:left;'>{nombre}</td><td><a href='{archivo_persona}' class='link-incidencias'>{total}</a></td></tr>"
                        cuerpo_p = generar_filas_con_fotos(df_suc_act[df_suc_act['RESPONSABLE'] == nombre])
                        with open(os.path.join(p_f, archivo_persona), "w", encoding="utf-8") as f:
                            f.write(f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{nombre}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in columnas_reporte])}</tr></thead><tbody>{cuerpo_p}</tbody></table><a onclick='window.history.back()' class='btn-volver'>VOLVER</a></div></body></html>")
                        html_count += 1

                df_otras = df_suc_act[df_suc_act['INC_LIMPIA'] == "OTRAS"]
                c_otras_act = len(df_otras)
                c_otras_ant = len(df_master[(df_master['SUCURSAL']==n_s_limpio) & (df_master['PERIODO']==p_ant) & (df_master['INC_LIMPIA']=="OTRAS")])
                v_otras_ant = "0" if c_otras_ant == 0 else f"<a href='../../{n_m_ant}/{n_s_limpio}/OTRAS.html' class='link-incidencias'>{c_otras_ant}</a>"
                v_otras_act = "0" if c_otras_act == 0 else f"<a href='OTRAS.html' class='link-incidencias'>{c_otras_act}</a>"
                filas_html_txt = f"<tr style='background-color:#ffffff; border-bottom: 4px solid #333;'><td style='text-align:left;'>OTRAS.</td><td>-</td><td>{v_otras_ant}</td><td>{v_otras_act}</td><td>0%</td></tr>"
                
                if c_otras_act > 0:
                    cuerpo_o = generar_filas_con_fotos(df_otras)
                    with open(os.path.join(p_f, "OTRAS.html"), "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>OTRAS</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in columnas_reporte])}</tr></thead><tbody>{cuerpo_o}</tbody></table><a onclick='window.history.back()' class='btn-volver'>VOLVER</a></div></body></html>")
                    html_count += 1

                suma_impacto, t_act, t_ant = 0, c_otras_act, c_otras_ant
                for grupo in grupos_defs:
                    g_act_grupo, g_ant_grupo = 0, 0
                    temp_filas = []
                    for idx_item, inc in enumerate(grupo["items"]):
                        busq = inc.upper().replace('.', '')
                        nombre_file = f"{limpiar_nombre_archivo(inc)}.html"
                        df_inc = df_suc_act[df_suc_act['INC_LIMPIA'] == busq]
                        c_act = len(df_inc); c_ant = len(df_master[(df_master['SUCURSAL']==n_s_limpio) & (df_master['PERIODO']==p_ant) & (df_master['INC_LIMPIA']==busq)])
                        g_act_grupo += c_act; g_ant_grupo += c_ant; t_act += c_act; t_ant += c_ant
                        v_ant = "0" if c_ant == 0 else f"<a href='../../{n_m_ant}/{n_s_limpio}/{nombre_file}' class='link-incidencias'>{c_ant}</a>"
                        v_act = "0" if c_act == 0 else f"<a href='{nombre_file}' class='link-incidencias'>{c_act}</a>"
                        estilo_separador = "border-bottom: 4px solid #333;" if idx_item == len(grupo["items"]) - 1 else ""
                        temp_filas.append(f"<tr style='background-color:{grupo['color']}; {estilo_separador}'><td style='text-align:left;'>{inc}.</td><td>{grupo['tipo']}</td><td>{v_ant}</td><td>{v_act}</td>")
                        if c_act > 0:
                            cuerpo = generar_filas_con_fotos(df_inc)
                            with open(os.path.join(p_f, nombre_file), "w", encoding="utf-8") as f:
                                f.write(f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{inc}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in columnas_reporte])}</tr></thead><tbody>{cuerpo}</tbody></table><a onclick='window.history.back()' class='btn-volver'>VOLVER</a></div></body></html>")
                            html_count += 1
                    
                    if g_act_grupo > g_ant_grupo: suma_impacto += grupo["porc"]
                    color_p = "#ffcccc" if g_act_grupo > g_ant_grupo else "#ccffcc"
                    for idx, f_base in enumerate(temp_filas):
                        filas_html_txt += f_base + (f"<td rowspan='{len(grupo['items'])}' style='background-color:{color_p};'>{grupo['porc']}%</td></tr>" if idx == 0 else "</tr>")

                nota_f = max(0, 100 - suma_impacto)
                color_eval = "#ed1c24" if nota_f < 75 else "#27ae60"
                link_total_ant = f"<a href='../../{n_m_ant}/{n_s_limpio}/todo_el_mes.html' style='color:white;'>{t_ant}</a>" if t_ant > 0 else "0"
                link_total_act = f"<a href='todo_el_mes.html' style='color:white;'>{t_act}</a>" if t_act > 0 else "0"

                cuerpo_mes = generar_filas_con_fotos(df_suc_act) if not df_suc_act.empty else "<tr><td colspan='7'>Sin incidencias</td></tr>"
                with open(os.path.join(p_f, "todo_el_mes.html"), "w", encoding="utf-8") as f: 
                    f.write(f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>TOTAL {n_m_act}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in columnas_reporte])}</tr></thead><tbody>{cuerpo_mes}</tbody></table><a onclick='window.history.back()' class='btn-volver'>VOLVER</a></div></body></html>")
                
                with open(os.path.join(p_f, "reporte.html"), "w", encoding="utf-8") as f: 
                    f.write(f"""<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body>
                        <div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{n_s_limpio} - {n_m_act}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div>
                        <div class='main-container'>
                            <table><thead><tr><th>INCIDENCIA</th><th>TIPO</th><th>{n_m_ant}</th><th>{n_m_act}</th><th>PORCENTAJE</th></tr></thead>
                            <tbody>{filas_html_txt}
                                <tr style='background-color:#0844a4;'><td style='text-align:right; color:white;' colspan='2'>TOTAL / CALIFICACIÓN</td><td style='color:white;'>{link_total_ant}</td><td style='color:white;'>{link_total_act}</td><td style='background:{color_eval}; color:white;'>{nota_f}%</td></tr>
                            </tbody></table>
                            <div class='ranking-title'>PERSONAS CON MAYOR NÚMERO DE INCIDENCIAS</div>
                            <table><tbody>{ranking_html if ranking_html else "<tr><td>Sin datos</td></tr>"}</tbody></table>
                            <br><a onclick='window.history.back()' class='btn-volver'>VOLVER</a>
                        </div></body></html>""")
                
                print(f"    -> Sucursal: {n_s_limpio.ljust(15)} | OK! ({html_count + 2} HTMLs)")

        print("\n" + "="*60)
        print("✅ TODOS LOS MESES PROCESADOS EXITOSAMENTE.")
        print("="*60)
        
    except Exception as e: print(f"\n❌ ERROR CRÍTICO: {e}")

    def cerrar_programa():
        for i in range(10, 0, -1):
            sys.stdout.write(f"\rCERRANDO CMD EN {i} SEGUNDOS... (Presione ENTER para salir ahora)")
            sys.stdout.flush()
            time.sleep(1)
        os._exit(0)

    threading.Thread(target=cerrar_programa, daemon=True).start()
    input()

if __name__ == "__main__":
    generar_reporte_v30_final()