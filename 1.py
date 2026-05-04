import pandas as pd
import os
import sys
import re
import warnings

# Silenciar advertencias de validación de Excel
warnings.filterwarnings("ignore")

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

RUTA_LOGO = "../../RECURSOS/logo.png"

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[^\w\s-]', '', str(nombre)).strip().replace(' ', '_')

CSS_UNIFICADO = f"""
<style>
    body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background-color: #f8f9fa; margin: 0; padding: 0; text-align: center; }}
    .top-bar {{ height: 100px; background: white; border-bottom: 4px solid #F9D908; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; margin-bottom: 30px; }}
    .logo-ext {{ height: 60px; max-width: 100px; object-fit: contain; }}
    h1 {{ color: #0844a4; margin: 0; text-transform: uppercase; font-weight: 900; font-size: 18px; text-align: center; flex-grow: 1; padding: 0 10px; }}
    .main-container {{ background-color: white; width: 90%; margin: 0 auto 40px auto; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); display: inline-block; text-align: center; overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; border: 2px solid #333; margin: 0 auto; }}
    th {{ background-color: #0844a4; color: #F9D908; padding: 10px; font-size: 12px; text-transform: uppercase; border: 2px solid #333; }}
    td {{ border: 2px solid #333; padding: 8px; text-align: center; font-size: 11px; font-weight: bold; color: black; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    .btn-volver {{ display: inline-block; margin-top: 25px; padding: 10px 20px; background-color: #0844a4; color: #ffffff !important; text-decoration: none; font-weight: 900; font-size: 13px; text-transform: uppercase; border-radius: 50px; border: 2px solid #F9D908; }}
</style>
<script>
    function volverSmart() {{
        const urlParams = new URLSearchParams(window.location.search);
        const retorno = urlParams.get('retorno');
        if (retorno) {{ window.location.href = retorno; }}
        else {{ window.location.href = 'reporte.html'; }}
    }}
</script>
"""

def generar_reporte_v30_final():
    try:
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_base, "cuadros")
        archivos = [os.path.join(root, f) for root, dirs, files in os.walk(ruta_cuadros) for f in files if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

        lista_df = []
        for f in archivos:
            try:
                df_t = pd.read_excel(f, engine='openpyxl')
                if len(df_t.columns) >= 7:
                    df_t['INC_LIMPIA'] = df_t.iloc[:, 6].astype(str).str.strip().str.upper().str.replace('.', '', regex=False)
                df_t.columns = [str(c).strip().upper() for c in df_t.columns]
                if 'RESPONSABLE' in df_t.columns: df_t['NOMBRE_AUX'] = df_t['RESPONSABLE']
                if len(df_t.columns) >= 4 and 'FECHA' not in df_t.columns:
                    df_t.rename(columns={df_t.columns[3]: 'FECHA'}, inplace=True)
                df_t['FECHA'] = pd.to_datetime(df_t['FECHA'], errors='coerce')
                lista_df.append(df_t)
            except: continue
        
        df_master = pd.concat(lista_df, ignore_index=True).dropna(subset=['FECHA', 'SUCURSAL'])
        df_master['PERIODO'] = df_master['FECHA'].dt.to_period('M')
        
        columnas_reporte = ['RESPONSABLE', 'PROVEEDOR', 'FECHA', 'FACTURA', 'INCIDENCIA', 'TIPO FISCALIZACIÓN', 'OBSERVACIÓN']
        grupos_defs = [
            {"tipo": "A", "porc": 10, "color": "#f4faf0", "items": ["NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO", "FALTA SELLO, FIRMA O CÉDULA", "DOCUMENTO NO LEGIBLE"]},
            {"tipo": "B", "porc": 15, "color": "#f0f7ff", "items": ["DOCUMENTACIÓN ERRÓNEA", "FISCALIZACIÓN A DESTIEMPO", "PRODUCTO O SKU DUPLICADO", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN"]},
            {"tipo": "C", "porc": 20, "color": "#fffef0", "items": ["FISCALIZACIÓN CON USUARIO NO CORRESPONDIENTE", "ERROR DE KG EN TARA", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS"]},
            {"tipo": "D", "porc": 25, "color": "#fff8f0", "items": ["NO SE INDICÓ DIFERENCIA AL DORSO DE LA FACTURA", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO"]},
            {"tipo": "E", "porc": 30, "color": "#fff0f0", "items": ["RECEPCIÓN SIN AUTORIZACIÓN DE CMF", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA"]}
        ]

        for p_act in sorted(df_master['PERIODO'].unique()):
            n_m_act = MESES_ES[p_act.month]; p_ant = p_act - 1
            n_m_ant = MESES_ES[p_ant.month] if p_ant.month in MESES_ES else "ANT."
            df_m_act = df_master[df_master['PERIODO'] == p_act]

            for suc in sorted(df_m_act['SUCURSAL'].dropna().unique()):
                n_s = str(suc).strip().upper(); df_suc_act = df_m_act[df_m_act['SUCURSAL'] == suc]
                filas_html, suma_impacto, t_act, t_ant = "", 0, 0, 0

                for grupo in grupos_defs:
                    g_act_grupo, g_ant_grupo = 0, 0
                    temp_filas = []
                    for inc in grupo["items"]:
                        busq = inc.upper().replace('.', '')
                        df_det_act = df_suc_act[df_suc_act['INC_LIMPIA'] == busq].copy().drop_duplicates()
                        c_act = len(df_det_act)
                        df_det_ant = df_master[(df_master['SUCURSAL']==suc) & (df_master['PERIODO']==p_ant) & (df_master['INC_LIMPIA']==busq)].copy().drop_duplicates()
                        c_ant = len(df_det_ant)
                        
                        g_act_grupo += c_act; g_ant_grupo += c_ant; t_act += c_act; t_ant += c_ant
                        
                        v_ant = "0" if c_ant == 0 else f"<a href='../../{n_m_ant}/{n_s}/{limpiar_nombre_archivo(inc)}.html?retorno=../../{n_m_act}/{n_s}/reporte.html'>{c_ant}</a>"
                        v_act = "0" if c_act == 0 else f"<a href='{limpiar_nombre_archivo(inc)}.html?retorno=reporte.html'>{c_act}</a>"
                        temp_filas.append(f"<tr style='background-color:{grupo['color']};'><td style='text-align:left;'>{inc}.</td><td>{grupo['tipo']}</td><td>{v_ant}</td><td>{v_act}</td>")
                        
                        if c_act > 0:
                            df_det_act['FECHA'] = pd.to_datetime(df_det_act['FECHA']).dt.strftime('%Y-%m-%d')
                            cols_p = [c for c in columnas_reporte if c in df_det_act.columns]
                            cuerpo = "".join([f"<tr>{''.join([f'<td>{v}</td>' for v in r])}</tr>" for r in df_det_act[cols_p].fillna('-').values])
                            html_inc = f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{inc}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in cols_p])}</tr></thead><tbody>{cuerpo}</tbody></table><a href='#' onclick='volverSmart()' class='btn-volver'>VOLVER</a></div></body></html>"
                            f_path = os.path.join(ruta_base, n_m_act, n_s, f"{limpiar_nombre_archivo(inc)}.html")
                            os.makedirs(os.path.dirname(f_path), exist_ok=True); 
                            with open(f_path, "w", encoding="utf-8") as f: f.write(html_inc)

                    if g_act_grupo > g_ant_grupo: suma_impacto += grupo["porc"]
                    color_p = "#ffcccc" if g_act_grupo > g_ant_grupo else "#ccffcc"
                    for idx, f_base in enumerate(temp_filas):
                        filas_html += f_base + (f"<td rowspan='{len(grupo['items'])}' style='background-color:{color_p};'>{grupo['porc']}%</td></tr>" if idx == 0 else "</tr>")

                nota_f = max(0, 100 - suma_impacto)
                p_f = os.path.join(ruta_base, n_m_act, n_s); os.makedirs(p_f, exist_ok=True)
                
                # ENLACES EN LOS TOTALES CON RETORNO INTELIGENTE
                link_t_ant = f"<a href='../../{n_m_ant}/{n_s}/reporte.html' style='color:white;'>{t_ant}</a>" if t_ant > 0 else "0"
                link_t_act = f"<a href='reporte.html' style='color:white;'>{t_act}</a>" if t_act > 0 else "0"

                with open(os.path.join(p_f, "reporte.html"), "w", encoding="utf-8") as f: 
                    # El botón VOLVER AL PANEL ahora apunta al ancla del mes actual
                    f.write(f"<html><head><meta charset='UTF-8'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{n_s} - {n_m_act}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr><th>INCIDENCIA</th><th>TIPO</th><th>{n_m_ant}</th><th>{n_m_act}</th><th>APROBATORIO 75%</th></tr></thead><tbody>{filas_html}<tr style='background-color:#0844a4;'><td style='text-align:right; color:white;' colspan='2'>TOTAL / CALIFICACIÓN</td><td style='color:white;'>{link_t_ant}</td><td style='color:white;'>{link_t_act}</td><td style='background:{'#ed1c24' if nota_f < 75 else '#27ae60'}; color:white;'>{nota_f}%</td></tr></tbody></table><br><a href='../../index.html#mes-{n_m_act}' class='btn-volver'>VOLVER AL PANEL</a></div></body></html>")
                with open(os.path.join(p_f, "solo_mes.html"), "w", encoding="utf-8") as f: 
                    f.write(f"<table><tr><td>{n_m_act}</td><td>{t_act}</td><td>{nota_f}%</td></tr></table>")

        print("\n✅ Botones de Volver al Panel configurados con anclas inteligentes.")
    except Exception as e: print(f"\n❌ ERROR: {e}")

if __name__ == "__main__":
    generar_reporte_v30_final()