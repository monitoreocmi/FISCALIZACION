import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import sys
import re

# Configuración de codificación para evitar errores en consola
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# Ajuste de ruta para compatibilidad con servidores web (GitHub)
RUTA_LOGO = "../../RECURSOS/logo.png"

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[^\w\s-]', '', str(nombre)).strip().replace(' ', '_')

CSS_UNIFICADO = f"""
<style>
    body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background-color: #f8f9fa; margin: 0; padding: 0; }}
    .top-bar {{ height: 100px; background: white; border-bottom: 4px solid #F9D908; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; margin-bottom: 30px; }}
    .logo-ext {{ height: 60px; max-width: 100px; object-fit: contain; }}
    h1 {{ color: #0844a4; margin: 0; text-transform: uppercase; font-weight: 900; font-size: 18px; text-align: center; flex-grow: 1; padding: 0 10px; }}
    .main-container {{ background-color: white; max-width: 95%; margin: 0 auto 40px auto; padding: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; border: 2px solid #333; min-width: 600px; }}
    th {{ background-color: #0844a4; color: #F9D908; padding: 10px; font-size: 12px; text-transform: uppercase; border: 2px solid #333; }}
    td {{ border: 2px solid #333; padding: 8px; text-align: center; font-size: 11px; font-weight: bold; color: black; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    .btn-volver {{ display: inline-block; margin-top: 25px; padding: 10px 20px; background-color: #0844a4; color: #ffffff !important; text-decoration: none; font-weight: 900; font-size: 13px; text-transform: uppercase; border-radius: 50px; border: 2px solid #F9D908; }}
    
    /* Ajustes para Resoluciones Pequeñas (Laptops y Móviles) */
    @media (max-width: 768px) {{
        .top-bar {{ height: 80px; padding: 0 10px; }}
        .logo-ext {{ height: 40px; }}
        h1 {{ font-size: 14px; }}
        td, th {{ font-size: 10px; padding: 5px; }}
        .main-container {{ width: 100%; max-width: 100%; padding: 5px; box-shadow: none; }}
    }}
</style>
"""

def generar_reporte_v30_final():
    try:
        print("\n" + "="*60)
        print(">>> GENERADOR V30 - OPTIMIZADO PARA MÓVIL Y GITHUB <<<")
        print("="*60 + "\n")
        
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        archivos = filedialog.askopenfilenames(title="Seleccionar archivos Excel")
        if not archivos: return

        lista_df = []
        for f in archivos:
            df_t = pd.read_excel(f, engine='openpyxl')
            df_t.columns = [str(c).strip().upper() for c in df_t.columns]
            if 'RESPONSABLE' in df_t.columns: df_t['NOMBRE_AUX'] = df_t['RESPONSABLE']
            if len(df_t.columns) >= 4 and 'FECHA' not in df_t.columns:
                df_t.rename(columns={df_t.columns[3]: 'FECHA'}, inplace=True)
            if 'INCIDENCIA' in df_t.columns:
                df_t['INC_LIMPIA'] = df_t['INCIDENCIA'].astype(str).str.strip().str.upper().str.replace('.', '', regex=False)
            df_t['FECHA'] = pd.to_datetime(df_t['FECHA'], errors='coerce')
            lista_df.append(df_t)
        
        df_master = pd.concat(lista_df, ignore_index=True).dropna(subset=['FECHA', 'SUCURSAL'])
        df_master['PERIODO'] = df_master['FECHA'].dt.to_period('M')
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        
        grupos_defs = [
            {"tipo": "A", "porc": 10, "color": "#f4faf0", "items": ["NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO", "FALTA SELLO, FIRMA O CÉDULA", "DOCUMENTO NO LEGIBLE"]},
            {"tipo": "B", "porc": 15, "color": "#f0f7ff", "items": ["DOCUMENTACIÓN ERRÓNEA", "FISCALIZACIÓN A DESTIEMPO", "PRODUCTO O SKU DUPLICADO", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN"]},
            {"tipo": "C", "porc": 20, "color": "#fffef0", "items": ["FISCALIZACIÓN CON USUARIO NO CORRESPONDIENTE", "ERROR DE KG EN TARA", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS"]},
            {"tipo": "D", "porc": 25, "color": "#fff8f0", "items": ["NO SE INDICÓ DIFERENCIA AL DORSO DE LA FACTURA", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO"]},
            {"tipo": "E", "porc": 30, "color": "#fff0f0", "items": ["RECEPCIÓN SIN AUTORIZACIÓN DE CMF", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA"]}
        ]

        for p_act in sorted(df_master['PERIODO'].unique()):
            n_m_act = MESES_ES[p_act.month]
            p_ant = p_act - 1
            n_m_ant = MESES_ES[p_ant.month] if p_ant.month in MESES_ES else "ANT."
            df_m_act = df_master[df_master['PERIODO'] == p_act]
            
            sucursales = sorted(df_m_act['SUCURSAL'].dropna().unique())
            for suc in sucursales:
                n_s = str(suc).strip().upper()
                df_suc_act = df_m_act[df_m_act['SUCURSAL'] == suc]
                filas_html, suma_impacto, t_act, t_ant = "", 0, 0, 0
                
                for grupo in grupos_defs:
                    g_act_grupo, g_ant_grupo = 0, 0
                    num_items = len(grupo["items"])
                    temp_filas = []
                    for i, inc in enumerate(grupo["items"]):
                        busq = inc.upper().replace('.', '')
                        c_act = int(len(df_suc_act[df_suc_act['INC_LIMPIA']==busq]))
                        c_ant = int(len(df_master[(df_master['SUCURSAL']==suc) & (df_master['PERIODO']==p_ant) & (df_master['INC_LIMPIA']==busq)]))
                        g_act_grupo += c_act; g_ant_grupo += c_ant; t_act += c_act; t_ant += c_ant
                        v_ant = "0" if c_ant == 0 else f"<a href='../../{n_m_ant}/{n_s}/{limpiar_nombre_archivo(inc)}.html'>{c_ant}</a>"
                        v_act = "0" if c_act == 0 else f"<a href='{limpiar_nombre_archivo(inc)}.html'>{c_act}</a>"
                        temp_filas.append(f"<tr style='background-color:{grupo['color']};'><td style='text-align:left;'>{inc}.</td><td>{grupo['tipo']}</td><td>{v_ant}</td><td>{v_act}</td>")
                        
                        if c_act > 0:
                            df_inc_det = df_suc_act[df_suc_act['INC_LIMPIA'] == busq].copy()
                            df_inc_det['FECHA'] = df_inc_det['FECHA'].dt.strftime('%Y-%m-%d')
                            cols = [c for c in ['PROVEEDOR', 'FACTURA', 'FECHA', 'RESPONSABLE', 'OBSERVACIÓN'] if c in df_inc_det.columns]
                            html_inc = f"<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{inc}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in cols])}</tr></thead><tbody>{''.join([f'<tr>{"".join([f"<td>{v}</td>" for v in r])}</tr>' for r in df_inc_det[cols].values])}</tbody></table><a href='reporte.html' class='btn-volver'>VOLVER AL REPORTE</a></div></body></html>"
                            f_det = os.path.join(ruta_base, n_m_act, n_s, f"{limpiar_nombre_archivo(inc)}.html")
                            os.makedirs(os.path.dirname(f_det), exist_ok=True); 
                            with open(f_det, "w", encoding="utf-8") as f: f.write(html_inc)
                    
                    color_porc = "#ffcccc" if g_act_grupo > g_ant_grupo else "#ccffcc"
                    if g_act_grupo > g_ant_grupo: suma_impacto += grupo["porc"]
                    for idx_f, fila_base in enumerate(temp_filas):
                        filas_html += fila_base + (f"<td rowspan='{num_items}' style='background-color:{color_porc};'>{grupo['porc']}%</td></tr>" if idx_f == 0 else "</tr>")

                nota_f = max(0, 100 - suma_impacto)
                color_calif = "#ed1c24" if nota_f < 75 else "#27ae60"
                link_t_ant = f"<a href='../../{n_m_ant}/{n_s}/TODAS.html' style='color:white;'>{t_ant}</a>" if t_ant > 0 else "0"
                link_t_act = f"<a href='TODAS.html' style='color:white;'>{t_act}</a>" if t_act > 0 else "0"

                filas_resp = ""
                if 'NOMBRE_AUX' in df_suc_act.columns:
                    for name, cant in df_suc_act['NOMBRE_AUX'].value_counts().head(3).items():
                        filas_resp += f"<tr style='background-color:#ffcccc;'><td style='text-align:left;'>{name}</td><td colspan='4'><a href='RESP_{limpiar_nombre_archivo(name)}.html' style='color:#ed1c24; font-weight:bold;'>{cant}</a></td></tr>"
                        df_res_det = df_suc_act[df_suc_act['NOMBRE_AUX'] == name].copy()
                        df_res_det['FECHA'] = df_res_det['FECHA'].dt.strftime('%Y-%m-%d')
                        cols_res = [c for c in ['FACTURA', 'FECHA', 'INCIDENCIA', 'OBSERVACIÓN'] if c in df_res_det.columns]
                        html_res = f"<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>RESPONSABLE: {name}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in cols_res])}</tr></thead><tbody>{''.join([f'<tr>{"".join([f"<td>{v}</td>" for v in r])}</tr>' for r in df_res_det[cols_res].values])}</tbody></table><a href='reporte.html' class='btn-volver'>VOLVER AL REPORTE</a></div></body></html>"
                        with open(os.path.join(ruta_base, n_m_act, n_s, f"RESP_{limpiar_nombre_archivo(name)}.html"), "w", encoding="utf-8") as f: f.write(html_res)

                # Reporte TODAS
                df_all = df_suc_act.copy()
                df_all['FECHA'] = df_all['FECHA'].dt.strftime('%Y-%m-%d')
                cols_finales = [c for c in ['PROVEEDOR', 'FACTURA', 'FECHA', 'TIPO FISCALIZACIÓN', 'RESPONSABLE', 'INCIDENCIA', 'OBSERVACIÓN'] if c in df_all.columns]
                html_all = f"<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{CSS_UNIFICADO}</head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>TODAS LAS INCIDENCIAS</h1><div class='main-container'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in cols_finales])}</tr></thead><tbody>{''.join([f'<tr>{"".join([f"<td>{v}</td>" for v in r])}</tr>' for r in df_all[cols_finales].fillna('-').values])}</tbody></table><a href='reporte.html' class='btn-volver'>VOLVER AL REPORTE</a></div></body></html>"
                with open(os.path.join(ruta_base, n_m_act, n_s, "TODAS.html"), "w", encoding="utf-8") as f: f.write(html_all)

                # Reporte Principal
                html_final = f"""<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{CSS_UNIFICADO}</head><body>
                    <div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext'><h1>{n_s} - {n_m_act}</h1><img src='{RUTA_LOGO}' class='logo-ext'></div>
                    <div class='main-container'>
                        <table>
                            <thead><tr><th>INCIDENCIA</th><th>TIPO</th><th>{n_m_ant}</th><th>{n_m_act}</th><th>APROBATORIO 75%</th></tr></thead>
                            <tbody>
                                {filas_html}
                                <tr class='row-total' style='background-color:#0844a4;'><td style='text-align:right; color:white;' colspan='2'>TOTAL GENERAL / CALIFICACIÓN</td><td style='color:white;'>{link_t_ant}</td><td style='color:white;'>{link_t_act}</td><td style='background:{color_calif}; color:white;'>{nota_f}%</td></tr>
                                <tr><td colspan='5' style='font-weight:900; background:#f0f0f0; padding:10px; color:#002060;'>RESPONSABLES CON MAYOR INCIDENCIA</td></tr>
                                {filas_resp}
                            </tbody>
                        </table><br><a href='../../index.html' class='btn-volver'>VOLVER AL PANEL PRINCIPAL</a>
                    </div></body></html>"""
                with open(os.path.join(ruta_base, n_m_act, n_s, "reporte.html"), "w", encoding="utf-8") as f: f.write(html_final)

                # --- INTEGRACIÓN SOLO_MES.HTML ---
                html_solo_mes = f"""<html><head><meta charset='UTF-8'></head><body>
                    <table>
                        <tr id='fila-datos'>
                            <td class='mes-nombre'>{n_m_act}</td>
                            <td class='total-incidencias'>{t_act}</td>
                            <td class='calificacion-final'>{nota_f}%</td>
                        </tr>
                    </table>
                </body></html>"""
                with open(os.path.join(ruta_base, n_m_act, n_s, "solo_mes.html"), "w", encoding="utf-8") as f: f.write(html_solo_mes)

        print("\n✅ Reporte actualizado. Recuerda renombrar tu imagen a logo.png en la carpeta RECURSOS."); input()
    except Exception as e: print(f"\n❌ ERROR: {e}"); input()

if __name__ == "__main__":
    generar_reporte_v30_final()