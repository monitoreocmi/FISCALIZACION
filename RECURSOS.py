import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import sys
import re

# Configuración de salida para evitar errores de caracteres
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

RUTA_CSS_EXTERNO = "../../RECURSOS/estilos.css"
RUTA_LOGO = "../../RECURSOS/LOGO.PNG"

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[^\w\s-]', '', str(nombre)).strip().replace(' ', '_')

def generar_reporte_integrado_v28():
    try:
        print("\n" + "="*60)
        print(">>> GENERADOR DE INCIDENCIAS V28 - INTEGRACIÓN FINAL <<<")
        print("="*60 + "\n")
        
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        archivos = filedialog.askopenfilenames(title="Seleccionar archivos Excel")
        if not archivos: 
            print("❌ No se seleccionaron archivos."); return

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

        resumen_cmd = {}

        # FASE 1: GENERACIÓN DE DETALLES (HTMLs de incidencias y responsables)
        for p in df_master['PERIODO'].unique():
            n_m = MESES_ES[p.month]
            df_p = df_master[df_master['PERIODO'] == p]
            for s in df_p['SUCURSAL'].dropna().unique():
                n_s = str(s).strip().upper()
                p_sucursal = os.path.join(ruta_base, n_m, n_s)
                os.makedirs(p_sucursal, exist_ok=True)
                df_s_p = df_p[df_p['SUCURSAL'] == s]
                cols_mostrar = [c for c in df_s_p.columns if c not in ['INC_LIMPIA', 'PERIODO', 'NOMBRE_AUX'] and not c.startswith('UNNAMED')]

                def escribir_detalle(ruta, titulo, df_f):
                    df_f = df_f.fillna("")
                    cuerpo = "".join([f"<tr><td>{i+1}</td>{''.join([f'<td>{r[c]}</td>' for c in cols_mostrar])}</tr>" for i, (_, r) in enumerate(df_f.iterrows())])
                    with open(ruta, "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'><link rel='stylesheet' href='{RUTA_CSS_EXTERNO}'></head><body><div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext logo-left'><div class='title-header'><h3>{titulo}</h3></div><img src='{RUTA_LOGO}' class='logo-ext logo-right'></div><div class='main-container'><table><thead><tr><th>#</th>{''.join([f'<th>{c}</th>' for c in cols_mostrar])}</tr></thead><tbody>{cuerpo}</tbody></table><br><a href='#' onclick='history.back()' class='btn-volver'>VOLVER</a></div></body></html>")

                if 'NOMBRE_AUX' in df_s_p.columns:
                    for resp in df_s_p['NOMBRE_AUX'].dropna().unique():
                        escribir_detalle(os.path.join(p_sucursal, f"RESP_{limpiar_nombre_archivo(resp)}.html"), f"RESPONSABLE: {resp}", df_s_p[df_s_p['NOMBRE_AUX'] == resp])
                escribir_detalle(os.path.join(p_sucursal, "TODAS.html"), f"TODAS LAS INCIDENCIAS - {n_m}", df_s_p)
                for g in grupos_defs:
                    for inc in g["items"]:
                        sub = df_s_p[df_s_p['INC_LIMPIA'] == inc.upper().replace('.', '')]
                        if not sub.empty: escribir_detalle(os.path.join(p_sucursal, f"{limpiar_nombre_archivo(inc)}.html"), inc, sub)

        # FASE 2: GENERACIÓN DE REPORTES PRINCIPALES
        periodos_cron = sorted(df_master['PERIODO'].unique())
        for p_act in periodos_cron:
            n_m_act = MESES_ES[p_act.month]
            resumen_cmd[n_m_act] = []
            p_ant = p_act - 1
            n_m_ant = MESES_ES[p_ant.month] if p_ant.month in MESES_ES else "ANT."
            df_m_act = df_master[df_master['PERIODO'] == p_act]
            
            for suc in sorted(df_m_act['SUCURSAL'].dropna().unique()):
                n_s = str(suc).strip().upper()
                resumen_cmd[n_m_act].append(n_s)
                df_suc_act = df_m_act[df_m_act['SUCURSAL'] == suc]
                filas_html, suma_impacto, t_act, t_ant = "", 0, 0, 0
                
                for grupo in grupos_defs:
                    g_act_grupo, g_ant_grupo = 0, 0
                    num_items = len(grupo["items"])
                    for i, inc in enumerate(grupo["items"]):
                        busq = inc.upper().replace('.', '')
                        c_act = int(len(df_suc_act[df_suc_act['INC_LIMPIA']==busq]))
                        c_ant = int(len(df_master[(df_master['SUCURSAL']==suc) & (df_master['PERIODO']==p_ant) & (df_master['INC_LIMPIA']==busq)]))
                        g_act_grupo += c_act; g_ant_grupo += c_ant; t_act += c_act; t_ant += c_ant
                        
                        v_ant = f"0" if c_ant == 0 else f"<a href='../../{n_m_ant}/{n_s}/{limpiar_nombre_archivo(inc)}.html'>{c_ant}</a>"
                        v_act = f"0" if c_act == 0 else f"<a href='{limpiar_nombre_archivo(inc)}.html'>{c_act}</a>"
                        
                        # Atributo de clase para aplicar borde en la última fila del grupo
                        clase_fila = "class='linea-grupo'" if i == num_items - 1 else ""
                        td_porc = f"<td rowspan='{num_items}' style='vertical-align:middle; font-weight:bold;'>{grupo['porc']}%</td>" if i == 0 else ""
                        
                        filas_html += f"""<tr style='background-color:{grupo['color']};' {clase_fila}>
                            <td style='text-align:left;'>{inc}.</td>
                            <td>{grupo['tipo']}</td>
                            <td>{v_ant}</td>
                            <td>{v_act}</td>
                            {td_porc}
                        </tr>"""
                    
                    if g_act_grupo > g_ant_grupo: suma_impacto += grupo["porc"]

                nota_f = max(0, 100 - suma_impacto)
                # Links de totales (Mes Anterior en Blanco)
                link_t_ant = f"<a href='../../{n_m_ant}/{n_s}/TODAS.html' style='color:white;'>{t_ant}</a>" if t_ant > 0 else "0"
                link_t_act = f"<a href='TODAS.html' style='color:white;'>{t_act}</a>" if t_act > 0 else "0"

                filas_resp = ""
                if 'NOMBRE_AUX' in df_suc_act.columns:
                    for name, cant in df_suc_act['NOMBRE_AUX'].value_counts().head(3).items():
                        url = f"RESP_{limpiar_nombre_archivo(name)}.html"
                        filas_resp += f"<tr><td style='text-align:left;'>{name}</td><td colspan='4'><a href='{url}' style='color:#ed1c24; font-weight:bold;'>{cant}</a></td></tr>"

                html_final = f"""<html><head><meta charset='UTF-8'><link rel='stylesheet' href='{RUTA_CSS_EXTERNO}'>
                    <style>
                        table {{ border-collapse: collapse; width: 100%; }}
                        .linea-grupo td {{ border-bottom: 3px solid #333 !important; }}
                    </style></head><body>
                    <div class='top-bar'><img src='{RUTA_LOGO}' class='logo-ext logo-left'><div class='title-header'><h1>{n_s} - {n_m_act}</h1></div><img src='{RUTA_LOGO}' class='logo-ext logo-right'></div>
                    <div class='main-container'>
                        <table>
                            <thead><tr><th>INCIDENCIA</th><th>TIPO</th><th>{n_m_ant}</th><th>{n_m_act}</th><th>APROBATORIO 75%</th></tr></thead>
                            <tbody>
                                {filas_html}
                                <tr class='row-total'><td style='text-align:right;' colspan='2'>TOTAL GENERAL / CALIFICACIÓN</td><td>{link_t_ant}</td><td>{link_t_act}</td><td style='background:{'#ed1c24' if nota_f < 75 else '#27ae60'}; color:white;'>{nota_f}%</td></tr>
                                <tr><td colspan='5' style='font-weight:900; background:#f0f0f0; padding:10px; color:#002060; border-top: 3px solid #002060;'>TOP 3 RESPONSABLES CON MAYOR INCIDENCIA</td></tr>
                                {filas_resp}
                            </tbody>
                        </table><br><a href='../../index.html' class='btn-volver'>VOLVER AL PANEL PRINCIPAL</a>
                    </div></body></html>"""
                
                f_p = os.path.join(ruta_base, n_m_act, n_s, "reporte.html")
                with open(f_p, "w", encoding="utf-8") as f: f.write(html_final)

        # RESUMEN CMD DETALLADO
        print("\n🏁 RESUMEN DE PROCESAMIENTO:")
        print("-" * 35)
        for mes, sucs in resumen_cmd.items():
            print(f"📅 MES: {mes}")
            for i, s in enumerate(sucs, 1):
                print(f"   {i}. {s}")
            print("-" * 35)

        print("\n✅ PROCESO COMPLETADO EXITOSAMENTE.")
        input("\nPresiona ENTER para cerrar...")
        
    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO: {e}"); input("\nPresiona ENTER para salir...")

if __name__ == "__main__":
    generar_reporte_integrado_v28()