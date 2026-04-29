import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import sys
import json
from openpyxl import load_workbook
import warnings
import re

# Silenciar advertencias de validación de Excel
warnings.filterwarnings("ignore", category=UserWarning)

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# Ruta optimizada para GitHub (Case Sensitive)
RUTA_LOGO_ESTANDAR = "../../RECURSOS/logo.png"

def limpiar_monto(valor):
    if valor is None: return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    texto = str(valor).strip().replace('$', '').replace(' ', '')
    if ',' in texto and '.' in texto:
        if texto.find('.') < texto.find(','): texto = texto.replace('.', '').replace(',', '.')
        else: texto = texto.replace(',', '')
    elif ',' in texto: texto = texto.replace(',', '.')
    try: return float(texto)
    except: return 0.0

def obtener_indices_flexibles(headers):
    indices = {'sucursal': -1, 'monto': -1, 'fecha': -1, 'foto': -1}
    for i, h in enumerate(headers):
        h_up = str(h).upper() if h else ""
        if 'SUCURSAL' in h_up: indices['sucursal'] = i
        if 'MONTO' in h_up: indices['monto'] = i
        if 'FECHA' in h_up: indices['fecha'] = i
        if 'F COBRADA' in h_up: indices['foto'] = i
    if indices['foto'] == -1: indices['foto'] = 10 # Columna K por defecto
    return indices

def generar_reporte_cobros_final():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: GENERADOR DE COBROS OPTIMIZADO <<<")
        print("="*60 + "\n")
        
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        archivos = filedialog.askopenfilenames(title="Seleccionar archivos Excel de Cobros")
        if not archivos: return

        datos_finales = []
        columnas_reales = [] 
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))

        for f in archivos:
            print(f"📖 Leyendo: {os.path.basename(f)}")
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            headers = [str(cell.value) if cell.value else f"Col_{i}" for i, cell in enumerate(ws[1])]
            
            if not columnas_reales:
                columnas_reales = headers[:10]
            
            idx = obtener_indices_flexibles(headers)
            
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                row_vals = [cell.value for cell in row]
                if not row_vals or len(row_vals) < 2: continue 
                
                try:
                    color = str(row[idx['monto']].fill.start_color.index).upper()
                except: color = "00000000"

                estatus = "OTRO"
                if color in ['FF00B050', 'FF92D050', '00FF00', 'FF00FF00', 'FFC6EFCE', '13']: estatus = "COBRADO"
                elif color in ['FFFF0000', 'FFC00000', 'FFFFC7CE', '10']: estatus = "NO_PAGADO"
                elif color in ['FFFFFF00', 'FFFFFFE1', 'FFFFEB9C', '17']: estatus = "RECUPERADO"
                
                if estatus != "OTRO":
                    fecha_val = pd.to_datetime(row_vals[idx['fecha']], errors='coerce')
                    if pd.isna(fecha_val): continue
                    
                    foto_raw = str(row_vals[idx['foto']]).strip() if len(row_vals) > idx['foto'] and row_vals[idx['foto']] else ""
                    nombre_foto = foto_raw.replace(".0", "") if foto_raw.endswith(".0") else foto_raw

                    datos_finales.append({
                        'SUCURSAL': str(row_vals[idx['sucursal']]).strip().upper(),
                        'MES': MESES_ES[fecha_val.month],
                        'PERIODO': fecha_val.to_period('M'),
                        'MONTO_CALC': limpiar_monto(row_vals[idx['monto']]),
                        'FOTO_LINK': nombre_foto,
                        'ESTATUS': estatus, 
                        'FILA': row_vals[:10],
                        'IDX_MONTO': idx['monto']
                    })

        if not datos_finales:
            print("❌ No se encontraron datos."); return

        df = pd.DataFrame(datos_finales)
        totales_globales = {}

        estilo_css = """<style>
            body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #333; padding: 10px; text-align: center; margin: 0; }
            .header-logos { display: flex; justify-content: space-between; align-items: center; padding: 10px 20px; background: white; margin-bottom: 20px; border-bottom: 4px solid #F9D908; }
            .logo-header { height: 50px; max-width: 100px; object-fit: contain; }
            h1 { color: #002060; margin: 0; font-size: 16px; text-transform: uppercase; font-weight: 900; flex-grow: 1; }
            .resumen-grid { display: flex; justify-content: center; gap: 15px; margin: 20px 0; flex-wrap: wrap; }
            .card-resumen { background: white; padding: 20px; border-radius: 12px; text-decoration: none; width: 260px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-bottom: 6px solid #ccc; color: inherit; transition: 0.3s; }
            .card-resumen:hover { transform: translateY(-3px); }
            .card-resumen h3 { margin: 0; color: #666; text-transform: uppercase; font-size: 10px; height: 30px; display: flex; align-items: center; justify-content: center; }
            .card-resumen .monto { font-size: 24px; font-weight: 900; color: #002060; margin: 10px 0; }
            .cobrado { border-color: #27ae60; } .recuperado { border-color: #f1c40f; } .no-pagado { border-color: #ed1c24; }
            .blue-box-container { background: #002060; padding: 15px; border-radius: 12px; width: 98%; margin: 10px auto; border: 2px solid #F9D908; color: white; box-sizing: border-box; }
            .table-responsive { background: white; border-radius: 8px; overflow-x: auto; color: #333; margin-top: 15px; }
            table { width: 100%; border-collapse: collapse; min-width: 800px; }
            th { background: #001a4d; color: #F9D908; padding: 10px; font-size: 9px; text-transform: uppercase; border-bottom: 2px solid #F9D908; }
            td { padding: 8px; border-bottom: 1px solid #eee; font-size: 9px; font-weight: bold; text-align: left; }
            .btn-group { margin-top: 20px; display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; }
            .btn { padding: 10px 18px; background: #002060; color: white !important; text-decoration: none; font-weight: bold; border-radius: 6px; border: 2px solid #F9D908; text-transform: uppercase; font-size: 11px; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); align-items: center; justify-content: center; }
            .modal-content { max-width: 95%; max-height: 95%; border: 3px solid #F9D908; }
            
            /* Ajustes Móvil y Laptop */
            @media (max-width: 768px) {
                h1 { font-size: 14px; }
                .logo-header { height: 35px; }
                .card-resumen { width: 100%; max-width: 300px; }
                .blue-box-container { width: 100%; border-radius: 0; }
                .btn { width: 100%; max-width: 250px; text-align: center; }
            }
        </style>"""

        script_modal = """<div id="myModal" class="modal" onclick="this.style.display='none'"><img class="modal-content" id="imgModal"></div>
        <script>function openModal(src) { document.getElementById('myModal').style.display = "flex"; document.getElementById('imgModal').src = src; }</script>"""

        for periodo in sorted(df['PERIODO'].unique()):
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            
            totales_globales[n_m] = {
                "TOTAL_COBRADO": float(df_p[df_p['ESTATUS'] == 'COBRADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_PERDIDA_PATRIMONIO": float(df_p[df_p['ESTATUS'] == 'RECUPERADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_NO_COBRADO": float(df_p[df_p['ESTATUS'] == 'NO_PAGADO']['MONTO_CALC'].sum() or 0)
            }

            for suc in sorted(df_p['SUCURSAL'].unique()):
                p_suc = os.path.join(ruta_base, n_m, suc)
                os.makedirs(p_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]
                
                t_c = df_s[df_s['ESTATUS']=='COBRADO']['MONTO_CALC'].sum() or 0
                t_r = df_s[df_s['ESTATUS']=='RECUPERADO']['MONTO_CALC'].sum() or 0
                t_n = df_s[df_s['ESTATUS']=='NO_PAGADO']['MONTO_CALC'].sum() or 0
                
                # HTML MENÚ
                html_menu = f"""<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{estilo_css}</head><body>
                <div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div>
                <h2>{suc} | {n_m}</h2>
                <div class='resumen-grid'>
                    <a href='cobrado.html' class='card-resumen cobrado'><h3>Cobrado</h3><div class='monto'>${t_c:,.2f}</div></a>
                    <a href='recuperado.html' class='card-resumen recuperado'><h3>Pérdida mitigada y patrimonio</h3><div class='monto'>${t_r:,.2f}</div></a>
                    <a href='no_pagado.html' class='card-resumen no-pagado'><h3>No Pagado</h3><div class='monto'>${t_n:,.2f}</div></a>
                </div>
                <div class='btn-group'>
                    <a href='todo_detallado.html' class='btn'>LISTADO COMPLETO</a>
                    <a href='../../index.html?tab=cobs' class='btn'>MENÚ PRINCIPAL</a>
                </div>
                </body></html>"""
                with open(os.path.join(p_suc, "cobros_detalles.html"), "w", encoding="utf-8") as f: f.write(html_menu)

                # REPORTES DETALLADOS
                for est_key, file_name, titulo_web in [('COBRADO', 'cobrado.html', 'DETALLE COBRADO'), ('RECUPERADO', 'recuperado.html', 'PÉRDIDA MITIGADA'), ('NO_PAGADO', 'no_pagado.html', 'DETALLE NO PAGADO'), ('TODO', 'todo_detallado.html', 'DETALLE COMPLETO')]:
                    df_view = df_s if est_key == 'TODO' else df_s[df_s['ESTATUS'] == est_key]
                    filas = ""
                    for _, r in df_view.iterrows():
                        idx_monto = r['IDX_MONTO']
                        tds = "".join([f"<td><span class='link-foto' onclick='openModal(\"../../FACTURAS/{n_m}/{suc}/{r['FOTO_LINK']}.jpeg\")'>${limpiar_monto(v):,.2f}</span></td>" if i == idx_monto and r['FOTO_LINK'] and r['FOTO_LINK'] != "None" else f"<td>${limpiar_monto(v):,.2f}</td>" if i == idx_monto else f"<td>{str(v) if v is not None else ''}</td>" for i, v in enumerate(r['FILA'])])
                        filas += f"<tr>{tds}</tr>"

                    cabeceras_html = "".join([f"<th>{col}</th>" for col in columnas_reales])
                    with open(os.path.join(p_suc, file_name), "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{estilo_css}</head><body>"
                                f"<div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div>"
                                f"<div class='blue-box-container'><h2>{titulo_web}</h2><h3>{suc} - {n_m}</h3>"
                                f"<div class='table-responsive'><table><thead><tr>{cabeceras_html}</tr></thead><tbody>{filas}</tbody></table></div>"
                                f"<div class='btn-group'>"
                                f"<a href='cobros_detalles.html' class='btn'>VOLVER AL RESUMEN</a>"
                                f"<a href='../../index.html?tab=cobs' class='btn'>MENÚ PRINCIPAL</a>"
                                f"</div></div>{script_modal}</body></html>")

        with open(os.path.join(ruta_base, "TOTALES_GLOBALES_COBROS.json"), "w", encoding="utf-8") as f:
            json.dump(totales_globales, f, indent=4, ensure_ascii=False)

        print("\n✅ Reportes generados y optimizados. Recuerda usar logo.png en RECURSOS."); input()
    except Exception as e: print(f"❌ Error: {e}"); input()

if __name__ == "__main__":
    generar_reporte_cobros_final()