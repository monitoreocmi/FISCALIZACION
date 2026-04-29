import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import sys
import json
from openpyxl import load_workbook
import warnings

# Silenciar advertencias de Excel
warnings.filterwarnings("ignore", category=UserWarning)

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

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
    if indices['foto'] == -1: indices['foto'] = 10 
    return indices

def generar_reporte_cobros_v4_botones():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: GENERADOR CON NAVEGACIÓN DOBLE <<<")
        print("="*60 + "\n")
        
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        archivos = [f for f in filedialog.askopenfilenames(title="Seleccionar archivos Excel") if not os.path.basename(f).startswith('~$')]
        
        if not archivos: return

        datos_finales = []
        columnas_reales = []

        for f in archivos:
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            headers = [str(cell.value) if cell.value else f"Col_{i}" for i, cell in enumerate(ws[1])]
            if not columnas_reales: columnas_reales = headers[:10]
            idx = obtener_indices_flexibles(headers)
            
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                row_vals = [cell.value for cell in row]
                if not row_vals or len(row_vals) < 2: continue 
                
                try: color = str(row[idx['monto']].fill.start_color.index).upper()
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
                        'FILA': row_vals[:10]
                    })

        df = pd.DataFrame(datos_finales)
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))

        estilo_css = """<style>
            body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #333; padding: 20px; text-align: center; }
            .header-logos { display: flex; justify-content: space-between; align-items: center; padding: 10px 40px; background: white; margin-bottom: 20px; border-radius: 10px; border-bottom: 4px solid #F9D908; }
            .logo-header { width: 110px; }
            h1 { color: #002060; margin: 0; font-size: 20px; text-transform: uppercase; font-weight: 900; }
            .blue-box-container { background: #002060; padding: 30px; border-radius: 15px; max-width: 98%; margin: 20px auto; border: 3px solid #F9D908; color: white; }
            .table-responsive { background: white; border-radius: 8px; overflow-x: auto; color: #333; margin-top: 20px;}
            table { width: 100%; border-collapse: collapse; min-width: 1000px;}
            th { background: #001a4d; color: #F9D908; padding: 12px; font-size: 10px; border-bottom: 2px solid #F9D908;}
            td { padding: 10px; border-bottom: 1px solid #eee; font-size: 10px; font-weight: bold; text-align: left;}
            .btn-group { margin-top: 25px; display: flex; justify-content: center; gap: 15px; }
            .btn { display: inline-block; padding: 12px 25px; background: #002060; color: white !important; text-decoration: none; font-weight: bold; border-radius: 8px; border: 2px solid #F9D908; text-transform: uppercase; font-size: 12px; }
            .btn:hover { background: #F9D908; color: #002060 !important; }
            .link-foto { color: #002060; text-decoration: underline; font-weight: 800; cursor: pointer; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); align-items: center; justify-content: center; }
            .modal-content { max-width: 90%; max-height: 90%; border: 5px solid #F9D908; }
        </style>"""

        script_modal = """<div id="myModal" class="modal" onclick="this.style.display='none'"><img class="modal-content" id="imgModal"></div>
        <script>function openModal(src) { document.getElementById('myModal').style.display = "flex"; document.getElementById('imgModal').src = src; }</script>"""

        cabeceras_html = "".join([f"<th>{col}</th>" for col in columnas_reales]) + "<th style='text-align:right;'>MONTO</th>"

        for periodo in sorted(df['PERIODO'].unique()):
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            
            for suc in sorted(df_p['SUCURSAL'].unique()):
                p_suc = os.path.join(ruta_base, n_m, suc)
                os.makedirs(p_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]
                
                for est_key, file_name, titulo_web in [('COBRADO','cobrado.html','DETALLE COBRADO'),('RECUPERADO','recuperado.html','PÉRDIDA MITIGADA'),('TODO','todo_detallado.html','DETALLE COMPLETO')]:
                    df_view = df_s if est_key == 'TODO' else df_s[df_s['ESTATUS'] == est_key]
                    filas = ""
                    for _, r in df_view.iterrows():
                        m_txt = f"${r['MONTO_CALC']:,.2f}"
                        ruta_foto = f"../../FACTURAS/{n_m}/{suc}/{r['FOTO_LINK']}.jpeg"
                        
                        if r['FOTO_LINK'] and r['FOTO_LINK'] != "None":
                            td_monto = f"<td style='text-align:right;'><span class='link-foto' onclick='openModal(\"{ruta_foto}\")'>{m_txt}</span></td>"
                        else:
                            td_monto = f"<td style='text-align:right;'>{m_txt}</td>"
                        
                        tds = "".join([f"<td>{str(v) if v is not None else ''}</td>" for v in r['FILA']])
                        filas += f"<tr>{tds}{td_monto}</tr>"

                    # Generación del HTML con los dos botones solicitados
                    with open(os.path.join(p_suc, file_name), "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body>"
                                f"<div class='header-logos'><img src='../../RECURSOS/LOGO.PNG' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='../../RECURSOS/LOGO.PNG' class='logo-header'></div>"
                                f"<div class='blue-box-container'><h2>{titulo_web}</h2><h3>{suc} - {n_m}</h3>"
                                f"<div class='table-responsive'><table><thead><tr>{cabeceras_html}</tr></thead><tbody>{filas}</tbody></table></div>"
                                f"<div class='btn-group'>"
                                f"<a href='cobros_detalles.html' class='btn'>Volver al Resumen</a>"
                                f"<a href='../../index.html' class='btn'>Menú Principal</a>"
                                f"</div></div>{script_modal}</body></html>")

        print("\n✅ Botones agregados correctamente."); input()
    except Exception as e: print(f"❌ Error: {e}"); input()

if __name__ == "__main__":
    generar_reporte_cobros_v4_botones()