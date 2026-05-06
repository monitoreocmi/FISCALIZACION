import pandas as pd
import os
import sys
import json
from openpyxl import load_workbook
import warnings
import re
import glob
import threading
import time

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
        if 'MONTO $' in h_up: indices['monto'] = i
        elif 'MONTO' in h_up and indices['monto'] == -1: indices['monto'] = i
        if 'FECHA' in h_up: indices['fecha'] = i
        if 'F COBRADA' in h_up: indices['foto'] = i
    
    if indices['monto'] == -1: indices['monto'] = 9 
    if indices['foto'] == -1: indices['foto'] = 10 
    return indices

def generar_reporte_cobros_final():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: GENERADOR DE COBROS (COLUMNA J) <<<")
        print("="*60)
        
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_base, "cuadros")

        if not os.path.isdir(ruta_cuadros):
            print(f"❌ ERROR: No se encuentra la carpeta 'cuadros' en:\n{ruta_cuadros}")
            return

        archivos = []
        for root, dirs, files in os.walk(ruta_cuadros):
            for file in files:
                if file.endswith(".xlsx") and not file.startswith("~$"):
                    archivos.append(os.path.join(root, file))
        
        if not archivos:
            print(f"⚠️ No hay archivos Excel válidos en: {ruta_cuadros}")
            return

        print(f"✅ Se encontraron {len(archivos)} archivos. Procesando...\n")

        datos_finales = []
        columnas_reales = [] 

        for f in archivos:
            rel_path = os.path.relpath(f, ruta_cuadros)
            print(f"📖 Leyendo: {rel_path}")
            wb = load_workbook(f, data_only=False)
            ws = wb.active
            headers = [str(cell.value) if cell.value else f"Col_{i}" for i, cell in enumerate(ws[1])]
            
            if not columnas_reales:
                columnas_reales = headers[:11]
            
            idx = obtener_indices_flexibles(headers)
            print(f"   -> Mapeo: Sucursal Col {idx['sucursal']}, Monto Col {idx['monto']}, Foto Col {idx['foto']}")
            
            count_rows = 0
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                row_vals = [cell.value for cell in row]
                if not row_vals or len(row_vals) < 2: continue 
                
                estatus = "OTRO"
                try:
                    celda_monto = row[idx['monto']]
                    color = str(celda_monto.fill.start_color.index).upper()
                    
                    if color in ['FF00B050', 'FF92D050', '00FF00', 'FF00FF00', 'FFC6EFCE', '13', 'FF548235']: 
                        estatus = "COBRADO"
                    elif color in ['FFFFFF00', 'FFFFFFE1', 'FFFFEB9C', '17', 'FFFFD966', 'FFFFC000']: 
                        estatus = "RECUPERADO"
                    elif color in ['FF0070C0', 'FF00B0F0', 'FFCCE5FF', '24', '30', 'FF3D85C6']:
                        estatus = "EXCEDENTE"
                except: 
                    pass

                foto_raw = str(row_vals[idx['foto']]).strip() if len(row_vals) > idx['foto'] and row_vals[idx['foto']] else ""
                if estatus == "OTRO" and foto_raw != "" and foto_raw.upper() not in ["NONE", "NAN", "SIN FOTO"]:
                    estatus = "COBRADO"

                if estatus != "OTRO":
                    fecha_val = pd.to_datetime(row_vals[idx['fecha']], errors='coerce')
                    if pd.isna(fecha_val): continue
                    
                    nombre_foto = foto_raw.replace(".0", "") if foto_raw.endswith(".0") else foto_raw
                    sucursal_nombre = str(row_vals[idx['sucursal']]).strip().upper() if idx['sucursal'] != -1 else "GENERAL"

                    datos_finales.append({
                        'SUCURSAL': sucursal_nombre,
                        'MES': MESES_ES.get(fecha_val.month, "VARIOS"),
                        'PERIODO': fecha_val.to_period('M'),
                        'MONTO_CALC': limpiar_monto(row_vals[idx['monto']]),
                        'FOTO_LINK': nombre_foto,
                        'ESTATUS': estatus, 
                        'FILA': row_vals[:11],
                        'IDX_MONTO': idx['monto']
                    })
                    count_rows += 1
            print(f"   -> ✓ {count_rows} filas válidas extraídas.")

        if not datos_finales:
            print("\n❌ No se encontraron datos válidos (coloreados o con foto).")
            return

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
            .cobrado { border-color: #27ae60; } .recuperado { border-color: #f1c40f; } .excedente { border-color: #0070c0; }
            .blue-box-container { background: #002060; padding: 15px; border-radius: 12px; width: 98%; margin: 10px auto; border: 2px solid #F9D908; color: white; box-sizing: border-box; }
            .table-responsive { background: white; border-radius: 8px; overflow-x: auto; color: #333; margin-top: 15px; }
            table { width: 100%; border-collapse: collapse; min-width: 800px; }
            th { background: #001a4d; color: #F9D908; padding: 10px; font-size: 9px; text-transform: uppercase; border-bottom: 2px solid #F9D908; }
            td { padding: 8px; border-bottom: 1px solid #eee; font-size: 9px; font-weight: bold; text-align: left; }
            .btn-group { margin-top: 20px; display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; }
            .btn { padding: 10px 18px; background: #002060; color: white !important; text-decoration: none; font-weight: bold; border-radius: 6px; border: 2px solid #F9D908; text-transform: uppercase; font-size: 11px; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); align-items: center; justify-content: center; }
            .modal-content { max-width: 95%; max-height: 95%; border: 3px solid #F9D908; }
        </style>"""

        script_modal = """<div id="myModal" class="modal" onclick="this.style.display='none'"><img class="modal-content" id="imgModal"></div>
        <script>function openModal(src) { document.getElementById('myModal').style.display = "flex"; document.getElementById('imgModal').src = src; }</script>"""

        print("\n📂 Generando estructuras de carpetas y HTMLs...")
        for periodo in sorted(df['PERIODO'].unique()):
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            print(f"📍 Procesando Mes: {n_m}")
            
            totales_globales[n_m] = {
                "TOTAL_COBRADO": float(df_p[df_p['ESTATUS'] == 'COBRADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_PERDIDA_PATRIMONIO": float(df_p[df_p['ESTATUS'] == 'RECUPERADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_EXCEDENTE": float(df_p[df_p['ESTATUS'] == 'EXCEDENTE']['MONTO_CALC'].sum() or 0)
            }

            sucursales = sorted(df_p['SUCURSAL'].unique())
            for suc in sucursales:
                p_suc = os.path.join(ruta_base, n_m, suc)
                os.makedirs(p_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]
                
                t_c = df_s[df_s['ESTATUS']=='COBRADO']['MONTO_CALC'].sum() or 0
                t_r = df_s[df_s['ESTATUS']=='RECUPERADO']['MONTO_CALC'].sum() or 0
                t_e = df_s[df_s['ESTATUS']=='EXCEDENTE']['MONTO_CALC'].sum() or 0
                
                html_menu = f"""<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{estilo_css}</head><body>
                <div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div>
                <h2>{suc} | {n_m}</h2>
                <div class='resumen-grid'>
                    <a href='cobrado.html' class='card-resumen cobrado'><h3>Cobrado</h3><div class='monto'>${t_c:,.2f}</div></a>
                    <a href='recuperado.html' class='card-resumen recuperado'><h3>Pérdida mitigada y patrimonio</h3><div class='monto'>${t_r:,.2f}</div></a>
                    <a href='excedente.html' class='card-resumen excedente'><h3>Excedentes</h3><div class='monto'>${t_e:,.2f}</div></a>
                </div>
                <div class='btn-group'>
                    <a href='todo_detallado.html' class='btn'>LISTADO COMPLETO</a>
                    <a href='../../index.html?tab=cobs' class='btn'>MENÚ PRINCIPAL</a>
                </div>
                </body></html>"""
                with open(os.path.join(p_suc, "cobros_detalles.html"), "w", encoding="utf-8") as f_out: f_out.write(html_menu)

                for est_key, file_name, titulo_web in [('COBRADO', 'cobrado.html', 'DETALLE COBRADO'), ('RECUPERADO', 'recuperado.html', 'PÉRDIDA MITIGADA'), ('EXCEDENTE', 'excedente.html', 'DETALLE EXCEDENTES'), ('TODO', 'todo_detallado.html', 'DETALLE COMPLETO')]:
                    df_view = df_s if est_key == 'TODO' else df_s[df_s['ESTATUS'] == est_key]
                    filas = ""
                    for _, r in df_view.iterrows():
                        idx_monto = r['IDX_MONTO']
                        tds = ""
                        for i, v in enumerate(r['FILA']):
                            if i == idx_monto:
                                if r['FOTO_LINK'] and r['FOTO_LINK'] != "":
                                    tds += f"<td><span style='cursor:pointer; color:#002060; text-decoration:underline;' onclick='openModal(\"../../FACTURAS/{n_m}/{suc}/{r['FOTO_LINK']}.jpeg\")'>${r['MONTO_CALC']:,.2f}</span></td>"
                                else:
                                    tds += f"<td>${r['MONTO_CALC']:,.2f}</td>"
                            else:
                                tds += f"<td>{str(v) if v is not None else ''}</td>"
                        filas += f"<tr>{tds}</tr>"

                    cabeceras_html = "".join([f"<th>{col}</th>" for col in columnas_reales])
                    with open(os.path.join(p_suc, file_name), "w", encoding="utf-8") as f_out:
                        f_out.write(f"<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>{estilo_css}</head><body>"
                                f"<div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div>"
                                f"<div class='blue-box-container'><h2>{titulo_web}</h2><h3>{suc} - {n_m}</h3>"
                                f"<div class='table-responsive'><table><thead><tr>{cabeceras_html}</tr></thead><tbody>{filas}</tbody></table></div>"
                                f"<div class='btn-group'>"
                                f"<a href='cobros_detalles.html' class='btn'>VOLVER AL RESUMEN</a>"
                                f"<a href='../../index.html?tab=cobs' class='btn'>MENÚ PRINCIPAL</a>"
                                f"</div></div>{script_modal}</body></html>")
            print(f"   ✓ {suc}: Archivos de detalle creados.")

        with open(os.path.join(ruta_base, "TOTALES_GLOBALES_COBROS.json"), "w", encoding="utf-8") as f_json:
            json.dump(totales_globales, f_json, indent=4, ensure_ascii=False)

        print("\n" + "="*60)
        print("✅ PROCESO FINALIZADO CON ÉXITO.")
        print("="*60)

    except Exception as e: 
        print(f"\n❌ Error Crítico: {e}")

    # LÓGICA DE CIERRE AUTOMÁTICO EN 10 SEGUNDOS
    print("\n" + "-"*30)
    print("Presiona ENTER para salir (o espera 10 segundos)...")
    
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__":
    generar_reporte_cobros_final()