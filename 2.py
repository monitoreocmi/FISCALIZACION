import pandas as pd
import os
import sys
import json
from openpyxl import load_workbook
import warnings
import re
import glob

# Silenciar advertencias de validación de Excel
warnings.filterwarnings("ignore", category=UserWarning)

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

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
    return indices

def generar_reporte_cobros_final():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: GENERADOR DE COBROS <<<")
        print("="*60)
        
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_base, "cuadros")

        archivos = [os.path.join(root, f) for root, dirs, files in os.walk(ruta_cuadros) 
                    for f in files if f.endswith(".xlsx") and not f.startswith("~$")]

        datos_finales = []
        MAX_COLS_MOSTRAR = 12 

        for f in archivos:
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            headers = [str(cell.value) if cell.value else f"Col_{i}" for i, cell in enumerate(ws[1])]
            
            idx = obtener_indices_flexibles(headers)
            cabeceras_reporte = headers[:MAX_COLS_MOSTRAR]
            
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
                
                foto_raw = str(row_vals[idx['foto']]).strip() if idx['foto'] != -1 and len(row_vals) > idx['foto'] else ""
                if estatus == "OTRO" and foto_raw not in ["", "None", "nan", "SIN FOTO"]:
                    estatus = "COBRADO"

                if estatus != "OTRO":
                    fecha_val = pd.to_datetime(row_vals[idx['fecha']], errors='coerce')
                    if pd.isna(fecha_val): continue
                    
                    # Limpiamos el nombre base quitando extensiones conocidas
                    nombre_base = re.sub(r'\.(jpg|jpeg|jfif|png|0)$', '', foto_raw, flags=re.I)

                    datos_finales.append({
                        'SUCURSAL': str(row_vals[idx['sucursal']]).strip().upper(),
                        'MES': MESES_ES.get(fecha_val.month, "VARIOS"),
                        'PERIODO': fecha_val.to_period('M'),
                        'MONTO_CALC': limpiar_monto(row_vals[idx['monto']]),
                        'FOTO_BASE': nombre_base,
                        'ESTATUS': estatus, 
                        'FILA': row_vals, 
                        'IDX_MONTO': idx['monto'],
                        'IDX_FOTO': idx['foto'],
                        'HEADERS': cabeceras_reporte
                    })

        if not datos_finales:
            print("\n❌ No se encontraron datos."); return

        df = pd.DataFrame(datos_finales)
        totales_globales = {}

        estilo_css = """<style>
            body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #333; padding: 10px; text-align: center; margin: 0; }
            .header-logos { display: flex; justify-content: space-between; align-items: center; padding: 10px 20px; background: white; margin-bottom: 20px; border-bottom: 4px solid #F9D908; }
            .logo-header { height: 50px; }
            h1 { color: #002060; margin: 0; font-size: 16px; text-transform: uppercase; font-weight: 900; flex-grow: 1; }
            .resumen-grid { display: flex; justify-content: center; gap: 15px; margin: 20px 0; flex-wrap: wrap; }
            .card-resumen { background: white; padding: 20px; border-radius: 12px; text-decoration: none; width: 260px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-bottom: 6px solid #ccc; color: inherit; transition: 0.3s; }
            .card-resumen .monto { font-size: 24px; font-weight: 900; color: #002060; margin: 10px 0; }
            .cobrado { border-color: #27ae60; } .recuperado { border-color: #f1c40f; } .no-pagado { border-color: #ed1c24; }
            .blue-box-container { background: #002060; padding: 15px; border-radius: 12px; width: 98%; margin: 10px auto; border: 2px solid #F9D908; color: white; box-sizing: border-box; }
            .table-responsive { background: white; border-radius: 8px; overflow-x: auto; color: #333; margin-top: 15px; }
            table { width: 100%; border-collapse: collapse; min-width: 800px; }
            th { background: #001a4d; color: #F9D908; padding: 10px; font-size: 9px; text-transform: uppercase; border-bottom: 2px solid #F9D908; }
            td { padding: 8px; border-bottom: 1px solid #eee; font-size: 9px; font-weight: bold; text-align: left; }
            .btn { padding: 10px 18px; background: #002060; color: white !important; text-decoration: none; font-weight: bold; border-radius: 6px; border: 2px solid #F9D908; display: inline-block; margin: 5px; font-size: 11px; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); align-items: center; justify-content: center; }
            .modal-content { max-width: 95%; max-height: 95%; border: 3px solid #F9D908; }
        </style>"""

        script_modal = """<div id="myModal" class="modal" onclick="this.style.display='none'"><img class="modal-content" id="imgModal"></div>
        <script>
        function openModal(nombreBase, mes, suc) {
            const extensiones = ['jpg', 'jpeg', 'jfif', 'png', 'JPG', 'JPEG', 'JFIF', 'PNG'];
            const modal = document.getElementById('myModal');
            const img = document.getElementById('imgModal');
            
            let index = 0;
            function intentarCargar() {
                if (index >= extensiones.length) { alert("No se encontró el archivo de imagen."); return; }
                const url = `../../FACTURAS/${mes}/${suc}/${nombreBase}.${extensiones[index]}`;
                const tempImg = new Image();
                tempImg.onload = () => { img.src = url; modal.style.display = "flex"; };
                tempImg.onerror = () => { index++; intentarCargar(); };
                tempImg.src = url;
            }
            intentarCargar();
        }
        </script>"""

        for periodo in sorted(df['PERIODO'].unique()):
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            totales_globales[n_m] = {
                "TOTAL_COBRADO": float(df_p[df_p['ESTATUS'] == 'COBRADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_PERDIDA_PATRIMONIO": float(df_p[df_p['ESTATUS'] == 'RECUPERADO']['MONTO_CALC'].sum() or 0),
                "TOTAL_NO_COBRADO": float(df_p[df_p['ESTATUS'] == 'NO_PAGADO']['MONTO_CALC'].sum() or 0)
            }

            for suc in sorted(df_p['SUCURSAL'].unique()):
                p_suc = os.path.join(ruta_base, n_m, suc); os.makedirs(p_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]

                with open(os.path.join(p_suc, "cobros_detalles.html"), "w", encoding="utf-8") as f:
                    t_c, t_r, t_n = df_s[df_s['ESTATUS']=='COBRADO']['MONTO_CALC'].sum() or 0, df_s[df_s['ESTATUS']=='RECUPERADO']['MONTO_CALC'].sum() or 0, df_s[df_s['ESTATUS']=='NO_PAGADO']['MONTO_CALC'].sum() or 0
                    f.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>SISTEMA LUXOR</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div><h2>{suc} | {n_m}</h2><div class='resumen-grid'><a href='cobrado.html' class='card-resumen cobrado'><h3>Cobrado</h3><div class='monto'>${t_c:,.2f}</div></a><a href='recuperado.html' class='card-resumen recuperado'><h3>Pérdida mitigada</h3><div class='monto'>${t_r:,.2f}</div></a><a href='no_pagado.html' class='card-resumen no-pagado'><h3>No Pagado</h3><div class='monto'>${t_n:,.2f}</div></a></div><a href='todo_detallado.html' class='btn'>LISTADO COMPLETO</a><a href='../../index.html?tab=cobs' class='btn'>INICIO</a></body></html>")

                for est_key, file_name, titulo in [('COBRADO', 'cobrado.html', 'DETALLE COBRADO'), ('RECUPERADO', 'recuperado.html', 'PÉRDIDA MITIGADA'), ('NO_PAGADO', 'no_pagado.html', 'DETALLE NO PAGADO'), ('TODO', 'todo_detallado.html', 'DETALLE COMPLETO')]:
                    df_view = df_s if est_key == 'TODO' else df_s[df_s['ESTATUS'] == est_key]
                    filas = ""
                    cabeceras_finales = df_view['HEADERS'].iloc[0] if not df_view.empty else []
                    
                    for _, r in df_view.iterrows():
                        tds = ""
                        for i in range(MAX_COLS_MOSTRAR):
                            val = r['FILA'][i] if i < len(r['FILA']) else ""
                            if i == r['IDX_MONTO'] and r['FOTO_BASE'] not in ["", "None", "nan", "SIN FOTO"]:
                                tds += f"<td><span style='cursor:pointer; color:#002060; text-decoration:underline;' onclick='openModal(\"{r['FOTO_BASE']}\", \"{n_m}\", \"{suc}\")'>${limpiar_monto(val):,.2f}</span></td>"
                            elif i == r['IDX_MONTO']:
                                tds += f"<td>${limpiar_monto(val):,.2f}</td>"
                            else:
                                tds += f"<td>{str(val) if val is not None else ''}</td>"
                        filas += f"<tr>{tds}</tr>"

                    with open(os.path.join(p_suc, file_name), "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>GERENCIA DE FISCALIZACIÓN</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div><div class='blue-box-container'><h2>{titulo}</h2><div class='table-responsive'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in cabeceras_finales])}</tr></thead><tbody>{filas}</tbody></table></div><a href='cobros_detalles.html' class='btn'>VOLVER</a></div>{script_modal}</body></html>")

        with open(os.path.join(ruta_base, "TOTALES_GLOBALES_COBROS.json"), "w", encoding="utf-8") as f:
            json.dump(totales_globales, f, indent=4, ensure_ascii=False)
        print("\n✅ PROCESO FINALIZADO.")
    except Exception as e: print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    generar_reporte_cobros_final()