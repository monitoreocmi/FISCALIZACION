import pandas as pd
import os
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
    indices = {'sucursal': -1, 'monto': 9, 'fecha': -1, 'foto': -1}
    for i, h in enumerate(headers):
        h_up = str(h).upper() if h else ""
        if 'SUCURSAL' in h_up: indices['sucursal'] = i
        if 'FECHA' in h_up: indices['fecha'] = i
        if 'F COBRADA' in h_up: indices['foto'] = i
    return indices

def generar_reporte_cobros_final():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: GENERACIÓN INTELIGENTE (JSON + HTML) <<<")
        print("="*60)
        
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_base, "cuadros")
        archivos = [os.path.join(root, f) for root, dirs, files in os.walk(ruta_cuadros) 
                    for f in files if f.endswith(".xlsx") and not f.startswith("~$")]

        datos_finales = []
        MAX_COLS_MOSTRAR = 12 

        for f in archivos:
            print(f"Procesando: {os.path.basename(f)}...")
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            headers = [str(cell.value) if cell.value else f"Col_{i}" for i, cell in enumerate(ws[1])]
            idx = obtener_indices_flexibles(headers)
            
            for row in ws.iter_rows(min_row=2):
                row_vals = [cell.value for cell in row]
                if not any(row_vals): continue
                
                try:
                    color = str(row[9].fill.start_color.index).upper()
                except: color = "00000000"

                estatus = "OTRO"
                if color in ['FF00B050', 'FF92D050', '00FF00', 'FF00FF00', 'FFC6EFCE', '13', 'FF548235']: estatus = "COBRADO"
                elif color in ['FFFF0000', 'FFC00000', 'FFFFC7CE', '10', 'FFE06666']: estatus = "NO_PAGADO"
                elif color in ['FFFFFF00', 'FFFFFFE1', 'FFFFEB9C', '17', 'FFFFD966']: estatus = "RECUPERADO"
                elif color in ['FF0070C0', 'FF00B0F0', 'FFCCE5FF', '24', '30', 'FF3D85C6']: estatus = "EXCEDENTE"
                
                foto_raw = ""
                if idx['foto'] != -1 and idx['foto'] < len(row_vals):
                    val_foto = row_vals[idx['foto']]
                    foto_raw = str(val_foto).strip() if val_foto is not None else ""

                if estatus == "OTRO" and foto_raw not in ["", "None", "nan", "SIN FOTO"]:
                    estatus = "COBRADO"

                if estatus != "OTRO":
                    fecha_val = pd.to_datetime(row_vals[idx['fecha']], errors='coerce')
                    mes_nombre = MESES_ES.get(fecha_val.month, "VARIOS") if not pd.isna(fecha_val) else "VARIOS"
                    periodo_val = fecha_val.to_period('M') if not pd.isna(fecha_val) else pd.Period('2026-01', freq='M')

                    datos_finales.append({
                        'SUCURSAL': str(row_vals[idx['sucursal']]).strip().upper() if idx['sucursal'] != -1 else "GENERAL",
                        'MES': mes_nombre, 'PERIODO': periodo_val,
                        'MONTO_CALC': limpiar_monto(row_vals[9]),
                        'FOTO_BASE': re.sub(r'\.(jpg|jpeg|jfif|png|0)$', '', foto_raw, flags=re.I),
                        'ESTATUS': estatus, 'FILA': row_vals, 'HEADERS': headers[:MAX_COLS_MOSTRAR]
                    })
            wb.close()

        if not datos_finales:
            print("\n❌ No se encontraron registros detectables."); return

        df = pd.DataFrame(datos_finales)
        
        # --- 1. GENERACIÓN DE JSONS ---
        totales_globales = {}
        totales_sucursales = {}
        for periodo in sorted(df['PERIODO'].unique()):
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            
            totales_globales[n_m] = {
                "TOTAL_COBRADO": float(df_p[df_p['ESTATUS']=="COBRADO"]['MONTO_CALC'].sum()),
                "TOTAL_PERDIDA_PATRIMONIO": float(df_p[df_p['ESTATUS']=="RECUPERADO"]['MONTO_CALC'].sum()),
                "TOTAL_NO_COBRADO": float(df_p[df_p['ESTATUS']=="NO_PAGADO"]['MONTO_CALC'].sum()),
                "TOTAL_EXCEDENTE": float(df_p[df_p['ESTATUS']=="EXCEDENTE"]['MONTO_CALC'].sum())
            }

            for suc in df_p['SUCURSAL'].unique():
                df_s = df_p[df_p['SUCURSAL'] == suc]
                s_key = f"{suc} ({n_m})"
                totales_sucursales[s_key] = {
                    "COBRADO": float(df_s[df_s['ESTATUS']=="COBRADO"]['MONTO_CALC'].sum()),
                    "PERDIDA_PATRIMONIO": float(df_s[df_s['ESTATUS']=="RECUPERADO"]['MONTO_CALC'].sum()),
                    "EXCEDENTE": float(df_s[df_s['ESTATUS']=="EXCEDENTE"]['MONTO_CALC'].sum())
                }

        with open(os.path.join(ruta_base, "TOTALES_GLOBALES_COBROS.json"), "w", encoding="utf-8") as f:
            json.dump(totales_globales, f, indent=4, ensure_ascii=False)
        with open(os.path.join(ruta_base, "TOTALES_SUCURSALES_COBROS.json"), "w", encoding="utf-8") as f:
            json.dump(totales_sucursales, f, indent=4, ensure_ascii=False)

        # --- 2. GENERACIÓN DE HTMLS ---
        estilo_css = """<style>
            body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #333; padding: 10px; text-align: center; margin: 0; }
            .header-logos { display: flex; justify-content: space-between; align-items: center; padding: 10px 20px; background: white; border-bottom: 4px solid #F9D908; }
            .logo-header { height: 50px; }
            h1 { color: #002060; margin: 0; font-size: 16px; text-transform: uppercase; font-weight: 900; flex-grow: 1; }
            .resumen-grid { display: flex; justify-content: center; gap: 15px; margin: 20px 0; flex-wrap: wrap; }
            .card-resumen { background: white; padding: 20px; border-radius: 12px; text-decoration: none; width: 220px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-bottom: 6px solid #ccc; color: inherit; transition: 0.3s; }
            .card-resumen .monto { font-size: 20px; font-weight: 900; color: #002060; margin: 10px 0; }
            .cobrado { border-color: #27ae60; } .recuperado { border-color: #f1c40f; } .no-pagado { border-color: #ed1c24; } .excedente { border-color: #0070c0; }
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
                if (index >= extensiones.length) { alert("No se encontró la imagen."); return; }
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
            for suc in sorted(df_p['SUCURSAL'].unique()):
                p_suc = os.path.join(ruta_base, n_m, suc); os.makedirs(p_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]

                # Index de la sucursal: Ahora apunta al ancla del mes Y pide abrir la pestaña 'cobs' (Cobros)
                with open(os.path.join(p_suc, "cobros_detalles.html"), "w", encoding="utf-8") as f:
                    v = [df_s[df_s['ESTATUS']==k]['MONTO_CALC'].sum() for k in ['COBRADO', 'RECUPERADO', 'EXCEDENTE', 'NO_PAGADO']]
                    f.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>SISTEMA LUXOR</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div><h2>{suc} | {n_m}</h2><div class='resumen-grid'>")
                    for l, m, cl, url in zip(['Cobrado', 'Pérdida mitigada', 'Excedentes', 'No Pagado'], v, ['cobrado', 'recuperado', 'excedente', 'no-pagado'], ['cobrado.html', 'recuperado.html', 'excedente.html', 'no_pagado.html']):
                        f.write(f"<a href='{url}' class='card-resumen {cl}'><h3>{l}</h3><div class='monto'>${m:,.2f}</div></a>")
                    f.write(f"</div><a href='todo_detallado.html' class='btn'>VER TODO</a><a href='../../index.html?tab=cobs#mes-{n_m}' class='btn'>INICIO</a></body></html>")

                # Detallados
                for est_key, file_name, titulo in [('COBRADO', 'cobrado.html', 'DETALLE COBRADO'), ('RECUPERADO', 'recuperado.html', 'PÉRDIDA MITIGADA'), ('NO_PAGADO', 'no_pagado.html', 'DETALLE NO PAGADO'), ('EXCEDENTE', 'excedente.html', 'DETALLE EXCEDENTES'), ('TODO', 'todo_detallado.html', 'DETALLE COMPLETO')]:
                    df_view = df_s if est_key == 'TODO' else df_s[df_s['ESTATUS'] == est_key]
                    filas = ""
                    for _, r in df_view.iterrows():
                        tds = ""
                        for i in range(MAX_COLS_MOSTRAR):
                            val = r['FILA'][i] if i < len(r['FILA']) else ""
                            if i == 9: 
                                if r['FOTO_BASE'] not in ["", "None", "nan", "0"]:
                                    tds += f"<td><span style='cursor:pointer; color:#002060; text-decoration:underline;' onclick='openModal(\"{r['FOTO_BASE']}\", \"{n_m}\", \"{suc}\")'>${limpiar_monto(val):,.2f}</span></td>"
                                else: tds += f"<td>${limpiar_monto(val):,.2f}</td>"
                            else: tds += f"<td>{str(val) if val is not None else ''}</td>"
                        filas += f"<tr>{tds}</tr>"

                    with open(os.path.join(p_suc, file_name), "w", encoding="utf-8") as f:
                        f.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><h1>{titulo}</h1></div><div class='blue-box-container'><div class='table-responsive'><table><thead><tr>{''.join([f'<th>{c}</th>' for c in df_s['HEADERS'].iloc[0]])}</tr></thead><tbody>{filas}</tbody></table></div><a href='cobros_detalles.html' class='btn'>VOLVER</a></div>{script_modal}</body></html>")

        print("\n✅ PROCESO FINALIZADO: Botones configurados para volver a la pestaña de Cobros.")
    except Exception as e: print(f"❌ Error: {e}")

if __name__ == "__main__":
    generar_reporte_cobros_final()