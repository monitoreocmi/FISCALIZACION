import pandas as pd
import os
import sys
import json
from openpyxl import load_workbook
import warnings
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

RUTA_LOGO_ESTANDAR = "../../RECURSOS/logo.png"

def limpiar_monto(valor):
    if valor is None or str(valor).strip().lower() in ['none', 'nan', '']: return 0.0
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
    return indices

def generar_reporte_cobros_final():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: PROCESANDO COLUMNA J (MONTO $) <<<")
        print("="*60)
        
        ruta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_base, "cuadros")
        
        if not os.path.exists(ruta_cuadros):
            print(f"❌ Error: No existe la carpeta 'cuadros' en {ruta_base}")
            return

        archivos = [os.path.join(root, f) for root, dirs, files in os.walk(ruta_cuadros) 
                    for f in files if f.endswith(".xlsx") and not f.startswith("~$")]

        print(f"📂 Archivos encontrados: {len(archivos)}")
        datos_finales = []

        for f in archivos:
            nombre_f = os.path.basename(f)
            print(f"📖 Analizando: {nombre_f}")
            wb = load_workbook(f, data_only=False)
            ws = wb.active
            headers_reales = [str(cell.value).strip() if cell.value else f"COL_{i+1}" for i, cell in enumerate(ws[1])]
            idx = obtener_indices_flexibles(headers_reales)
            
            count_filas = 0
            for row in ws.iter_rows(min_row=2):
                row_vals = [cell.value for cell in row]
                if not any(v is not None for v in row_vals): continue
                
                fecha_val = pd.to_datetime(row_vals[idx['fecha']], errors='coerce') if idx['fecha'] != -1 else None
                mes_num = fecha_val.month if fecha_val and not pd.isna(fecha_val) else 0
                
                estatus = "OTRO"
                try:
                    target_cell = row[idx['monto']]
                    color = str(target_cell.fill.start_color.index).upper()
                    
                    if color in ['FF00B050', 'FF92D050', '00FF00', 'FF00FF00', 'FFC6EFCE', '13', 'FF548235']: 
                        estatus = "COBRADO"
                    elif color in ['FFFFFF00', 'FFFFFFE1', 'FFFFEB9C', '17', 'FFFFD966', 'FFFFC000']: 
                        estatus = "RECUPERADO"
                    elif color in ['FF0070C0', 'FF00B0F0', 'FFCCE5FF', '24', '30', 'FF3D85C6']: 
                        estatus = "EXCEDENTE"
                except: pass

                if estatus != "OTRO":
                    mes_nombre = MESES_ES.get(mes_num, "VARIOS")
                    datos_finales.append({
                        'SUCURSAL': str(row_vals[idx['sucursal']]).strip().upper() if idx['sucursal'] != -1 else "GENERAL",
                        'MES': mes_nombre, 
                        'PERIODO': fecha_val.to_period('M') if fecha_val and not pd.isna(fecha_val) else pd.Period('2026-01', freq='M'),
                        'MONTO_CALC': limpiar_monto(row_vals[idx['monto']]),
                        'FOTO_BASE': str(row_vals[idx['foto']]).strip() if idx['foto'] != -1 and row_vals[idx['foto']] else "",
                        'ESTATUS': estatus, 'FILA': row_vals, 'HEADERS': headers_reales, 'IDX': idx
                    })
                    count_filas += 1
            wb.close()
            print(f"   ✅ {count_filas} registros extraídos.")

        if not datos_finales:
            print("⚠️ No se encontraron filas coloreadas.")
            return

        df = pd.DataFrame(datos_finales)
        print(f"\n📊 Generando JSON y reportes HTML para {len(df)} registros totales...")

        # TOTALES JSON
        t_glob = {}
        for m_n in df['MES'].unique():
            df_m = df[df['MES'] == m_n]
            t_glob[str(m_n).upper()] = {
                "TOTAL_COBRADO": float(df_m[df_m['ESTATUS'] == 'COBRADO']['MONTO_CALC'].sum()),
                "TOTAL_PERDIDA_PATRIMONIO": float(df_m[df_m['ESTATUS'] == 'RECUPERADO']['MONTO_CALC'].sum()),
                "TOTAL_EXCEDENTE": float(df_m[df_m['ESTATUS'] == 'EXCEDENTE']['MONTO_CALC'].sum())
            }
        
        with open(os.path.join(ruta_base, "TOTALES_GLOBALES_COBROS.json"), "w", encoding="utf-8") as f_json:
            json.dump(t_glob, f_json, indent=4)

        # HTML GENERATION
        estilo_css = """<style>body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #333; padding: 10px; text-align: center; margin: 0; } .header-logos { display: flex; justify-content: space-between; align-items: center; padding: 10px 20px; background: white; border-bottom: 4px solid #F9D908; } .logo-header { height: 50px; } h1 { color: #002060; margin: 0; font-size: 16px; text-transform: uppercase; font-weight: 900; flex-grow: 1; } .resumen-grid { display: flex; justify-content: center; gap: 15px; margin: 20px 0; flex-wrap: wrap; } .card-resumen { background: white; padding: 20px; border-radius: 12px; text-decoration: none; width: 220px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-bottom: 6px solid #ccc; color: inherit; transition: 0.3s; } .card-resumen .monto { font-size: 20px; font-weight: 900; color: #002060; margin: 10px 0; } .cobrado { border-color: #27ae60; } .recuperado { border-color: #f1c40f; } .excedente { border-color: #0070c0; } .blue-box-container { background: #002060; padding: 15px; border-radius: 12px; width: 98%; margin: 10px auto; border: 2px solid #F9D908; color: white; box-sizing: border-box; } .table-responsive { background: white; border-radius: 8px; overflow-x: auto; color: #333; margin-top: 15px; } table { width: 100%; border-collapse: collapse; min-width: 1000px; } th { background: #001a4d; color: #F9D908; padding: 8px; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #F9D908; white-space: nowrap; } td { padding: 6px; border-bottom: 1px solid #eee; font-size: 10px; font-weight: bold; text-align: left; } .btn { padding: 10px 18px; background: #002060; color: white !important; text-decoration: none; font-weight: bold; border-radius: 6px; border: 2px solid #F9D908; display: inline-block; margin: 5px; font-size: 11px; } .foto-link { color: #002060; text-decoration: underline; font-weight: bold; cursor: pointer; }</style>"""
        script_modal = """<div id="myModal" class="modal" style="display:none; position:fixed; z-index:1000; left:0; top:0; width:100%; height:100%; background:rgba(0,0,0,0.9);" onclick="this.style.display='none'"><span style="position:absolute; top:15px; right:35px; color:#fff; font-size:40px; font-weight:bold; cursor:pointer;">&times;</span><img style="margin:auto; display:block; max-width:90%; max-height:90%; border:3px solid #F9D908; position:relative; top:50%; transform:translateY(-50%);" id="img01"></div><script>function openModal(src) { document.getElementById('myModal').style.display = "block"; document.getElementById('img01').src = src; }</script>"""

        for periodo in df['PERIODO'].unique():
            df_p = df[df['PERIODO'] == periodo]
            n_m = str(df_p['MES'].iloc[0])
            print(f"🔨 Generando archivos para {n_m}...")
            for suc in df_p['SUCURSAL'].unique():
                suc_f = str(suc).strip().upper()
                r_suc = os.path.join(ruta_base, n_m, suc_f)
                os.makedirs(r_suc, exist_ok=True)
                df_s = df_p[df_p['SUCURSAL'] == suc]
                idx_a = df_s['IDX'].iloc[0]

                for est_k, f_n, tit in [('COBRADO', 'cobrado.html', 'DETALLE COBRADO'), ('RECUPERADO', 'recuperado.html', 'PÉRDIDA MITIGADA'), ('EXCEDENTE', 'excedente.html', 'DETALLE EXCEDENTES')]:
                    df_v = df_s[df_s['ESTATUS'] == est_k]
                    filas = ""
                    for _, r in df_v.iterrows():
                        tds = ""
                        for i, v in enumerate(r['FILA']):
                            v_s = str(v).strip() if v is not None else ""
                            if i == idx_a['monto']: tds += f"<td>${r['MONTO_CALC']:,.2f}</td>"
                            elif i == idx_a['foto'] and v_s not in ["", "None", "nan", "SIN FOTO"]:
                                nom_f = v_s if v_s.lower().endswith(('.jpeg', '.jpg', '.png')) else v_s + ".jpeg"
                                r_foto = f"../../FACTURAS/{n_m}/{suc_f}/{nom_f}"
                                tds += f"<td><span class='foto-link' onclick='openModal(\"{r_foto}\")'>{v_s}</span></td>"
                            else: tds += f"<td>{v_s}</td>"
                        filas += f"<tr>{tds}</tr>"

                    h_h = "".join([f"<th>{h}</th>" for h in df_s.iloc[0]['HEADERS']])
                    with open(os.path.join(r_suc, f_n), "w", encoding="utf-8") as f_out:
                        f_out.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><h1>{tit}</h1></div><div class='blue-box-container'><div class='table-responsive'><table><thead><tr>{h_h}</tr></thead><tbody>{filas}</tbody></table></div><a href='cobros_detalles.html' class='btn'>VOLVER</a></div>{script_modal}</body></html>")

                s_c = df_s[df_s['ESTATUS']=='COBRADO']['MONTO_CALC'].sum()
                s_r = df_s[df_s['ESTATUS']=='RECUPERADO']['MONTO_CALC'].sum()
                s_e = df_s[df_s['ESTATUS']=='EXCEDENTE']['MONTO_CALC'].sum()
                
                with open(os.path.join(r_suc, "cobros_detalles.html"), "w", encoding="utf-8") as f_out:
                    f_out.write(f"<html><head><meta charset='UTF-8'>{estilo_css}</head><body><div class='header-logos'><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'><h1>SISTEMA LUXOR</h1><img src='{RUTA_LOGO_ESTANDAR}' class='logo-header'></div><h2>{suc_f} | {n_m}</h2><div class='resumen-grid'>")
                    f_out.write(f"<a href='cobrado.html' class='card-resumen cobrado'><h3>Cobrado</h3><div class='monto'>${s_c:,.2f}</div></a>")
                    f_out.write(f"<a href='recuperado.html' class='card-resumen recuperado'><h3>Pérdida mitigada</h3><div class='monto'>${s_r:,.2f}</div></a>")
                    f_out.write(f"<a href='excedente.html' class='card-resumen excedente'><h3>Excedentes</h3><div class='monto'>${s_e:,.2f}</div></a>")
                    f_out.write(f"</div><a href='../../index.html?tab=cobs#mes-{n_m}' class='btn'>INICIO</a></body></html>")

        print("\n✅ PROCESO COMPLETADO EXITOSAMENTE.")

    except Exception as e:
        print(f"❌ Error durante la ejecución: {e}")

    print("\nPresiona ENTER para salir (o espera 10 segundos)...")
    def auto_close():
        time.sleep(10)
        os._exit(0)

    threading.Thread(target=auto_close, daemon=True).start()
    try: input()
    except: pass

if __name__ == "__main__":
    generar_reporte_cobros_final()