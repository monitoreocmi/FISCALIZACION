import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog
import sys
import warnings

warnings.filterwarnings("ignore")

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

def procesar_luxor_columna_h():
    try:
        print("\n" + "="*60)
        print(">>> SISTEMA LUXOR: DETECTOR DE 'TIPO D' Y 'TIPO E' <<<")
        print("="*60)
        
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        directorio = filedialog.askdirectory(title="Selecciona la carpeta con los Excels")
        if not directorio: return

        lista_acumulada = []
        archivos = [f for f in os.listdir(directorio) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

        for i, archivo in enumerate(archivos, 1):
            ruta_completa = os.path.join(directorio, archivo)
            print(f"[{i}/{len(archivos)}] Analizando: {archivo}")
            
            try:
                # Cargar el archivo
                df = pd.read_excel(ruta_completa)
                
                # Estandarizar nombres de columnas
                df.columns = [str(c).upper().strip() for c in df.columns]
                
                # Buscar columnas necesarias
                col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
                col_suc = next((c for c in df.columns if 'SUCURSAL' in c), None)

                # Si no hay columnas de identidad, saltamos
                if not col_fecha or not col_suc:
                    continue

                # ESCANEO DE FILAS: Buscamos el texto que contenga 'TIPO D' o 'TIPO E'
                # Convertimos todo a string para evitar errores con números
                for _, fila in df.iterrows():
                    fila_str = " ".join([str(val).upper() for val in fila.values])
                    
                    tipo_encontrado = None
                    if "TIPO D" in fila_str: tipo_encontrado = "D"
                    elif "TIPO E" in fila_str: tipo_encontrado = "E"
                    
                    if tipo_encontrado:
                        try:
                            fecha_dt = pd.to_datetime(fila[col_fecha], errors='coerce')
                            mes_num = fecha_dt.month if pd.notnull(fecha_dt) else 4
                            
                            lista_acumulada.append({
                                'SUCURSAL': str(fila[col_suc]).upper().strip(),
                                'MES': MESES_ES.get(mes_num, "ABRIL"),
                                'TIPO': tipo_encontrado
                            })
                        except: continue

            except Exception as e:
                print(f"   ❌ Error en archivo: {e}")

        if not lista_acumulada:
            print("\n❌ ERROR: No se encontró el texto 'TIPO D' o 'TIPO E' en los archivos.")
            input(); return

        # Consolidar
        df_final = pd.DataFrame(lista_acumulada)
        resumen = df_final.groupby(['SUCURSAL', 'MES', 'TIPO']).size().unstack(fill_value=0)
        
        for letra in ['D', 'E']:
            if letra not in resumen: resumen[letra] = 0
        
        resumen['TOTAL'] = resumen['D'] + resumen['E']
        resumen = resumen.sort_values(by='TOTAL', ascending=False).reset_index()

        # Formato JSON
        resultado_json = [{
            "n": f"{r['SUCURSAL']} ({r['MES']})",
            "v": int(r['TOTAL']),
            "detalle": f"D: {int(r['D'])} | E: {int(r['E'])}"
        } for _, r in resumen.iterrows()]

        # Guardar
        ruta_script = os.path.dirname(os.path.abspath(sys.argv[0]))
        with open(os.path.join(ruta_script, "incidencias_graves.json"), "w", encoding="utf-8") as f:
            json.dump(resultado_json, f, indent=4, ensure_ascii=False)

        print(f"\n✅ ¡CONSEGUIDO! Se detectaron {len(lista_acumulada)} incidencias.")
        print(f"El ranking tiene {len(resultado_json)} sucursales.")
        input("Presiona ENTER para finalizar...")

    except Exception as e:
        print(f"\n❌ ERROR GENERAL: {e}")
        input("ENTER para salir...")

if __name__ == "__main__":
    procesar_luxor_columna_h()