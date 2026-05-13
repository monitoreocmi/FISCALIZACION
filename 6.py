import pandas as pd
import os
import json
import sys
import warnings
import threading 

# =================================================================
# ID: DETECTOR DE INCIDENCIAS GRAVES (LUXOR)
# FUNCIÓN: Clasificación y conteo de alertas Tipo D y Tipo E
# =================================================================

# Silenciar advertencias de validación de Excel para una consola limpia
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
        print(">>> SISTEMA LUXOR: DETECTOR AUTO DE 'TIPO D' Y 'TIPO E' <<<")
        print("="*60)
        
        ruta_script = os.path.dirname(os.path.abspath(sys.argv[0]))
        ruta_cuadros = os.path.join(ruta_script, "cuadros")

        if not os.path.isdir(ruta_cuadros):
            print(f"❌ ERROR: No se encuentra la carpeta 'cuadros' en:\n{ruta_cuadros}")
            return

        archivos = []
        for root, dirs, files in os.walk(ruta_cuadros):
            for file in files:
                if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                    archivos.append(os.path.join(root, file))

        if not archivos:
            print(f"⚠️ No hay archivos Excel para analizar en: {ruta_cuadros}")
            return

        print(f"✅ Se encontraron {len(archivos)} archivos. Iniciando análisis...\n")

        lista_acumulada = []

        for i, ruta_completa in enumerate(archivos, 1):
            nombre_archivo = os.path.basename(ruta_completa)
            print(f"[{i}/{len(archivos)}] Analizando: {nombre_archivo}")
            
            try:
                # Lectura rápida del Excel
                df = pd.read_excel(ruta_completa)
                df.columns = [str(c).upper().strip() for c in df.columns]
                
                # Identificación de columnas clave
                col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
                col_suc = next((c for c in df.columns if 'SUCURSAL' in c), None)

                if not col_fecha or not col_suc:
                    continue

                for _, fila in df.iterrows():
                    # Convertimos toda la fila a texto para buscar las etiquetas de tipo
                    fila_str = " ".join([str(val).upper() for val in fila.values])
                    
                    tipo_encontrado = None
                    if "TIPO D" in fila_str: tipo_encontrado = "D"
                    elif "TIPO E" in fila_str: tipo_encontrado = "E"
                    
                    if tipo_encontrado:
                        try:
                            fecha_dt = pd.to_datetime(fila[col_fecha], errors='coerce')
                            # Si no hay fecha, se asigna un mes por defecto o el actual
                            mes_num = fecha_dt.month if pd.notnull(fecha_dt) else 4 
                            
                            lista_acumulada.append({
                                'SUCURSAL': str(fila[col_suc]).upper().strip(),
                                'MES': MESES_ES.get(mes_num, "ABRIL"),
                                'TIPO': tipo_encontrado
                            })
                        except: continue

            except Exception as e:
                print(f"   ❌ Error en archivo {nombre_archivo}: {e}")

        if not lista_acumulada:
            print("\n❌ No se detectaron incidencias críticas en los archivos procesados.")
        else:
            # Procesamiento de datos para el resumen
            df_final = pd.DataFrame(lista_acumulada)
            resumen = df_final.groupby(['SUCURSAL', 'MES', 'TIPO']).size().unstack(fill_value=0)
            
            # Asegurar que ambas columnas existan en el DataFrame
            for letra in ['D', 'E']:
                if letra not in resumen: resumen[letra] = 0
            
            resumen['TOTAL'] = resumen['D'] + resumen['E']
            resumen = resumen.sort_values(by='TOTAL', ascending=False).reset_index()

            # Formatear para el JSON del Dashboard
            resultado_json = [{
                "n": f"{r['SUCURSAL']} ({r['MES']})",
                "v": int(r['TOTAL']),
                "detalle": f"D: {int(r['D'])} | E: {int(r['E'])}"
            } for _, r in resumen.iterrows()]

            # Guardar el archivo final
            with open(os.path.join(ruta_script, "incidencias_graves.json"), "w", encoding="utf-8") as f:
                json.dump(resultado_json, f, indent=4, ensure_ascii=False)

            print(f"\n✅ ¡CONSEGUIDO! Se registraron {len(lista_acumulada)} incidencias graves.")

    except Exception as e:
        print(f"\n❌ ERROR GENERAL EN EL PROCESO: {e}")

    # Temporizador de salida
    print("\n" + "-"*30)
    print("Presiona ENTER para salir (o espera 10 segundos)...")
    
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__":
    procesar_luxor_columna_h()