import os
import json
import sys
import threading
import time
from bs4 import BeautifulSoup

# =================================================================
# ID: DETECTOR DE INCIDENCIAS GRAVES (LUXOR) - AGRUPADO POR MES
# FUNCIÓN: Sumar cantidades de columnas de meses para Tipos D y E
# =================================================================

def generar_json_agrupado_por_mes():
    try:
        print("\n" + "="*60)
        print(">>> GENERANDO INFORME AGRUPADO POR MES (SOLO D y E) <<<")
        print("="*60)
        
        ruta_script = os.path.dirname(os.path.abspath(sys.argv[0]))
        
        # Diccionario principal que agrupará todo: { "MAYO": [...], "ABRIL": [...] }
        resultado_agrupado = {}

        for root, dirs, files in os.walk(ruta_script):
            for file in files:
                if file.lower() in ["reporte.html", "cobrado.html", "index.html"]:
                    ruta_html = os.path.join(root, file)
                    
                    # 1. IDENTIFICAR EL MES (Basado en la carpeta superior)
                    ruta_partes = ruta_html.upper().split(os.sep)
                    mes_encontrado = None
                    for m in ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO"]:
                        if m in ruta_partes:
                            mes_encontrado = m
                            break
                    
                    if not mes_encontrado: continue
                    
                    sucursal = os.path.basename(root).upper()

                    try:
                        with open(ruta_html, 'r', encoding='utf-8') as f:
                            soup = BeautifulSoup(f, 'html.parser')
                            tablas = soup.find_all('table')
                            
                            suma_d = 0
                            suma_e = 0
                            
                            for tabla in tablas:
                                filas = tabla.find_all('tr')
                                if not filas: continue

                                headers = [th.get_text().strip().upper() for th in filas[0].find_all(['th', 'td'])]
                                
                                try:
                                    idx_tipo = headers.index("TIPO")
                                    # Buscar la columna exacta del mes (ej. "MAYO")
                                    idx_col_mes = next(i for i, h in enumerate(headers) if mes_encontrado in h)
                                except (ValueError, StopIteration):
                                    continue

                                for fila in filas[1:]:
                                    celdas = fila.find_all(['td', 'th'])
                                    if len(celdas) > max(idx_tipo, idx_col_mes):
                                        tipo_letra = celdas[idx_tipo].get_text().strip().upper()
                                        
                                        # 2. FILTRAR SOLO D Y E
                                        if tipo_letra in ["D", "E"]:
                                            try:
                                                txt_cant = celdas[idx_col_mes].get_text().strip()
                                                # Extraer solo el número
                                                valor = int(''.join(filter(str.isdigit, txt_cant)) or 0)
                                                
                                                if tipo_letra == "D":
                                                    suma_d += valor
                                                else:
                                                    suma_e += valor
                                            except ValueError:
                                                continue

                            # 3. GUARDAR EN EL GRUPO DEL MES CORRESPONDIENTE
                            if (suma_d + suma_e) > 0:
                                if mes_encontrado not in resultado_agrupado:
                                    resultado_agrupado[mes_encontrado] = []
                                
                                resultado_agrupado[mes_encontrado].append({
                                    "n": f"{sucursal} ({mes_encontrado})",
                                    "v": int(suma_d + suma_e),
                                    "detalle": f"D: {suma_d} | E: {suma_e}"
                                })
                                print(f"✅ Agregado a {mes_encontrado}: {sucursal} (Total: {suma_d + suma_e})")

                    except Exception as e:
                        print(f"❌ Error en {file}: {e}")

        # 4. ESCRITURA DEL JSON FINAL
        if resultado_agrupado:
            ruta_final = os.path.join(ruta_script, "incidencias_graves.json")
            with open(ruta_final, "w", encoding="utf-8") as f:
                json.dump(resultado_agrupado, f, indent=4, ensure_ascii=False)
            print(f"\n🚀 JSON generado exitosamente y AGRUPADO POR MES.")
        else:
            print("\n⚠️ No se encontraron datos para los tipos D o E.")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")

    # Lógica de cierre automático o por teclado
    print("\nPresiona ENTER para salir o espera 10 segundos...")
    timer = threading.Timer(10.0, lambda: os._exit(0))
    timer.start()
    try:
        input()
    finally:
        timer.cancel()

if __name__ == "__main__":
    generar_json_agrupado_por_mes()