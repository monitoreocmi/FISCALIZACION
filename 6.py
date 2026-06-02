import os
import json
import sys
from bs4 import BeautifulSoup

# =================================================================
# ID: DETECTOR DE INCIDENCIAS GRAVES (LUXOR) - AGRUPADO POR MES
# FUNCIÓN: Sumar cantidades de columnas de meses para Tipos D y E
# =================================================================

# Configuración de salida para Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def ejecutar_auditoria_incidencias():
    try:
        print("\n" + "="*60)
        print(">>> LUXOR: AUDITORIA DE INCIDENCIAS GRAVES (D y E) <<<")
        print("="*60)
        
        ruta_script = os.path.dirname(os.path.abspath(__file__))
        resultado_agrupado = {}
        
        MESES_VALIDOS = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

        # Escaneo recursivo de reportes
        for root, dirs, files in os.walk(ruta_script):
            for file in files:
                if file.lower() in ["reporte.html", "solo_mes.html"]:
                    ruta_html = os.path.join(root, file)
                    
                    # 1. Identificar el Mes por la ruta
                    ruta_partes = ruta_html.upper().split(os.sep)
                    mes_encontrado = next((m for m in MESES_VALIDOS if m in ruta_partes), None)
                    
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

                            # Localizar columnas de interés
                            headers = [th.get_text().strip().upper() for th in filas[0].find_all(['th', 'td'])]
                            
                            try:
                                idx_tipo = headers.index("TIPO")
                                # Buscamos la columna que coincida con el mes de la carpeta
                                idx_col_mes = next(i for i, h in enumerate(headers) if mes_encontrado in h)
                            except (ValueError, StopIteration):
                                continue

                            for fila in filas[1:]:
                                celdas = fila.find_all(['td', 'th'])
                                if len(celdas) > max(idx_tipo, idx_col_mes):
                                    tipo_letra = celdas[idx_tipo].get_text().strip().upper()
                                    
                                    if tipo_letra in ["D", "E"]:
                                        txt_cant = celdas[idx_col_mes].get_text().strip()
                                        # Extraer solo números (limpieza de basura)
                                        valor_str = ''.join(filter(str.isdigit, txt_cant))
                                        valor = int(valor_str) if valor_str else 0
                                        
                                        if tipo_letra == "D": suma_d += valor
                                        elif tipo_letra == "E": suma_e += valor

                        # 2. Registrar hallazgos si la suma es mayor a cero
                        if (suma_d + suma_e) > 0:
                            if mes_encontrado not in resultado_agrupado:
                                resultado_agrupado[mes_encontrado] = []
                            
                            resultado_agrupado[mes_encontrado].append({
                                "sucursal": sucursal,
                                "total": int(suma_d + suma_e),
                                "detalle": f"D: {suma_d} | E: {suma_e}"
                            })
                            print(f"Hallazgo: {sucursal} ({mes_encontrado}) -> D:{suma_d} E:{suma_e}")

                    except Exception as e:
                        print(f"Error procesando {file} en {sucursal}: {e}")

        # 3. Guardar JSON final
        if resultado_agrupado:
            ruta_final = os.path.join(ruta_script, "incidencias_graves.json")
            with open(ruta_final, "w", encoding="utf-8") as f:
                json.dump(resultado_agrupado, f, indent=4, ensure_ascii=False)
            print(f"\nOK: 'incidencias_graves.json' generado con exito.")
        else:
            print("\nResultado: No se detectaron fallas criticas.")

    except Exception as e:
        print(f"\nERROR CRITICO: {e}")

if __name__ == "__main__":
    ejecutar_auditoria_incidencias()