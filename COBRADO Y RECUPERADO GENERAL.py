import os
import json
import re
import sys

# Configuración de codificación para evitar errores en consola
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def limpiar_monto(texto):
    """Extrae solo números y puntos decimales, eliminando símbolos y comas"""
    if not texto: return 0.0
    num = "".join(re.findall(r"[\d.]", texto.replace(',', '')))
    try:
        return float(num)
    except:
        return 0.0

def generar_total_global_mensual_con_color():
    try:
        print("\n" + "="*50)
        print(">>> CONSOLIDADOR DE COBROS CON COMPARATIVA <<<")
        print("="*50)
        
        ruta_raiz = os.path.dirname(os.path.abspath(sys.argv[0]))
        totales_por_mes = {}

        # 1. Recorrer carpetas de meses y sumar
        for carpeta in os.listdir(ruta_raiz):
            mes_key = carpeta.upper()
            if mes_key in MESES_ES:
                ruta_mes = os.path.join(ruta_raiz, carpeta)
                if mes_key not in totales_por_mes:
                    totales_por_mes[mes_key] = {"TOTAL_COBRADO": 0.0, "TOTAL_PERDIDA_PATRIMONIO": 0.0, "COLOR_COBRADO": "NEGRO"}

                for suc in os.listdir(ruta_mes):
                    p_suc = os.path.join(ruta_mes, suc)
                    archivo_cobros = os.path.join(p_suc, "cobros_detalles.html")
                    if os.path.exists(archivo_cobros):
                        with open(archivo_cobros, "r", encoding="utf-8") as f:
                            html = f.read()
                            montos = re.findall(r"<div class='monto'>(.*?)</div>", html)
                            if len(montos) >= 2:
                                totales_por_mes[mes_key]["TOTAL_COBRADO"] += limpiar_monto(montos[0])
                                totales_por_mes[mes_key]["TOTAL_PERDIDA_PATRIMONIO"] += limpiar_monto(montos[1])

        # 2. Lógica de Comparativa de Colores
        # Ordenamos los meses encontrados según el orden real del calendario
        meses_ordenados = [m for m in MESES_ES if m in totales_por_mes]
        
        for i in range(len(meses_ordenados)):
            mes_actual = meses_ordenados[i]
            if i > 0:
                mes_anterior = meses_ordenados[i-1]
                val_act = totales_por_mes[mes_actual]["TOTAL_COBRADO"]
                val_ant = totales_por_mes[mes_anterior]["TOTAL_COBRADO"]

                # Si este mes es mayor al anterior -> VERDE, si es menor -> ROJO
                if val_act > val_ant:
                    totales_por_mes[mes_actual]["COLOR_COBRADO"] = "VERDE"
                elif val_act < val_ant:
                    totales_por_mes[mes_actual]["COLOR_COBRADO"] = "ROJO"
                else:
                    totales_por_mes[mes_actual]["COLOR_COBRADO"] = "NEGRO"
            else:
                # El primer mes de la lista no tiene contra qué comparar
                totales_por_mes[mes_actual]["COLOR_COBRADO"] = "NEGRO"

        # 3. Guardar el archivo JSON
        ruta_salida = os.path.join(ruta_raiz, "TOTALES_GLOBALES_COBROS.json")
        with open(ruta_salida, "w", encoding="utf-8") as jf:
            json.dump(totales_por_mes, jf, ensure_ascii=False, indent=4)

        print("\n✅ Archivo JSON generado con indicadores de color.")
        print(f"Ruta: {ruta_salida}")
        input("\nPresione ENTER para finalizar...")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        input("Presione ENTER para cerrar...")

if __name__ == "__main__":
    generar_total_global_mensual_con_color()