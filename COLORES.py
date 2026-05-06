import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import warnings

# Silenciar advertencias de validación
warnings.filterwarnings("ignore", category=UserWarning)

def colorear_excels_raiz():
    print("\n" + "="*60)
    print(">>> COLOREANDO MONTOS (EJECUTANDO EN RAÍZ) <<<")
    print("="*60)

    # Detecta la carpeta donde está el script actual
    ruta_raiz = os.path.dirname(os.path.abspath(__file__))

    # Definición de Colores Standard
    verde_fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")
    amarillo_fill = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
    azul_fill = PatternFill(start_color="FF00B0F0", end_color="FF00B0F0", fill_type="solid")

    # Busca archivos .xlsx en la misma carpeta del script
    archivos = [f for f in os.listdir(ruta_raiz) if f.endswith(".xlsx") and not f.startswith("~$")]

    if not archivos:
        print("⚠️ No se encontraron archivos Excel en esta carpeta.")
        input("\nPresiona ENTER para salir...")
        return

    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_raiz, archivo)
        print(f"📄 Procesando: {archivo}...")
        
        try:
            # data_only=False para no perder fórmulas al guardar
            wb = load_workbook(ruta_archivo)
            ws = wb.active

            # 1. Encontrar la columna 'MONTO' dinámicamente
            headers = [str(cell.value).upper() if cell.value else "" for cell in ws[1]]
            idx_monto = -1
            for i, h in enumerate(headers):
                if 'MONTO' in h:
                    idx_monto = i + 1
                    break

            if idx_monto == -1:
                print(f"   ⚠️ No se encontró la columna 'MONTO'.")
                continue

            # 2. Leer Columna L (12) y pintar Monto
            col_l = 12 

            for row in range(2, ws.max_row + 1):
                # Extraer valor de clasificación
                celda_l = ws.cell(row=row, column=col_l).value
                valor_l = str(celda_l).strip().upper() if celda_l else ""
                
                celda_monto = ws.cell(row=row, column=idx_monto)

                if valor_l == "COBRO":
                    celda_monto.fill = verde_fill
                elif "RECUPERACI" in valor_l:
                    celda_monto.fill = amarillo_fill
                elif "EXCEDENTE" in valor_l:
                    celda_monto.fill = azul_fill
                elif valor_l == "NINGUNA":
                    celda_monto.fill = PatternFill(fill_type=None) # Sin color

            wb.save(ruta_archivo)
            print(f"   ✅ Coloreado y guardado.")

        except Exception as e:
            print(f"   ❌ Error en {archivo}: {e}")

    print("\n" + "="*60)
    print(">>> PROCESO FINALIZADO <<<")
    input("Presiona ENTER para cerrar...")

if __name__ == "__main__":
    colorear_excels_raiz()