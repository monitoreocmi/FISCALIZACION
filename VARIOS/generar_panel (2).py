import os
import sys
import json
import time

def generar_panel_luxor():
    try:
        ruta = os.path.dirname(os.path.abspath(sys.argv[0]))
        MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        
        # Cargar datos para verificar
        def cargar(n):
            p = os.path.join(ruta, n)
            return json.load(open(p, "r", encoding="utf-8")) if os.path.exists(p) else {}

        data_totales = cargar("incidencias_totales.json")
        
        # Buscar carpetas de meses
        carpetas = sorted([m for m in os.listdir(ruta) if m.upper() in MESES and os.path.isdir(os.path.join(ruta, m))], 
                         key=lambda x: MESES.index(x.upper()))

        if not carpetas:
            print("No se encontraron carpetas de meses.")
            return

        html_data = ""
        options = ""

        for m in carpetas:
            m_key = m.upper()
            options += f'<option value="mes-{m}">{m}</option>'
            ruta_m = os.path.join(ruta, m)
            sucs = sorted([s for s in os.listdir(ruta_m) if os.path.isdir(os.path.join(ruta_m, s))])
            
            html_data += f'<div id="mes-{m}" class="mes-container">'
            html_data += f'<h2 class="sub-title">REPORTE {m_key}</h2>'
            html_data += '<div class="grid">'
            for s in sucs:
                html_data += f'<a href="{m}/{s}/reporte.html" class="card">{s}</a>'
            html_data += '</div></div>'

        # Template HTML con CSS separado de las llaves de Python
        css = """
        :root { --azul: #0844a4; --amarillo: #F9D908; }
        body { font-family: sans-serif; background: #f4f4f4; margin: 0; padding: 0; }
        header { background: white; padding: 20px; border-bottom: 5px solid var(--amarillo); text-align: center; }
        .controls { padding: 20px; text-align: center; background: #eee; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 15px; padding: 20px; max-width: 1200px; margin: auto; }
        .card { background: var(--azul); color: white; padding: 25px 10px; text-decoration: none; border-radius: 8px; text-align: center; font-weight: bold; font-size: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .card:hover { background: var(--amarillo); color: var(--azul); }
        .mes-container { display: none; }
        .active { display: block; }
        .sub-title { text-align: center; color: var(--azul); margin-top: 20px; }
        select { padding: 10px; font-weight: bold; border-radius: 5px; border: 2px solid var(--azul); }
        """

        html_final = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Panel Luxor V3</title>
    <style>{css}</style>
</head>
<body>
    <header><h1>FISCALIZACIÓN LUXOR</h1></header>
    <div class="controls">
        <label>Seleccionar Mes: </label>
        <select id="mes-selector" onchange="cambiar()">{options}</select>
    </div>
    <main id="main-content">{html_data}</main>
    <script>
        function cambiar() {{
            document.querySelectorAll(".mes-container").forEach(e => e.classList.remove("active"));
            let val = document.getElementById("mes-selector").value;
            let target = document.getElementById(val);
            if(target) target.classList.add("active");
        }}
        window.onload = cambiar;
    </script>
</body>
</html>"""
        
        with open(os.path.join(ruta, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_final)
        print("Finalizado: index.html generado.")
        time.sleep(2)
    except Exception as e:
        print(f"Error: {e}")
        time.sleep(5)

if __name__ == "__main__":
    generar_panel_luxor()
