import os
import json
import sys
import pandas as pd
from datetime import datetime

try:
    from flask import Flask, request, jsonify, render_template_string, send_from_directory
    from werkzeug.utils import secure_filename
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    print("\n! ERROR: Faltan librerias (Flask, Werkzeug u Openpyxl).")
    input("Presiona ENTER para salir...")
    sys.exit()

app = Flask(__name__)
RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
os.chdir(RUTA_RAIZ)

FILL_VERDE = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")    # COBRO
FILL_AMARILLO = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid") # RECUPERACIÓN

SUCURSALES = [
    "BARQUISIMETO", "CASTAÑO", "CENTRAL", "CIRCULO", "BOSQUE", 
    "GUACARA", "IPSFA", "MORA", "VICTORIA", "ACACIAS", 
    "NAGUANAGUA", "SAN DIEGO", "SAN JUAN", "SANTA RITA", "TUCACAS", "VILLAS"
]

MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}

INCIDENCIAS_REF = {
    "NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO.": "TIPO A", "FALTA SELLO, FIRMA O CÉDULA.": "TIPO A", "DOCUMENTO NO LEGIBLE": "TIPO A",
    "DOCUMENTACIÓN ERRÓNEA": "TIPO B", "FISCALIZACIÓN A DESTIEMPO": "TIPO B", "PRODUCTO O SKU DUPLICADO.": "TIPO B", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN.": "TIPO B",
    "FISCALIZACIÓN con USUARIO NO CORRESPONDIENTE": "TIPO C", "ERROR DE KG EN TARA.": "TIPO C", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN.": "TIPO C", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS": "TIPO C",
    "NO SE INDICÓ DIFERENCIA AL DORSO DE LA FACTURA.": "TIPO D", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO.": "TIPO D",
    "RECEPCIÓN SIN AUTORIZACIÓN DE CMF.": "TIPO E", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA.": "TIPO E"
}

def asegurar_id(df):
    columnas_ordenadas = [
        'SUCURSAL', 'PROVEEDOR', 'FACTURA', 'FECHA', 'TIPO FISCALIZACIÓN', 
        'RESPONSABLE', 'INCIDENCIA', 'TIPO DE ERROR', 'OBSERVACIÓN', 
        'MONTO $', 'F COBRADA', 'CLASIFICACIÓN MONTO', 'ID'
    ]
    if 'ID' not in df.columns:
        df['ID'] = [f"HIST_{datetime.now().strftime('%Y%m%d')}_{i}" for i in range(len(df))]
    for col in columnas_ordenadas:
        if col not in df.columns: df[col] = ""
    df = df[columnas_ordenadas]
    df['ID'] = df['ID'].astype(str).str.strip().replace('nan', '')
    return df

def buscar_archivo_lectura(mes, sucursal):
    ruta_dir = os.path.join(RUTA_RAIZ, "cuadros", mes)
    if not os.path.exists(ruta_dir): return None
    suc_norm = sucursal.upper().replace(" ", "").replace("_", "")
    for archivo in os.listdir(ruta_dir):
        if archivo.startswith('~$') or not archivo.endswith(".xlsx"): continue
        nombre_norm = archivo.upper().replace(" ", "").replace("_", "")
        if suc_norm in nombre_norm: return os.path.join(ruta_dir, archivo)
    return None

@app.route('/')
def home():
    hoy = datetime.now().strftime('%Y-%m-%d')
    ruta_cuadros = os.path.join(RUTA_RAIZ, "cuadros")
    meses_existentes = []
    if os.path.exists(ruta_cuadros):
        meses_existentes = sorted([m for m in os.listdir(ruta_cuadros) 
                                 if os.path.isdir(os.path.join(ruta_cuadros, m)) and not m.startswith('~$')])
    return render_template_string(HTML_FORM, sucursales=SUCURSALES, incidencias=INCIDENCIAS_REF, 
                                  fecha_hoy=hoy, meses_con_datos=meses_existentes)

@app.route('/guardar', methods=['POST'])
def guardar():
    try:
        datos = request.form
        mes_nombre = MESES_ES[datetime.strptime(datos.get('FECHA'), '%Y-%m-%d').month]
        sucursal = datos.get('SUCURSAL')
        archivo_excel = buscar_archivo_lectura(mes_nombre, sucursal)
        if not archivo_excel:
            ruta_dir = os.path.join(RUTA_RAIZ, "cuadros", mes_nombre)
            os.makedirs(ruta_dir, exist_ok=True)
            archivo_excel = os.path.join(ruta_dir, f"{sucursal}.xlsx")
        
        df = pd.read_excel(archivo_excel, dtype={'ID': str}) if os.path.exists(archivo_excel) else pd.DataFrame()
        id_edicion = datos.get('ID_EDICION')
        if id_edicion and str(id_edicion).strip() != "":
            df = df[df['ID'] != str(id_edicion).strip()]
            
        archivo_foto = request.files.get('foto')
        f_cobrada_val = datos.get('F_COBRADA_INPUT') or "SIN FOTO"
        
        if archivo_foto and archivo_foto.filename != '':
            ruta_foto = os.path.join(RUTA_RAIZ, "facturas", mes_nombre, sucursal)
            os.makedirs(ruta_foto, exist_ok=True)
            # Se usa el nombre ingresado en F Cobrada para el archivo
            nombre_foto = f"{secure_filename(f_cobrada_val)}.jpg"
            archivo_foto.save(os.path.join(ruta_foto, nombre_foto))
        
        nuevo_id = str(id_edicion).strip() if (id_edicion and str(id_edicion).strip() != "") else datetime.now().strftime('%Y%m%d%H%M%S')
        nueva_fila = {
            'SUCURSAL': sucursal, 'PROVEEDOR': datos.get('PROVEEDOR'), 'FACTURA': datos.get('FACTURA'), 
            'FECHA': datos.get('FECHA'), 'TIPO FISCALIZACIÓN': datos.get('TIPO_FISC'), 'RESPONSABLE': datos.get('RESPONSABLE'),
            'INCIDENCIA': datos.get('INCIDENCIA'), 'TIPO DE ERROR': INCIDENCIAS_REF.get(datos.get('INCIDENCIA'), "N/A"),
            'OBSERVACIÓN': datos.get('OBSERVACION'), 'MONTO $': float(datos.get('MONTO')) if datos.get('CLASIFICACION_MONTO') != "NINGUNA" else 0, 
            'F COBRADA': f_cobrada_val, 'CLASIFICACIÓN MONTO': datos.get('CLASIFICACION_MONTO'), 'ID': nuevo_id
        }
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        df = asegurar_id(df)
        df.to_excel(archivo_excel, index=False)

        if datos.get('CLASIFICACION_MONTO') != "NINGUNA":
            wb = load_workbook(archivo_excel); ws = wb.active; col_monto_idx = -1
            for cell in ws[1]:
                if cell.value == 'MONTO $': col_monto_idx = cell.column; break
            if col_monto_idx != -1:
                target_cell = ws.cell(row=ws.max_row, column=col_monto_idx)
                if datos.get('CLASIFICACION_MONTO') == "COBRO": target_cell.fill = FILL_VERDE
                elif datos.get('CLASIFICACION_MONTO') == "RECUPERACIÓN": target_cell.fill = FILL_AMARILLO
            wb.save(archivo_excel)
        return jsonify({"status": "ok"})
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/listar/<mes>/<sucursal>')
def listar(mes, sucursal):
    ruta = buscar_archivo_lectura(mes, sucursal)
    if ruta and os.path.exists(ruta):
        try:
            df = pd.read_excel(ruta, dtype={'ID': str}).fillna(""); df = asegurar_id(df)
            df_filtrado = df[df['SUCURSAL'].astype(str).str.upper().str.contains(sucursal.upper())]
            return jsonify(df_filtrado.to_dict(orient='records'))
        except: return jsonify([])
    return jsonify([])

@app.route('/borrar', methods=['POST'])
def borrar():
    try:
        data = request.json
        mes = MESES_ES[datetime.strptime(data.get('fecha'), '%Y-%m-%d').month]
        ruta = buscar_archivo_lectura(mes, data.get('sucursal'))
        if ruta and os.path.exists(ruta):
            df = pd.read_excel(ruta, dtype={'ID': str}); df = asegurar_id(df)
            df = df[df['ID'] != str(data.get('id')).strip()]; df.to_excel(ruta, index=False)
            return jsonify({"status": "ok"})
    except: return jsonify({"status": "error"}), 500
    return jsonify({"status": "error"})

@app.route('/RECURSOS/<path:path>')
def recursos(path): return send_from_directory(os.path.join(RUTA_RAIZ, 'RECURSOS'), path)

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Sistema Luxor</title>
    <style>
        :root { --azul: #0844a4; --amarillo: #F9D908; --fondo: #f4f7f6; }
        body { font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; font-size: 12px; }
        .header { background: white; padding: 10px 30px; display: flex; justify-content: space-between; align-items: center; border-bottom: 5px solid var(--amarillo); }
        .header img { height: 40px; }
        .tabs { display: flex; justify-content: center; background: #ddd; }
        .tab-btn { padding: 12px 25px; cursor: pointer; border: none; background: none; font-weight: bold; color: #555; }
        .tab-btn.active { background: white; color: var(--azul); border-top: 4px solid var(--azul); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .card { background: white; max-width: 98%; margin: 15px auto; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        label { font-weight: bold; color: #555; }
        select, input, button, textarea { width: 100%; padding: 8px; border-radius: 5px; border: 1px solid #ccc; box-sizing: border-box; }
        .btn-main { background: var(--azul); color: white; border: none; font-size: 14px; cursor: pointer; font-weight: bold; margin-top: 10px; padding: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 10px; }
        th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
        th { background: var(--azul); color: white; position: sticky; top: 0; }
        .col-factura { max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .col-tipo-error { max-width: 60px; text-align: center; }
        .btn-icon { border: none; width: 22px; height: 22px; cursor: pointer; border-radius: 3px; font-size: 11px; display: inline-flex; align-items: center; justify-content: center; margin-right: 2px; }
        .btn-edit { background: #f0ad4e; color: white; }
        .btn-del { background: #d9534f; color: white; }
    </style>
</head>
<body>
    <div class="header"><img src="/RECURSOS/LOGO.png"> <h1 style="color:var(--azul); font-size:16px;">MATRIZ DE FISCALIZACION</h1> <img src="/RECURSOS/LOGO.png"></div>
    <div class="tabs">
        <button id="btn-tab-crear" class="tab-btn active" onclick="openTab(event, 'tab-crear')">AÑADIR / MODIFICAR</button>
        <button id="btn-tab-ver" class="tab-btn" onclick="openTab(event, 'tab-ver')">GESTIONAR</button>
    </div>
    <div id="tab-crear" class="tab-content active">
        <div class="card" style="max-width: 700px;">
            <form id="formMatriz">
                <input type="hidden" name="ID_EDICION" id="id_edicion">
                <div class="grid">
                    <div style="grid-column: span 2;"><label>SUCURSAL</label>
                        <select name="SUCURSAL" id="f_suc" required>{% for s in sucursales %}<option value="{{s}}">{{s}}</option>{% endfor %}</select>
                    </div>
                    <div><label>FECHA</label><input type="date" name="FECHA" id="f_fec" value="{{fecha_hoy}}" required></div>
                    <div><label>PROVEEDOR</label><input type="text" name="PROVEEDOR" id="f_pro" required></div>
                    <div><label>FACTURA</label><input type="text" name="FACTURA" id="f_fac" required></div>
                    <div><label>TIPO FISCALIZACIÓN</label>
                        <select name="TIPO_FISC" id="f_tipf" required><option value="RECEPCION">RECEPCIÓN</option><option value="DEVOLUCION">DEVOLUCIÓN</option><option value="TRANSFERENCIA">TRANSFERENCIA</option></select>
                    </div>
                    <div><label>RESPONSABLE</label><input type="text" name="RESPONSABLE" id="f_res" required></div>
                    <div style="grid-column: span 2;"><label>CLASIFICACIÓN MONTO</label>
                        <select name="CLASIFICACION_MONTO" id="f_clasm" onchange="toggleMonto()" required>
                            <option value="NINGUNA">NINGUNA</option><option value="COBRO">COBRO</option><option value="RECUPERACIÓN">RECUPERACIÓN</option>
                        </select>
                    </div>
                    <div id="box-monto"><label>MONTO $</label><input type="number" name="MONTO" id="f_mon" step="0.01" value="0"></div>
                    <div id="box-fcobrada"><label>F COBRADA</label><input type="text" name="F_COBRADA_INPUT" id="f_fcob"></div>
                    <div id="box-foto"><label>CARGAR FOTO</label><input type="file" name="foto" id="f_foto" accept="image/*"></div>
                    <div style="grid-column: span 2;"><label>INCIDENCIA</label>
                        <select name="INCIDENCIA" id="inc_sel" required><option value="">Seleccione...</option>{% for inc in incidencias %}<option value="{{inc}}">{{inc}}</option>{% endfor %}</select>
                    </div>
                    <div style="grid-column: span 2;"><label>OBSERVACIONES</label><textarea name="OBSERVACION" id="f_obs" rows="2"></textarea></div>
                </div>
                <button type="submit" class="btn-main" id="btn-submit">GUARDAR REGISTRO</button>
            </form>
        </div>
    </div>
    <div id="tab-ver" class="tab-content">
        <div class="card">
            <div class="grid" style="max-width: 450px; margin-bottom:10px;">
                <select id="filtro_mes">{% for m in meses_con_datos %}<option value="{{m}}">{{m}}</option>{% endfor %}</select>
                <select id="filtro_suc">{% for s in sucursales %}<option value="{{s}}">{{s}}</option>{% endfor %}</select>
            </div>
            <button class="btn-main" style="background:#555; width: 150px; margin:0;" onclick="cargarIncidencias()">FILTRAR DATOS</button>
            <div style="overflow-x: auto; margin-top:15px; max-height: 500px;"><table id="tablaIncidencias"><thead><tr><th>Sucursal</th><th>Proveedor</th><th class="col-factura">Factura</th><th>Fecha</th><th>Tipo</th><th>Responsable</th><th>Incidencia</th><th class="col-tipo-error">Error</th><th>Observación</th><th>Monto</th><th>F Cobrada</th><th>Acciones</th></tr></thead><tbody></tbody></table></div>
        </div>
    </div>
    <script>
        window.onload = () => {
            const lastMes = localStorage.getItem('lastMes');
            const lastSuc = localStorage.getItem('lastSuc');
            if(lastMes) document.getElementById('filtro_mes').value = lastMes;
            if(lastSuc) document.getElementById('filtro_suc').value = lastSuc;
            toggleMonto();
        };
        function toggleMonto() {
            const val = document.getElementById('f_clasm').value;
            
            // Monto: visible si no es NINGUNA
            document.getElementById('box-monto').style.display = (val === "NINGUNA") ? 'none' : 'block';
            document.getElementById('f_mon').required = (val !== "NINGUNA");
            
            // F Cobrada y Cargar Foto: SOLO si es COBRO
            const isCobro = (val === "COBRO");
            document.getElementById('box-fcobrada').style.display = isCobro ? 'block' : 'none';
            document.getElementById('box-foto').style.display = isCobro ? 'block' : 'none';
            document.getElementById('f_fcob').required = isCobro;
            
            if(val === "NINGUNA") document.getElementById('f_mon').value = 0;
            if(!isCobro) {
                document.getElementById('f_fcob').value = "SIN FOTO";
                document.getElementById('f_foto').value = "";
            }
        }
        function openTab(evt, tabName) {
            let i, content, btns;
            content = document.getElementsByClassName("tab-content");
            for (i = 0; i < content.length; i++) content[i].classList.remove("active");
            btns = document.getElementsByClassName("tab-btn");
            for (i = 0; i < btns.length; i++) btns[i].classList.remove("active");
            document.getElementById(tabName).classList.add("active");
            if(evt) evt.currentTarget.classList.add("active");
        }
        document.getElementById('formMatriz').onsubmit = async (e) => {
            e.preventDefault();
            const res = await fetch('/guardar', { method: 'POST', body: new FormData(e.target) });
            if(res.ok) { alert("¡Guardado correctamente!"); location.reload(); }
            else { alert("Error al guardar."); }
        };
        async function cargarIncidencias() {
            const mes = document.getElementById('filtro_mes').value;
            const suc = document.getElementById('filtro_suc').value;
            if(!mes || !suc) return;
            localStorage.setItem('lastMes', mes); localStorage.setItem('lastSuc', suc);
            const res = await fetch(`/listar/${mes}/${suc}`);
            const datos = await res.json();
            const tbody = document.querySelector("#tablaIncidencias tbody");
            tbody.innerHTML = "";
            datos.forEach(r => {
                let rowData = encodeURIComponent(JSON.stringify(r));
                const clean = (val) => (val && val.toString().trim() !== "" && val !== "N/A" && val !== "NINGUNA") ? val : "-";
                tbody.innerHTML += `<tr><td>${clean(r.SUCURSAL)}</td><td>${clean(r.PROVEEDOR)}</td><td class="col-factura">${clean(r.FACTURA)}</td><td>${clean(r.FECHA)}</td><td>${clean(r['TIPO FISCALIZACIÓN'])}</td><td>${clean(r.RESPONSABLE)}</td><td>${clean(r.INCIDENCIA)}</td><td class="col-tipo-error">${clean(r['TIPO DE ERROR'])}</td><td>${clean(r.OBSERVACIÓN)}</td><td>${clean(r['MONTO $'])}</td><td>${clean(r['F COBRADA'])}</td><td style="white-space: nowrap;"><button class="btn-icon btn-edit" onclick="prepararEdicion('${rowData}')">✏️</button><button class="btn-icon btn-del" onclick="borrarIncidencia('${r.ID}', '${r.FECHA}', '${r.SUCURSAL}')">🗑️</button></td></tr>`;
            });
        }
        function prepararEdicion(encodedData) {
            const r = JSON.parse(decodeURIComponent(encodedData));
            document.getElementById('id_edicion').value = r.ID;
            document.getElementById('f_suc').value = r.SUCURSAL;
            document.getElementById('f_fec').value = r.FECHA;
            document.getElementById('f_pro').value = r.PROVEEDOR;
            document.getElementById('f_fac').value = r.FACTURA;
            document.getElementById('f_tipf').value = r['TIPO FISCALIZACIÓN'];
            document.getElementById('f_res').value = r.RESPONSABLE;
            document.getElementById('f_clasm').value = r['CLASIFICACIÓN MONTO'] || "NINGUNA";
            document.getElementById('f_mon').value = r['MONTO $'];
            document.getElementById('f_fcob').value = r['F COBRADA'];
            document.getElementById('inc_sel').value = r.INCIDENCIA;
            document.getElementById('f_obs').value = r.OBSERVACIÓN;
            toggleMonto();
            document.getElementById('btn-tab-crear').click();
            window.scrollTo(0,0);
        }
        async function borrarIncidencia(id, fecha, suc) {
            if(!confirm("¿Desea eliminar el registro?")) return;
            const res = await fetch('/borrar', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({id: id, fecha: fecha, sucursal: suc}) });
            if(res.ok) { alert("Eliminado."); cargarIncidencias(); }
            else { alert("Error al eliminar."); }
        }
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    try: app.run(host='0.0.0.0', port=5000)
    except Exception as e: print(f"ERROR: {e}")