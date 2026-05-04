import os
import json
import sys
import pandas as pd
from datetime import datetime

try:
    from flask import Flask, request, jsonify, render_template_string, send_from_directory, session, redirect, url_for
    from werkzeug.utils import secure_filename
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    from functools import wraps
except ImportError as e:
    print(f"\n! ERROR: Faltan librerias. Detalle: {e}")
    print("Ejecuta: pip install flask pandas openpyxl")
    input("\nPresiona ENTER para salir...")
    sys.exit()

app = Flask(__name__)
app.secret_key = 'luxor_secret_key_2026' 
RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
os.chdir(RUTA_RAIZ)

# --- CONFIGURACIÓN DE ACCESOS ---
USUARIOS = {
    "admin": {"pw": "admin123", "sucursales": ["TODAS"]},
    "ahenriquez": {"pw": "2026", "sucursales": ["IPSFA", "GUACARA"]},
    "ldiaz": {"pw": "2026", "sucursales": ["SAN JUAN", "VICTORIA"]},
    "dflores": {"pw": "2026", "sucursales": ["BARQUISIMETO", "CASTAÑO"]},
    "vroman": {"pw": "2026", "sucursales": ["SAN DIEGO", "CIRCULO"]},
    "sfuente": {"pw": "2026", "sucursales": ["TUCACAS", "NAGUANAGUA", "BOSQUE"]},
    "ialviarez": {"pw": "2026", "sucursales": ["SANTA RITA", "MORA"]},
    "rguzman": {"pw": "2026", "sucursales": ["CENTRAL"]},
    "wcarmona": {"pw": "2026", "sucursales": ["CENTRAL"]},
    "kcalderon": {"pw": "2026", "sucursales": ["ACACIAS", "VILLAS DE ARAGUA"]},
    "hdelgado": {"pw": "2026", "sucursales": ["TODAS"]}
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- LÓGICA DE EXCEL ---
SUCURSALES_FULL = ["BARQUISIMETO", "CASTAÑO", "CENTRAL", "CIRCULO", "BOSQUE", "GUACARA", "IPSFA", "MORA", "VICTORIA", "ACACIAS", "NAGUANAGUA", "SAN DIEGO", "SAN JUAN", "SANTA RITA", "TUCACAS", "VILLAS"]
MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
INCIDENCIAS_REF = {"NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO.": "TIPO A", "FALTA SELLO, FIRMA O CÉDULA.": "TIPO A", "DOCUMENTO NO LEGIBLE": "TIPO A", "DOCUMENTACIÓN ERRÓNEA": "TIPO B", "FISCALIZACIÓN A DESTIEMPO": "TIPO B", "PRODUCTO O SKU DUPLICADO.": "TIPO B", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN.": "TIPO B", "FISCALIZACIÓN con USUARIO NO CORRESPONDIENTE": "TIPO C", "ERROR DE KG EN TARA.": "TIPO C", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN.": "TIPO C", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS": "TIPO C", "NO SE INDICÓ DIFERENCIA AL DORSO DE LA FACTURA.": "TIPO D", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO.": "TIPO D", "RECEPCIÓN SIN AUTORIZACIÓN DE CMF.": "TIPO E", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA.": "TIPO E"}

def asegurar_id(df):
    columnas_ordenadas = ['SUCURSAL', 'PROVEEDOR', 'FACTURA', 'FECHA', 'TIPO FISCALIZACIÓN', 'RESPONSABLE', 'INCIDENCIA', 'TIPO DE ERROR', 'OBSERVACIÓN', 'MONTO $', 'F COBRADA', 'CLASIFICACIÓN MONTO', 'ID']
    if 'ID' not in df.columns:
        df['ID'] = [f"HIST_{datetime.now().strftime('%Y%m%d')}_{i}" for i in range(len(df))]
    for col in columnas_ordenadas:
        if col not in df.columns: df[col] = ""
    return df[columnas_ordenadas]

def buscar_archivo_lectura(mes, sucursal):
    ruta_dir = os.path.join(RUTA_RAIZ, "cuadros", mes)
    if not os.path.exists(ruta_dir): return None
    suc_norm = sucursal.upper().replace(" ", "").replace("_", "")
    for archivo in os.listdir(ruta_dir):
        if archivo.startswith('~$') or not archivo.endswith(".xlsx"): continue
        nombre_norm = archivo.upper().replace(" ", "").replace("_", "")
        if suc_norm in nombre_norm: return os.path.join(ruta_dir, archivo)
    return None

# --- RUTAS ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u, p = request.form.get('user'), request.form.get('pass')
        if u in USUARIOS and USUARIOS[u]['pw'] == p:
            session['user'], session['sucursales'] = u, USUARIOS[u]['sucursales']
            return redirect(url_for('home'))
        return render_template_string(HTML_LOGIN, error="Acceso denegado")
    return render_template_string(HTML_LOGIN)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def home():
    hoy = datetime.now().strftime('%Y-%m-%d')
    ruta_cuadros = os.path.join(RUTA_RAIZ, "cuadros")
    meses_existentes = sorted([m for m in os.listdir(ruta_cuadros) if os.path.isdir(os.path.join(ruta_cuadros, m))]) if os.path.exists(ruta_cuadros) else []
    mis_suc = SUCURSALES_FULL if "TODAS" in session['sucursales'] else session['sucursales']
    return render_template_string(HTML_FORM, sucursales=mis_suc, incidencias=INCIDENCIAS_REF, fecha_hoy=hoy, meses_con_datos=meses_existentes, usuario=session['user'])

@app.route('/guardar', methods=['POST'])
@login_required
def guardar():
    try:
        datos = request.form
        sucursal = datos.get('SUCURSAL')
        mes_nombre = MESES_ES[datetime.strptime(datos.get('FECHA'), '%Y-%m-%d').month]
        archivo_excel = buscar_archivo_lectura(mes_nombre, sucursal) or os.path.join(RUTA_RAIZ, "cuadros", mes_nombre, f"{sucursal}.xlsx")
        os.makedirs(os.path.dirname(archivo_excel), exist_ok=True)
        
        df = pd.read_excel(archivo_excel, dtype={'ID': str}) if os.path.exists(archivo_excel) else pd.DataFrame()
        if datos.get('ID_EDICION'): df = df[df['ID'].astype(str) != str(datos.get('ID_EDICION'))]
            
        nueva_fila = {
            'SUCURSAL': sucursal, 'PROVEEDOR': datos.get('PROVEEDOR'), 'FACTURA': datos.get('FACTURA'), 
            'FECHA': datos.get('FECHA'), 'TIPO FISCALIZACIÓN': datos.get('TIPO_FISC'), 'RESPONSABLE': datos.get('RESPONSABLE'),
            'INCIDENCIA': datos.get('INCIDENCIA'), 'TIPO DE ERROR': INCIDENCIAS_REF.get(datos.get('INCIDENCIA'), "N/A"),
            'OBSERVACIÓN': datos.get('OBSERVACION'), 'MONTO $': float(datos.get('MONTO') or 0), 
            'F COBRADA': datos.get('F_COBRADA_INPUT') or "SIN FOTO", 'CLASIFICACIÓN MONTO': datos.get('CLASIFICACION_MONTO'), 
            'ID': datos.get('ID_EDICION') or datetime.now().strftime('%Y%m%d%H%M%S')
        }
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        asegurar_id(df).to_excel(archivo_excel, index=False)
        return jsonify({"status": "ok"})
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/listar/<mes>/<sucursal>')
@login_required
def listar(mes, sucursal):
    try:
        ruta = buscar_archivo_lectura(mes, sucursal)
        if ruta and os.path.exists(ruta):
            df = pd.read_excel(ruta, dtype={'ID': str})
            df = df[df['SUCURSAL'].astype(str).str.upper() == sucursal.upper()]
            if 'FECHA' in df.columns:
                df['FECHA'] = df['FECHA'].astype(str).replace(['NaT', 'nan'], '')
            return jsonify(df.fillna("").to_dict(orient='records'))
        return jsonify([])
    except: return jsonify([])

@app.route('/borrar', methods=['POST'])
@login_required
def borrar():
    try:
        data = request.json
        f_str = data.get('fecha').split(' ')[0]
        mes = MESES_ES[datetime.strptime(f_str, '%Y-%m-%d').month]
        ruta = buscar_archivo_lectura(mes, data.get('sucursal'))
        if ruta and os.path.exists(ruta):
            df = pd.read_excel(ruta, dtype={'ID': str})
            df = df[df['ID'].astype(str).str.strip() != str(data.get('id')).strip()]
            df.to_excel(ruta, index=False)
            return jsonify({"status": "ok"})
    except: return jsonify({"status": "error"}), 500

@app.route('/RECURSOS/<path:path>')
def recursos(path): return send_from_directory(os.path.join(RUTA_RAIZ, 'RECURSOS'), path)

# --- VISTAS ---
HTML_LOGIN = """
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><style>:root { --azul: #0844a4; --amarillo: #F9D908; } body { font-family: 'Segoe UI', sans-serif; background: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; } .login-card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center; border-top: 8px solid var(--azul); width: 300px; } input { width: 100%; padding: 12px; margin: 10px 0; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; } button { width: 100%; padding: 12px; background: var(--azul); color: white; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; }</style></head>
<body><div class="login-card"><img src="/RECURSOS/LOGO.png" height="40" style="margin-bottom:20px;"><h2 style="color:var(--azul); margin-top:0;">FISCALIZACION</h2>{% if error %}<p style="color:red; font-size:12px;">{{error}}</p>{% endif %}<form method="POST"><input type="text" name="user" placeholder="Usuario" required><input type="password" name="pass" placeholder="Contraseña" required><button type="submit">INGRESAR</button></form></div></body>
</html>
"""

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Sistema Luxor</title>
    <style>
        :root { --azul: #0844a4; --amarillo: #F9D908; --fondo: #f4f7f6; }
        body { font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; font-size: 11px; }
        .header { background: white; padding: 10px 30px; display: flex; justify-content: space-between; align-items: center; border-bottom: 5px solid var(--amarillo); }
        .header img { height: 40px; }
        .tabs { display: flex; justify-content: center; background: #ddd; }
        .tab-btn { padding: 12px 25px; cursor: pointer; border: none; background: none; font-weight: bold; color: #555; }
        .tab-btn.active { background: white; color: var(--azul); border-top: 4px solid var(--azul); }
        .tab-content { display: none; padding: 20px; }
        .tab-content.active { display: block; }
        .card { background: white; max-width: 700px; margin: auto; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        label { font-weight: bold; color: #555; }
        select, input, textarea { width: 100%; padding: 8px; border-radius: 5px; border: 1px solid #ccc; box-sizing: border-box; font-size: 11px; }
        .btn-main { background: var(--azul); color: white; border: none; font-size: 14px; cursor: pointer; font-weight: bold; margin-top: 10px; padding: 10px; width: 100%; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 9px; background: white; }
        th, td { border: 1px solid #ddd; padding: 5px; text-align: left; white-space: nowrap; }
        th { background: var(--azul); color: white; position: sticky; top:0; }
        .btn-icon { border: none; width: 22px; height: 22px; cursor: pointer; border-radius: 3px; display: inline-flex; align-items: center; justify-content: center; margin-right: 2px; }
        .btn-edit { background: #f0ad4e; color: white; }
        .btn-del { background: #d9534f; color: white; }
    </style>
</head>
<body>
    <div class="header">
        <img src="/RECURSOS/LOGO.png">
        <div style="text-align:center">
            <h1 style="color:var(--azul); font-size:16px; margin:0;">MATRIZ DE FISCALIZACION</h1>
            <span style="font-size:10px;">👤 {{usuario}} | <a href="/logout">Salir</a></span>
        </div>
        <img src="/RECURSOS/LOGO.png">
    </div>
    <div class="tabs">
        <button id="btn-tab-crear" class="tab-btn active" onclick="openTab(event, 'tab-crear')">AÑADIR / MODIFICAR</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-ver')">GESTIONAR</button>
    </div>
    <div id="tab-crear" class="tab-content active">
        <div class="card">
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
                        <select name="TIPO_FISC" id="f_tip"><option value="RECEPCION">RECEPCIÓN</option><option value="DEVOLUCION">DEVOLUCIÓN</option></select>
                    </div>
                    <div><label>RESPONSABLE</label><input type="text" name="RESPONSABLE" id="f_res" required></div>
                    <div style="grid-column: span 2;"><label>CLASIFICACIÓN MONTO</label>
                        <select name="CLASIFICACION_MONTO" id="f_clasm" onchange="toggleMonto()">
                            <option value="NINGUNA">NINGUNA</option><option value="COBRO">COBRO</option><option value="RECUPERACIÓN">RECUPERACIÓN</option>
                        </select>
                    </div>
                    <div id="box-monto" style="display:none;"><label>MONTO $</label><input type="number" name="MONTO" id="f_mon" step="0.01"></div>
                    <div id="box-fco" style="display:none;"><label>F COBRADA</label><input type="text" name="F_COBRADA_INPUT" id="f_fco_in"></div>
                    <div style="grid-column: span 2;"><label>INCIDENCIA</label>
                        <select name="INCIDENCIA" id="f_inc" required><option value="">Seleccione...</option>{% for inc in incidencias %}<option value="{{inc}}">{{inc}}</option>{% endfor %}</select>
                    </div>
                    <div style="grid-column: span 2;"><label>OBSERVACIONES</label><textarea name="OBSERVACION" id="f_obs" rows="2"></textarea></div>
                </div>
                <button type="submit" class="btn-main">GUARDAR REGISTRO</button>
            </form>
        </div>
    </div>
    <div id="tab-ver" class="tab-content">
        <div class="card" style="max-width:98%">
            <div class="grid" style="grid-template-columns: 1fr 1fr 150px; margin-bottom:15px;">
                <select id="filtro_mes">{% for m in meses_con_datos %}<option value="{{m}}">{{m}}</option>{% endfor %}</select>
                <select id="filtro_suc">{% for s in sucursales %}<option value="{{s}}">{{s}}</option>{% endfor %}</select>
                <button onclick="cargarDatos()" style="background:#555; color:white; border:none; cursor:pointer; font-weight:bold;">FILTRAR</button>
            </div>
            <div style="overflow-x:auto; max-height:550px; border:1px solid #ccc;">
                <table id="tabla">
                    <thead>
                        <tr>
                            <th>ACCIONES</th>
                            <th>SUCURSAL</th>
                            <th>PROVEEDOR</th>
                            <th>FACTURA</th>
                            <th>FECHA</th>
                            <th>TIPO FISC.</th>
                            <th>RESPONSABLE</th>
                            <th>INCIDENCIA</th>
                            <th>TIPO ERROR</th>
                            <th>OBSERVACIÓN</th>
                            <th>MONTO $</th>
                            <th>F COBRADA</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
    </div>
    <script>
        function toggleMonto() {
            let v = document.getElementById('f_clasm').value;
            document.getElementById('box-monto').style.display = v !== "NINGUNA" ? "block" : "none";
            document.getElementById('box-fco').style.display = v === "COBRO" ? "block" : "none";
        }
        function openTab(e, n) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.getElementById(n).classList.add('active');
            e.currentTarget.classList.add('active');
        }
        document.getElementById('formMatriz').onsubmit = async (e) => {
            e.preventDefault();
            const res = await fetch('/guardar', { method: 'POST', body: new FormData(e.target) });
            if(res.ok) { alert("Guardado"); location.reload(); }
        };
        async function cargarDatos() {
            const m = document.getElementById('filtro_mes').value, s = document.getElementById('filtro_suc').value;
            const res = await fetch(`/listar/${m}/${s}`);
            const data = await res.json();
            const b = document.querySelector("#tabla tbody"); b.innerHTML = "";
            data.forEach(r => {
                let rData = encodeURIComponent(JSON.stringify(r));
                b.innerHTML += `<tr>
                    <td>
                        <button class="btn-icon btn-edit" onclick="editar('${rData}')">✏️</button>
                        <button class="btn-icon btn-del" onclick="borrar('${r.ID}', '${r.FECHA}', '${r.SUCURSAL}')">🗑️</button>
                    </td>
                    <td>${r.SUCURSAL}</td>
                    <td>${r.PROVEEDOR}</td>
                    <td>${r.FACTURA}</td>
                    <td>${r.FECHA.split(' ')[0]}</td>
                    <td>${r['TIPO FISCALIZACIÓN']}</td>
                    <td>${r.RESPONSABLE}</td>
                    <td>${r.INCIDENCIA}</td>
                    <td>${r['TIPO DE ERROR']}</td>
                    <td>${r.OBSERVACIÓN}</td>
                    <td>${r['MONTO $']}</td>
                    <td>${r['F COBRADA']}</td>
                </tr>`;
            });
        }
        function editar(data) {
            const r = JSON.parse(decodeURIComponent(data));
            document.getElementById('id_edicion').value = r.ID;
            document.getElementById('f_suc').value = r.SUCURSAL;
            document.getElementById('f_fec').value = r.FECHA.split(' ')[0];
            document.getElementById('f_pro').value = r.PROVEEDOR;
            document.getElementById('f_fac').value = r.FACTURA;
            document.getElementById('f_tip').value = r['TIPO FISCALIZACIÓN'];
            document.getElementById('f_res').value = r.RESPONSABLE;
            document.getElementById('f_clasm').value = r['CLASIFICACIÓN MONTO'] || "NINGUNA";
            document.getElementById('f_mon').value = r['MONTO $'];
            document.getElementById('f_fco_in').value = r['F COBRADA'];
            document.getElementById('f_inc').value = r.INCIDENCIA;
            document.getElementById('f_obs').value = r.OBSERVACIÓN;
            toggleMonto();
            document.getElementById('btn-tab-crear').click();
        }
        async function borrar(id, fecha, suc) {
            if(!confirm("¿Eliminar registro?")) return;
            const res = await fetch('/borrar', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({id: id, fecha: fecha, sucursal: suc}) });
            if(res.ok) cargarDatos();
        }
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    try:
        app.run(host='0.0.0.0', port=5000, debug=False)
    except Exception as e:
        print(f"\n--- ERROR AL INICIAR EL SERVIDOR ---\n{e}")
        input("\nPresiona ENTER para salir...")