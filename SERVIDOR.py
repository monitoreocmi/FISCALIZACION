import os
import sys
import re
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from functools import wraps

# --- VERIFICACIÓN DE LIBRERÍAS ---
try:
    from flask import Flask, request, jsonify, render_template_string, send_from_directory, session, redirect, url_for
    from flask_cors import CORS
except ImportError:
    print("\n[!] ERROR: Faltan librerias. Ejecuta: pip install flask flask-cors pandas openpyxl requests beautifulsoup4")
    sys.exit()

app = Flask(__name__)
app.secret_key = 'luxor_full_system_2026'
CORS(app)

# --- CONFIGURACIÓN DE RUTAS ---
RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
os.chdir(RUTA_RAIZ)

# Memoria temporal para el marcador
ULTIMO_ANALISIS = {"SUCURSAL": "", "PROVEEDOR": "", "FACTURA": "", "RESPONSABLE": ""}

# --- USUARIOS LOCALES ---
USUARIOS = {
    "admin": {"pw": "admin123", "sucursales": ["TODAS"]},
    "ahenriquez": {"pw": "2026", "sucursales": ["IPSFA", "GUACARA"]},
    "ldiaz": {"pw": "2026", "sucursales": ["SAN JUAN", "VICTORIA"]},
    "dflores": {"pw": "2026", "sucursales": ["BARQUISIMETO", "CASTAÑO"]},
    "vroman": {"pw": "2026", "sucursales": ["SAN DIEGO", "CIRCULO MILITAR"]},
    "sfuente": {"pw": "2026", "sucursales": ["TUCACAS", "NAGUANAGUA", "BOSQUE"]},
    "ialviarez": {"pw": "2026", "sucursales": ["SANTA RITA", "MORA"]},
    "rguzman": {"pw": "2026", "sucursales": ["CENTRAL"]},
    "wcarmona": {"pw": "2026", "sucursales": ["CENTRAL"]},
    "kcalderon": {"pw": "2026", "sucursales": ["ACACIAS", "VILLAS DE ARAGUA"]},
    "hdelgado": {"pw": "2026", "sucursales": ["TODAS"]}
}

SUCURSALES_FULL = ["BARQUISIMETO", "CASTAÑO", "CENTRAL", "CIRCULO MILITAR", "BOSQUE", "GUACARA", "IPSFA", "MORA", "VICTORIA", "ACACIAS", "NAGUANAGUA", "SAN DIEGO", "SAN JUAN", "SANTA RITA", "TUCACAS", "VILLAS DE ARAGUA"]
MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
INCIDENCIAS_REF = {"NÚMERO DE CONTROL O DOCUMENTO ERRÓNEO.": "TIPO A", "FALTA SELLO, FIRMA O CÉDULA.": "TIPO A", "DOCUMENTO NO LEGIBLE": "TIPO A", "DOCUMENTACIÓN ERRÓNEA": "TIPO B", "FISCALIZACIÓN A DESTIEMPO": "TIPO B", "PRODUCTO O SKU DUPLICADO.": "TIPO B", "RECEPCIÓN FUERA DE VISUAL / CON OBSTRUCCIÓN.": "TIPO B", "FISCALIZACIÓN con USUARIO NO CORRESPONDIENTE": "TIPO C", "ERROR DE KG EN TARA.": "TIPO C", "PRODUCTO O SKU NO PERTENECE A LA RECEPCIÓN.": "TIPO C", "NO FISCALIZÓ UNO O VARIOS PRODUCTOS": "TIPO C", "NO SE INDICÓ DIFERENCIA AL DORSO DE LA FACTURA.": "TIPO D", "DIFERENCIA ENTRE CANTIDAD FISCALIZADA Y DOCUMENTO.": "TIPO D", "RECEPCIÓN SIN AUTORIZACIÓN DE CMF.": "TIPO E", "NO SE COMPLETA EL PROCESO DE FISCALIZACION Y SE ELIMINA.": "TIPO E", " OTRAS.": ""}

def login_required(f):
    @wraps(f)
    def dec(*args, **kwargs):
        if 'user' not in session: return redirect(url_for('login'))
        return f(*args, **kwargs)
    return dec

def buscar_archivo(mes, sucursal):
    rd = os.path.join(RUTA_RAIZ, "cuadros", mes.upper())
    if not os.path.exists(rd): return None
    sn = sucursal.upper().replace(" ", "")
    for a in os.listdir(rd):
        if not a.startswith('~$') and a.endswith(".xlsx") and sn in a.upper().replace(" ", ""):
            return os.path.join(rd, a)
    return None

# --- MOTOR DE EXTRACCIÓN INTEGRADO ---
@app.route('/analizar_codigo', methods=['POST'])
def analizar_codigo():
    global ULTIMO_ANALISIS
    try:
        html = request.json.get('html', '')
        soup = BeautifulSoup(html, 'html.parser')
        texto_sucio = " ".join(soup.get_text(" ", strip=True).split()).upper()
        
        m_control = re.search(r"CONTROL\s*[:]?\s*(\d+)", texto_sucio)
        val_control = m_control.group(1) if m_control else "---"
        
        factura_detectada = ""
        m_fact = re.search(r"FACTURA\s*(?:N°|#|:)?\s*(\d+)", texto_sucio)
        if m_fact and m_fact.group(1) != val_control:
            factura_detectada = m_fact.group(1)
        else:
            todos_nums = re.findall(r"\b\d{4,12}\b", texto_sucio)
            for n in todos_nums:
                if n != val_control:
                    factura_detectada = n
                    break

        proveedor = ""
        m_prov = re.search(r"RAZÓN\s*SOCIAL\s*[:]?\s*(.*?)\s*(?:SUCURSAL|RIF|DIRECCIÓN|$)", texto_sucio)
        if m_prov: proveedor = m_prov.group(1).replace(":", "").strip()

        sucursal_final = ""
        for s in SUCURSALES_FULL:
            if s in texto_sucio: sucursal_final = s; break

        responsable = ""
        m_resp = re.search(r"RESPONSABLE\s*[:]?\s*([A-Z\s.]+?)\s*(?:ESTADO|FECHA|SUCURSAL|$)", texto_sucio)
        if m_resp: responsable = m_resp.group(1).replace(":", "").strip()

        ULTIMO_ANALISIS = {
            "SUCURSAL": sucursal_final,
            "PROVEEDOR": proveedor,
            "FACTURA": factura_detectada,
            "RESPONSABLE": responsable
        }
        return jsonify({"status": "ok", "data": ULTIMO_ANALISIS})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route('/obtener_ultimo_analisis', methods=['GET'])
def obtener_ultimo_analisis():
    return jsonify({"status": "ok", "data": ULTIMO_ANALISIS})

# --- CONTINUACIÓN DE RUTAS ORIGINALES ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u, p = request.form.get('user'), request.form.get('pass')
        if u in USUARIOS and USUARIOS[u]['pw'] == p:
            session.clear()
            session['user'] = u
            session['sucursales_permitidas'] = USUARIOS[u]['sucursales']
            return redirect(url_for('home'))
    return render_template_string(HTML_LOGIN)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/guardar', methods=['POST'])
@login_required
def guardar():
    try:
        d = request.form
        suc, fec, ide = d.get('SUCURSAL'), d.get('FECHA'), str(d.get('ID_EDICION') or "").strip()
        mes = MESES_ES[datetime.strptime(fec, '%Y-%m-%d').month]
        path_xlsx = buscar_archivo(mes, suc) or os.path.join(RUTA_RAIZ, "cuadros", mes, f"{suc}.xlsx")
        os.makedirs(os.path.dirname(path_xlsx), exist_ok=True)
        f_cob_val = d.get('F_COBRADA_INPUT') or "SIN_FOTO"
        f_inc_name = d.get('F_INC_EXISTENTE') or ""
        
        for key, folder in [('FOTO_FILE', 'FACTURAS'), ('FOTO_INCIDENCIA_FILE', 'fotos_incidencias')]:
            file = request.files.get(key)
            if file and file.filename != '':
                ext = file.filename.rsplit('.', 1)[1].lower()
                n_final = f"{f_cob_val}.{ext}" if key == 'FOTO_FILE' else f"INC_{datetime.now().strftime('%H%M%S')}.{ext}"
                dest = os.path.join(RUTA_RAIZ, folder, mes, suc)
                os.makedirs(dest, exist_ok=True)
                file.save(os.path.join(dest, n_final))
                if key == 'FOTO_FILE': f_cob_val = n_final
                else: f_inc_name = n_final

        df = pd.read_excel(path_xlsx) if os.path.exists(path_xlsx) else pd.DataFrame()
        if not df.empty and ide: df = df[df['ID'].astype(str) != ide]
        nueva = {
            'SUCURSAL': suc, 'PROVEEDOR': d.get('PROVEEDOR'), 'FACTURA': d.get('FACTURA'),
            'FECHA': fec, 'TIPO FISCALIZACIÓN': d.get('TIPO_FISC'), 'RESPONSABLE': d.get('RESPONSABLE'),
            'INCIDENCIA': d.get('INCIDENCIA'), 'TIPO DE ERROR': INCIDENCIAS_REF.get(d.get('INCIDENCIA'), "N/A"),
            'OBSERVACIÓN': d.get('OBSERVACION'), 'MONTO $': float(d.get('MONTO') or 0),
            'F COBRADA': f_cob_val, 'CLASIFICACIÓN MONTO': d.get('CLASIFICACION_MONTO'),
            'FOTO_INCIDENCIA': f_inc_name, 'ID': ide if ide else f"ID_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        }
        df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
        df.to_excel(path_xlsx, index=False)
        return jsonify({"status": "ok"})
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/ver_foto/<tipo>/<mes>/<sucursal>/<nombre>')
@login_required
def ver_foto(tipo, mes, sucursal, nombre):
    carpeta = "FACTURAS" if tipo.upper() == "FACTURAS" else "fotos_incidencias"
    dir_f = os.path.join(RUTA_RAIZ, carpeta, mes.upper(), sucursal.upper())
    return send_from_directory(dir_f, nombre)

@app.route('/listar/<mes>/<sucursal>')
@login_required
def listar(mes, sucursal):
    ruta = buscar_archivo(mes, sucursal)
    if ruta and os.path.exists(ruta):
        df = pd.read_excel(ruta).fillna("")
        for col in df.columns:
            if 'FECHA' in col.upper(): df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').fillna("")
        return jsonify(df.to_dict(orient='records'))
    return jsonify([])

@app.route('/borrar', methods=['POST'])
@login_required
def borrar():
    data = request.json
    mes = MESES_ES[datetime.strptime(data['fecha'], '%Y-%m-%d').month]
    rx = buscar_archivo(mes, data['sucursal'])
    if rx:
        df = pd.read_excel(rx)
        df = df[df['ID'].astype(str) != str(data['id'])]
        df.to_excel(rx, index=False)
    return jsonify({"status": "ok"})

@app.route('/borrar_foto', methods=['POST'])
@login_required
def borrar_foto():
    data = request.json
    mes = MESES_ES[datetime.strptime(data['fecha'], '%Y-%m-%d').month]
    rx = buscar_archivo(mes, data['sucursal'])
    if rx:
        df = pd.read_excel(rx).fillna("")
        col = 'F COBRADA' if data['tipo'].upper() == 'FACTURAS' else 'FOTO_INCIDENCIA'
        df.loc[df['ID'].astype(str) == str(data['id']), col] = ("SIN_FOTO" if col == 'F COBRADA' else "")
        df.to_excel(rx, index=False)
    return jsonify({"status": "ok"})

@app.route('/')
@login_required
def home():
    hoy = datetime.now().strftime('%Y-%m-%d')
    p_cuadros = os.path.join(RUTA_RAIZ, "cuadros")
    mc = sorted([m for m in os.listdir(p_cuadros) if os.path.isdir(os.path.join(p_cuadros, m))]) if os.path.exists(p_cuadros) else []
    permisos = session.get('sucursales_permitidas', [])
    sucs_usuario = SUCURSALES_FULL if "TODAS" in permisos else [s for s in SUCURSALES_FULL if s in permisos]
    return render_template_string(HTML_FORM, sucursales=sucs_usuario, incidencias=INCIDENCIAS_REF, fecha_hoy=hoy, meses_con_datos=mc, usuario=session.get('user'))

HTML_LOGIN = """<!DOCTYPE html><html><head><meta charset="UTF-8"><style>:root { --azul: #0844a4; } body { font-family: sans-serif; background: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; } .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center; border-top: 8px solid var(--azul); width: 300px; } input { width: 100%; padding: 12px; margin: 10px 0; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; } button { width: 100%; padding: 12px; background: var(--azul); color: white; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; }</style></head><body><div class="card"><h2>LUXOR</h2><form method="POST"><input type="text" name="user" placeholder="Usuario"><input type="password" name="pass" placeholder="Clave"><button>INGRESAR</button></form></div></body></html>"""

HTML_FORM = """
<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
:root { --azul: #0844a4; --amarillo: #F9D908; --fondo: #f4f7f6; }
body { font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; font-size: 11px; }
.header { background: white; padding: 10px 30px; display: flex; justify-content: space-between; align-items: center; border-bottom: 5px solid var(--amarillo); }
.tabs { display: flex; justify-content: center; background: #ddd; border-bottom: 1px solid #ccc; }
.tab-btn { padding: 12px 25px; cursor: pointer; border: none; background: none; font-weight: bold; color: #555; }
.tab-btn.active { background: white; color: var(--azul); border-top: 4px solid var(--azul); border-bottom: 2px solid white; margin-bottom: -1px; }
.tab-content { display: none; padding: 20px; }
.tab-content.active { display: block; }
.card { background: white; max-width: 850px; margin: auto; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
.grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
label { font-weight: bold; color: #555; }
input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 5px; font-size: 11px; box-sizing: border-box; }
.btn-main { background: var(--azul); color: white; border: none; font-size: 14px; cursor: pointer; font-weight: bold; padding: 12px; width: 100%; margin-top: 10px; border-radius: 4px; }
.btn-sync-top { background: #ff9800; color: white; border: none; font-size: 12px; cursor: pointer; padding: 4px 10px; border-radius: 4px; font-weight: bold; margin-bottom: 5px; display: flex; align-items: center; gap: 5px; width: fit-content; }
table { width: 100%; border-collapse: collapse; background: white; font-size: 9px; margin-top: 10px; }
th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
th { background: var(--azul); color: white; position: sticky; top: 0; }
.btn-icon { border: none; width: 28px; height: 28px; cursor: pointer; border-radius: 4px; color: white; margin-right: 4px; display: inline-flex; align-items: center; justify-content: center; vertical-align: middle; font-size: 14px; }
.btn-del-x { background: #d9534f; color: white; font-size: 9px; padding: 2px 5px; border-radius: 50%; cursor: pointer; border: none; vertical-align: top; margin-left: -8px; font-weight: bold; position: relative; }
#modalImg { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); align-items: center; justify-content: center; }
#modalImg img { max-width: 90%; max-height: 90%; border: 4px solid white; }
.ws-icon { width: 18px; height: 18px; fill: white; }
</style></head><body>
<div id="modalImg" onclick="this.style.display='none'"><img id="imgFull"></div>
<div class="header">
    <div style="display: flex; align-items: center; gap: 10px;">
        <img src="https://tusupermercadoluxor.com/src/assets/img/general/logo_fusion_luxor_euromaxx.png" width="25">
        <h3>DEPARTAMENTO DE FISCALIZACION LUXOR</h3>
    </div>
    <span>{{usuario}} | <a href="/logout">Salir</a></span>
</div>
<div class="tabs">
    <button id="t1" class="tab-btn active" onclick="openTab(event, 'crear')">AÑADIR / EDITAR</button>
    <button id="t2" class="tab-btn" onclick="openTab(event, 'ver')">GESTIONAR</button>
</div>
<div id="crear" class="tab-content active"><div class="card">
<form id="fM">
<input type="hidden" name="ID_EDICION" id="e_id"><input type="hidden" name="F_INC_EXISTENTE" id="e_inc_ex">

<button type="button" class="btn-sync-top" onclick="sincronizar()" title="Sincronizar con Luxor">⚡ SINCRONIZAR</button>

<div class="grid">
<div style="grid-column:span 2"><label>SUCURSAL</label><select name="SUCURSAL" id="e_suc">{% for s in sucursales %}<option>{{s}}</option>{% endfor %}</select></div>
<div><label>FECHA</label><input type="date" name="FECHA" id="e_fec" value="{{fecha_hoy}}"></div>
<div><label>PROVEEDOR</label><input type="text" name="PROVEEDOR" id="e_pro"></div>
<div><label>FACTURA</label><input type="text" name="FACTURA" id="e_fac"></div>
<div><label>RESPONSABLE</label><input type="text" name="RESPONSABLE" id="e_res"></div>
<div style="grid-column:span 2"><label>TIPO FISCALIZACIÓN</label><select name="TIPO_FISC" id="e_tip"><option value="RECEPCION">RECEPCIÓN</option><option value="DEVOLUCION">DEVOLUCIÓN</option><option value="TRANSFERENCIA">TRANSFERENCIA</option></select></div>
<div style="grid-column:span 2"><label>CLASIFICACIÓN MONTO</label><select name="CLASIFICACION_MONTO" id="e_clasm" onchange="toggleM()"><option value="NINGUNA">NINGUNA</option><option value="COBRO">COBRO</option><option value="RECUPERACIÓN">RECUPERACIÓN</option><option value="EXCEDENTES">EXCEDENTES</option></select></div>
<div id="b_m" style="display:none"><label>MONTO $</label><input type="number" name="MONTO" id="e_mon" step="0.01"></div>
<div><label>ID FACTURA COBRADA</label><input type="text" name="F_COBRADA_INPUT" id="e_fco"></div>
<div><label>📸 FOTO FACTURA</label><input type="file" name="FOTO_FILE" accept="image/*"></div>
<div><label>📷 FOTO INCIDENCIA</label><input type="file" name="FOTO_INCIDENCIA_FILE" accept="image/*"></div>
<div style="grid-column:span 2"><label>INCIDENCIA</label><select name="INCIDENCIA" id="e_inc">{% for inc in incidencias %}<option>{{inc}}</option>{% endfor %}</select></div>
<div style="grid-column:span 2"><label>OBSERVACIONES</label><textarea name="OBSERVACION" id="e_obs" rows="2"></textarea></div>
</div>
<button type="submit" class="btn-main">GUARDAR REGISTRO</button>
</form></div></div>

<div id="ver" class="tab-content"><div class="card" style="max-width:98%"><div class="grid" style="grid-template-columns: 1fr 1fr 100px;">
<select id="s_m">{% for m in meses_con_datos %}<option>{{m}}</option>{% endfor %}</select>
<select id="s_s">{% for s in sucursales %}<option>{{s}}</option>{% endfor %}</select>
<button onclick="cargar()" style="background:var(--azul); color:white; font-weight:bold; cursor:pointer;">FILTRAR</button>
</div><div style="overflow-x:auto;"><table id="tabla"><thead><tr>
<th>ACCIONES</th><th>SUCURSAL</th><th>PROVEEDOR</th><th>FACTURA</th><th>FECHA</th><th>TIPO FISC.</th><th>RESPONSABLE</th><th>INCIDENCIA</th><th>TIPO ERROR</th><th>OBS</th><th>MONTO $</th><th>ID COBRO</th><th>F. INC</th>
</tr></thead><tbody></tbody></table></div></div></div>

<script>
const meses_lista = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
function toggleM(){ let v=document.getElementById('e_clasm').value; document.getElementById('b_m').style.display=(v!=='NINGUNA')?'block':'none'; }
function openTab(e, n){ 
    document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active')); 
    document.querySelectorAll('.tab-btn').forEach(b=>(b.classList.remove('active'))); 
    document.getElementById(n).classList.add('active'); e.currentTarget.classList.add('active'); 
    if(n === 'crear' && !document.getElementById('e_id').value) resetBloqueo();
}

function resetBloqueo() {
    const campos = ['e_suc', 'e_fec', 'e_pro', 'e_fac', 'e_res', 'e_tip', 'e_inc'];
    campos.forEach(id => {
        const el = document.getElementById(id);
        if(el.tagName === 'SELECT') { el.style.pointerEvents = 'auto'; el.style.background = 'white'; }
        else { el.readOnly = false; el.style.background = 'white'; }
    });
}

async function sincronizar() {
    try {
        const res = await fetch('/obtener_ultimo_analisis');
        const r = await res.json();
        if(r.status === 'ok' && r.data.PROVEEDOR) {
            document.getElementById('e_pro').value = r.data.PROVEEDOR;
            document.getElementById('e_fac').value = r.data.FACTURA;
            document.getElementById('e_res').value = r.data.RESPONSABLE;
            if(r.data.SUCURSAL) {
                const el = document.getElementById('e_suc');
                for(let i=0; i<el.options.length; i++) {
                    if(el.options[i].text.includes(r.data.SUCURSAL)) { el.selectedIndex = i; break; }
                }
            }
        } else { alert("No hay datos de Luxor capturados."); }
    } catch(e) { alert("Error al conectar."); }
}

document.getElementById('fM').onsubmit=async(e)=>{
    e.preventDefault();
    const fd = new FormData(e.target);
    const res = await fetch('/guardar',{method:'POST', body:fd});
    if((await res.json()).status === 'ok') { alert("¡Guardado!"); location.reload(); }
};

async function cargar(){
    const m = document.getElementById('s_m').value, s = document.getElementById('s_s').value;
    const res = await fetch(`/listar/${m}/${s}`);
    const data = await res.json();
    const b = document.querySelector("#tabla tbody"); b.innerHTML = "";
    data.forEach(r => {
        let f1 = (r['F COBRADA'] && r['F COBRADA']!='SIN_FOTO') ? `<div style="display:inline-block;"><button class="btn-icon" style="background:#5bc0de" onclick="vI('FACTURAS','${r.FECHA}','${r.SUCURSAL}','${r['F COBRADA']}')">📄</button><button class="btn-del-x" onclick="bF('${r.ID}','${r.FECHA}','${r.SUCURSAL}','FACTURAS')">x</button></div>` : '-';
        let f2 = (r['FOTO_INCIDENCIA']) ? `<div style="display:inline-block;"><button class="btn-icon" style="background:#5bc0de" onclick="vI('fotos_incidencias','${r.FECHA}','${r.SUCURSAL}','${r['FOTO_INCIDENCIA']}')">📸</button><button class="btn-del-x" onclick="bF('${r.ID}','${r.FECHA}','${r.SUCURSAL}','fotos_incidencias')">x</button></div>` : '-';
        let wsBtn = `<button class="btn-icon" style="background:#25D366" onclick='sendWA(${JSON.stringify(r)})'><svg class="ws-icon" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217s.231.001.332.005c.109.004.258-.041.404.314l.542 1.312c.058.14.096.303.003.488l-.204.412c-.09.186-.185.31-.08.488.106.177.468.772 1.003 1.248.689.612 1.27.803 1.448.891.178.088.284.073.389-.047.105-.121.447-.52.566-.697.119-.177.239-.148.403-.087s1.042.491 1.222.581c.179.089.299.133.343.209.044.076.044.441-.1.846z"/></svg></button>`;
        b.innerHTML += `<tr><td><button class="btn-icon" style="background:orange" onclick='editar(${JSON.stringify(r)})'>✏️</button><button class="btn-icon" style="background:red" onclick="borrarR('${r.ID}','${r.FECHA}','${r.SUCURSAL}')">🗑️</button>${wsBtn}</td><td>${r.SUCURSAL}</td><td>${r.PROVEEDOR}</td><td>${r.FACTURA}</td><td>${r.FECHA}</td><td>${r['TIPO FISCALIZACIÓN']}</td><td>${r.RESPONSABLE}</td><td>${r.INCIDENCIA}</td><td>${r['TIPO DE ERROR']}</td><td>${r.OBSERVACIÓN}</td><td>${r['MONTO $']}</td><td>${f1}</td><td>${f2}</td></tr>`;
    });
}

function vI(t,f,s,n){ 
    const mes = meses_lista[new Date(f).getUTCMonth()];
    document.getElementById('imgFull').src = `/ver_foto/${t}/${mes}/${s}/${n}`;
    document.getElementById('modalImg').style.display = 'flex';
}
async function bF(id,f,s,t){ if(confirm("¿Quitar foto?")){ await fetch('/borrar_foto',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id,fecha:f,sucursal:s,tipo:t})}); cargar(); } }
async function borrarR(id,f,s){ if(confirm("¿Borrar registro?")){ await fetch('/borrar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id,fecha:f,sucursal:s})}); cargar(); } }
function sendWA(r){
    const t=`REPORTE FISCALIZACIÓN%0A%0A SUCURSAL: ${r.SUCURSAL}%0A FECHA: ${r.FECHA}%0A PROVEEDOR: ${r.PROVEEDOR}%0A FACTURA: ${r.FACTURA}%0A RESPONSABLE: ${r.RESPONSABLE}%0A INCIDENCIA: ${r.INCIDENCIA}`;
    window.open(`https://wa.me/584140511731?text=${t}`,'_blank');
}

function editar(r){
    document.getElementById('e_id').value=r.ID; document.getElementById('e_suc').value=r.SUCURSAL; document.getElementById('e_fec').value=r.FECHA;
    document.getElementById('e_pro').value=r.PROVEEDOR; document.getElementById('e_fac').value=r.FACTURA; document.getElementById('e_res').value=r.RESPONSABLE;
    document.getElementById('e_tip').value=r['TIPO FISCALIZACIÓN']; document.getElementById('e_clasm').value=r['CLASIFICACIÓN MONTO'];
    document.getElementById('e_mon').value=r['MONTO $']; document.getElementById('e_fco').value=r['F COBRADA']; 
    document.getElementById('e_inc').value=r.INCIDENCIA; document.getElementById('e_obs').value=r.OBSERVACIÓN;
    document.getElementById('e_inc_ex').value=r['FOTO_INCIDENCIA'];
    const campos = ['e_suc', 'e_fec', 'e_pro', 'e_fac', 'e_res', 'e_tip', 'e_inc'];
    campos.forEach(id => {
        const el = document.getElementById(id);
        if(el.tagName === 'SELECT') { el.style.pointerEvents = 'none'; el.style.background = '#e9ecef'; }
        else { el.readOnly = true; el.style.background = '#e9ecef'; }
    });
    toggleM(); document.getElementById('t1').click();
}
</script></body></html>
"""

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)