import os
import sys
import re
import subprocess
import pandas as pd
import time
from bs4 import BeautifulSoup
from datetime import datetime
from functools import wraps

# --- VERIFICACIÓN DE LIBRERÍAS ---
try:
    from flask import Flask, request, jsonify, render_template_string, send_from_directory, session, redirect, url_for, send_file
    from flask_cors import CORS
except ImportError:
    print("\n[!] ERROR: Faltan librerias. Ejecuta: pip install flask flask-cors pandas openpyxl requests beautifulsoup4")
    sys.exit()

app = Flask(__name__)
app.secret_key = 'luxor_full_system_2026'
CORS(app, resources={r"/*": {"origins": "*"}}) # Esto permite peticiones desde cualquier puerto

# --- CONFIGURACIÓN DE RUTAS ---
RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
os.chdir(RUTA_RAIZ)

# IMPORTANTE: Asegúrate de que tu archivo index.html esté en la misma carpeta que este script
RUTA_PANEL_HTML = os.path.join(RUTA_RAIZ, "index.html") 

ULTIMO_ANALISIS = {"SUCURSAL": "", "PROVEEDOR": "", "FACTURA": "", "RESPONSABLE": ""}
PROCESO_ACTIVO = False  # Flag para monitorear el script sincronizar.py

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

# --- NUEVA RUTA PARA EL PANEL ---
# --- MODIFICACIÓN EN RUTA PANEL ---
@app.route('/panel')
@login_required
def servir_panel():
    # En lugar de buscar el archivo localmente, redirigimos al servidor del puerto 80
    # Cambia la IP por la de tu servidor
    return redirect("http://192.168.7.77")

def buscar_archivo(mes, sucursal):
    rd = os.path.join(RUTA_RAIZ, "cuadros", mes.upper())
    if not os.path.exists(rd): return None
    sn = sucursal.upper().replace(" ", "")
    for a in os.listdir(rd):
        if not a.startswith('~$') and a.endswith(".xlsx"):
            nombre_limpio = a.upper().replace(" ", "")
            if sn in nombre_limpio:
                return os.path.join(rd, a)
    return None

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
        ULTIMO_ANALISIS = {"SUCURSAL": sucursal_final, "PROVEEDOR": proveedor, "FACTURA": str(factura_detectada), "RESPONSABLE": responsable}
        return jsonify({"status": "ok", "data": ULTIMO_ANALISIS})
    except Exception as e: return jsonify({"status": "error", "message": str(e)})

@app.route('/obtener_ultimo_analisis', methods=['GET'])
def obtener_ultimo_analisis():
    return jsonify({"status": "ok", "data": ULTIMO_ANALISIS})

@app.route('/status_sincronizacion')
def status_sincronizacion():
    return jsonify({"activo": PROCESO_ACTIVO})

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
    global PROCESO_ACTIVO
    try:
        d = request.form
        suc, fec, ide = d.get('SUCURSAL'), d.get('FECHA'), str(d.get('ID_EDICION') or "").strip()
        mes = MESES_ES[datetime.strptime(fec, '%Y-%m-%d').month]
        path_xlsx = buscar_archivo(mes, suc) or os.path.join(RUTA_RAIZ, "cuadros", mes, f"{suc}.xlsx")
        os.makedirs(os.path.dirname(path_xlsx), exist_ok=True)
        
        fotos_cobro_list = [d.get('F_COB_EXISTENTE')] if d.get('F_COB_EXISTENTE') and d.get('F_COB_EXISTENTE') != "SIN_FOTO" else []
        fotos_inc_list = [d.get('F_INC_EXISTENTE')] if d.get('F_INC_EXISTENTE') else []

        files_cobro = request.files.getlist('FOTO_FILE')
        dest_cobro = os.path.join(RUTA_RAIZ, 'FACTURAS', mes, suc)
        if files_cobro: os.makedirs(dest_cobro, exist_ok=True)
        
        for i, file in enumerate(files_cobro):
            if file and file.filename != '':
                ext = file.filename.rsplit('.', 1)[1].lower()
                n_final = f"COB_{datetime.now().strftime('%H%M%S')}_{i}.{ext}"
                file.save(os.path.join(dest_cobro, n_final))
                fotos_cobro_list.append(n_final)

        files_inc = request.files.getlist('FOTO_INCIDENCIA_FILE')
        dest_inc = os.path.join(RUTA_RAIZ, 'fotos_incidencias', mes, suc)
        if files_inc: os.makedirs(dest_inc, exist_ok=True)

        for i, file in enumerate(files_inc):
            if file and file.filename != '':
                ext = file.filename.rsplit('.', 1)[1].lower()
                n_final = f"INC_{datetime.now().strftime('%H%M%S')}_{i}.{ext}"
                file.save(os.path.join(dest_inc, n_final))
                fotos_inc_list.append(n_final)

        f_cob_val = " ".join(fotos_cobro_list).strip() if fotos_cobro_list else "SIN_FOTO"
        f_inc_name = " ".join(fotos_inc_list).strip() if fotos_inc_list else ""

        if os.path.exists(path_xlsx):
            df = pd.read_excel(path_xlsx, dtype={'FACTURA': str})
        else:
            df = pd.DataFrame()

        if not df.empty and ide: df = df[df['ID'].astype(str) != ide]
        
        nueva = {
            'SUCURSAL': suc, 'PROVEEDOR': d.get('PROVEEDOR'), 'FACTURA': str(d.get('FACTURA')),
            'FECHA': fec, 'TIPO FISCALIZACIÓN': d.get('TIPO_FISC'), 'RESPONSABLE': d.get('RESPONSABLE'),
            'INCIDENCIA': d.get('INCIDENCIA'), 'TIPO DE ERROR': INCIDENCIAS_REF.get(d.get('INCIDENCIA'), "N/A"),
            'OBSERVACIÓN': d.get('OBSERVACION'), 'MONTO $': float(d.get('MONTO') or 0),
            'F COBRADA': f_cob_val, 'CLASIFICACIÓN MONTO': d.get('CLASIFICACION_MONTO'),
            'FOTO_INCIDENCIA': f_inc_name, 'ID': ide if ide else f"ID_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        }
        
        df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
        
        try:
            df.to_excel(path_xlsx, index=False)
            
            # --- EJECUTAR SINCRONIZACIÓN AL GUARDAR ---
            script_maestro = os.path.join(RUTA_RAIZ, "sincronizar.py")
            if os.path.exists(script_maestro):
                PROCESO_ACTIVO = True
                subprocess.run([sys.executable, script_maestro])
                PROCESO_ACTIVO = False
            
        except PermissionError:
            return jsonify({"status": "error", "message": f"El archivo de {suc} está abierto en el servidor. Ciérralo."}), 500

        return jsonify({"status": "ok"})
    except Exception as e: 
        PROCESO_ACTIVO = False
        return jsonify({"status": "error", "message": str(e)}), 500

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
        df = pd.read_excel(ruta, dtype={'FACTURA': str}).fillna("")
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
        df = pd.read_excel(rx, dtype={'FACTURA': str})
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
        df = pd.read_excel(rx, dtype={'FACTURA': str}).fillna("")
        col = 'F COBRADA' if data['tipo'].upper() == 'FACTURAS' else 'FOTO_INCIDENCIA'
        registro = df[df['ID'].astype(str) == str(data['id'])]
        if not registro.empty:
            fotos_actuales = str(registro.iloc[0][col]).split(" ")
            nuevas_fotos = [f for f in fotos_actuales if f != data['nombre']]
            valor_final = " ".join(nuevas_fotos).strip()
            if col == 'F COBRADA' and not valor_final: valor_final = "SIN_FOTO"
            df.loc[df['ID'].astype(str) == str(data['id']), col] = valor_final
            df.to_excel(rx, index=False)
            carpeta = "FACTURAS" if data['tipo'].upper() == 'FACTURAS' else "fotos_incidencias"
            ruta_física = os.path.join(RUTA_RAIZ, carpeta, mes.upper(), data['sucursal'].upper(), data['nombre'])
            if os.path.exists(ruta_física): os.remove(ruta_física)
    return jsonify({"status": "ok"})

@app.route('/')
@login_required
def home():
    hoy = datetime.now()
    hoy_str = hoy.strftime('%Y-%m-%d')
    mes_actual = MESES_ES[hoy.month]
    p_cuadros = os.path.join(RUTA_RAIZ, "cuadros")
    mc = sorted([m for m in os.listdir(p_cuadros) if os.path.isdir(os.path.join(p_cuadros, m))]) if os.path.exists(p_cuadros) else []
    if mes_actual not in mc: mc.append(mes_actual)
    permisos = session.get('sucursales_permitidas', [])
    sucs_usuario = SUCURSALES_FULL if "TODAS" in permisos else [s for s in SUCURSALES_FULL if s in permisos]
    return render_template_string(HTML_FORM, sucursales=sucs_usuario, incidencias=INCIDENCIAS_REF, fecha_hoy=hoy_str, meses_con_datos=mc, mes_actual=mes_actual, usuario=session.get('user'))

HTML_LOGIN = """<!DOCTYPE html><html><head><meta charset="UTF-8"><style>:root { --azul: #0844a4; } body { font-family: sans-serif; background: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; } .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center; border-top: 8px solid var(--azul); width: 300px; } input { width: 100%; padding: 12px; margin: 10px 0; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; } button { width: 100%; padding: 12px; background: var(--azul); color: white; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; }</style></head><body><div class="card"><h2>LUXOR</h2><form method="POST"><input type="text" name="user" placeholder="Usuario"><input type="password" name="pass" placeholder="Clave"><button>INGRESAR</button></form></div></body></html>"""

HTML_FORM = """
<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
:root { --azul: #0844a4; --amarillo: #F9D908; --fondo: #f4f7f6; --verde: #28a745; }
body { font-family: 'Segoe UI', sans-serif; background: var(--fondo); margin: 0; font-size: 11px; }

#progress-container { width: 100%; background: #eee; height: 6px; position: fixed; top: 0; left: 0; z-index: 2000; display: none; }
#progress-bar { width: 0%; height: 100%; background: var(--verde); transition: width 0.3s; }

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

.btn-panel { background: var(--verde); color: white; border: none; font-size: 11px; cursor: pointer; padding: 6px 12px; border-radius: 4px; font-weight: bold; text-decoration: none; display: inline-flex; align-items: center; gap: 5px; }

table { width: 100%; border-collapse: collapse; background: white; font-size: 9px; margin-top: 10px; }
th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
th { background: var(--azul); color: white; position: sticky; top: 0; }
.btn-icon { border: none; width: 28px; height: 28px; cursor: pointer; border-radius: 4px; color: white; margin-right: 4px; display: inline-flex; align-items: center; justify-content: center; vertical-align: middle; font-size: 14px; }
.foto-item { display: flex; align-items: center; gap: 5px; margin-bottom: 3px; background: #f0f0f0; padding: 2px 5px; border-radius: 3px; }
.btn-del-mini { background: #d9534f; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 8px; padding: 2px 4px; }
#modalImg { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); align-items: center; justify-content: center; }
#modalImg img { max-width: 90%; max-height: 90%; border: 4px solid white; }
.ws-icon { width: 18px; height: 18px; fill: white; }
.foto-mini-link { color: var(--azul); text-decoration: underline; cursor: pointer; flex-grow: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
</style></head><body>

<div id="progress-container"><div id="progress-bar"></div></div>

<div id="modalImg" onclick="this.style.display='none'"><img id="imgFull"></div>
<div class="header">
    <div style="display: flex; align-items: center; gap: 15px;">
        <img src="https://tusupermercadoluxor.com/src/assets/img/general/logo_fusion_luxor_euromaxx.png" width="25">
        <h3 style="margin:0">DEPARTAMENTO DE FISCALIZACION LUXOR</h3>
        <button class="btn-panel" id="btnVerPanel" onclick="irAlPanel()">📊 VER PANEL</button>
    </div>
    <span><strong id="current_user">{{usuario}}</strong> | <a href="/logout">Salir</a></span>
</div>
<div class="tabs">
    <button id="t1" class="tab-btn active" onclick="openTab(event, 'crear')">AÑADIR / EDITAR</button>
    <button id="t2" class="tab-btn" onclick="openTab(event, 'ver')">GESTIONAR</button>
</div>
<div id="crear" class="tab-content active"><div class="card">
<form id="fM">
<input type="hidden" name="ID_EDICION" id="e_id">
<input type="hidden" name="F_COB_EXISTENTE" id="e_cob_ex">
<input type="hidden" name="F_INC_EXISTENTE" id="e_inc_ex">

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
<div style="grid-column:span 2; border: 1px dashed #ccc; padding: 10px; border-radius: 5px;">
    <div class="grid">
        <div><label>📸 FOTOS FACTURAS</label><input type="file" name="FOTO_FILE" id="file_f" accept="image/*" multiple></div>
        <div><label>📷 FOTOS INCIDENCIAS</label><input type="file" name="FOTO_INCIDENCIA_FILE" accept="image/*" multiple></div>
    </div>
</div>
<div style="grid-column:span 2"><label>INCIDENCIA</label><select name="INCIDENCIA" id="e_inc">{% for inc in incidencias %}<option>{{inc}}</option>{% endfor %}</select></div>
<div style="grid-column:span 2"><label>OBSERVACIONES</label><textarea name="OBSERVACION" id="e_obs" rows="2"></textarea></div>
</div>
<button type="submit" class="btn-main">GUARDAR REGISTRO</button>
</form></div></div>

<div id="ver" class="tab-content"><div class="card" style="max-width:98%"><div class="grid" style="grid-template-columns: 1fr 1fr 100px;">
<select id="s_m">{% for m in meses_con_datos %}<option {% if m == mes_actual %}selected{% endif %}>{{m}}</option>{% endfor %}</select>
<select id="s_s">{% for s in sucursales %}<option>{{s}}</option>{% endfor %}</select>
<button onclick="cargar()" style="background:var(--azul); color:white; font-weight:bold; cursor:pointer;">FILTRAR</button>
</div><div style="overflow-x:auto;"><table id="tabla"><thead><tr>
<th>ACCIONES</th><th>SUCURSAL</th><th>PROVEEDOR</th><th>FACTURA</th><th>FECHA</th><th>TIPO FISC.</th><th>RESPONSABLE</th><th>INCIDENCIA</th><th>TIPO ERROR</th><th>OBS</th><th>MONTO $</th><th>F. FACTURA</th><th>F. INC</th>
</tr></thead><tbody></tbody></table></div></div></div>

<script>
const meses_lista = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// FUNCIÓN ACTUALIZADA EN EL HTML_FORM
function irAlPanel() {
    const mesSeleccionado = document.getElementById('s_m').value;
    // Apuntamos directamente a la raíz del servidor del Panel (Puerto 80)
    const urlPanel = `http://192.168.77.7:80`;
    window.open(urlPanel, '_blank');
}

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
    document.getElementById('fM').reset();
    document.getElementById('e_id').value = "";
    document.getElementById('e_fec').value = "{{fecha_hoy}}";
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

document.getElementById('fM').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.target.querySelector('.btn-main');
    const bar = document.getElementById('progress-bar');
    const container = document.getElementById('progress-container');
    
    btn.disabled = true;
    btn.innerText = "⏳ GUARDANDO Y ACTUALIZANDO PANEL...";
    container.style.display = 'block';
    bar.style.width = '10%';

    const fd = new FormData(e.target);
    
    try {
        let width = 10;
        const interval = setInterval(() => {
            if (width < 95) {
                width += 2;
                bar.style.width = width + '%';
            }
        }, 400);

        const res = await fetch('/guardar', { method: 'POST', body: fd });
        const r = await res.json();

        clearInterval(interval);
        
        if (r.status === 'ok') {
            bar.style.width = '100%';
            setTimeout(() => {
                alert("¡Guardado y Sincronizado exitosamente!");
                location.reload();
            }, 500);
        } else {
            alert("Error: " + r.message);
            btn.disabled = false;
            btn.innerText = "GUARDAR REGISTRO";
            container.style.display = 'none';
        }
    } catch (err) {
        alert("Error de conexión");
        btn.disabled = false;
        container.style.display = 'none';
    }
};

async function cargar(){
    const m = document.getElementById('s_m').value, s = document.getElementById('s_s').value;
    const res = await fetch(`/listar/${m}/${s}`);
    const data = await res.json();
    const b = document.querySelector("#tabla tbody"); b.innerHTML = "";
    data.forEach(r => {
        let f1_html = "-";
        if(r['F COBRADA'] && r['F COBRADA']!='SIN_FOTO') {
            const fotos = String(r['F COBRADA']).split(" ");
            f1_html = fotos.map(f => `<div class="foto-item"><span class="foto-mini-link" onclick="vI('FACTURAS','${r.FECHA}','${r.SUCURSAL}','${f}')">📄 ${f}</span><button class="btn-del-mini" onclick="bF('${r.ID}','${r.FECHA}','${r.SUCURSAL}','FACTURAS','${f}')">x</button></div>`).join("");
        }
        let f2_html = "-";
        if(r['FOTO_INCIDENCIA']) {
            const fotos_inc = String(r['FOTO_INCIDENCIA']).split(" ");
            f2_html = fotos_inc.map(f => `<div class="foto-item"><span class="foto-mini-link" onclick="vI('fotos_incidencias','${r.FECHA}','${r.SUCURSAL}','${f}')">📸 ${f}</span><button class="btn-del-mini" onclick="bF('${r.ID}','${r.FECHA}','${r.SUCURSAL}','INCIDENCIAS','${f}')">x</button></div>`).join("");
        }
        let wsBtn = `<button class="btn-icon" style="background:#25D366" onclick='sendWA(${JSON.stringify(r)})'><svg class="ws-icon" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217s.231.001.332.005c.109.004.258-.041.404.314l.542 1.312c.058.14.096.303.003.488l-.204.412c-.09.186-.185.31-.08.488.106.177.468.772 1.003 1.248.689.612 1.27.803 1.448.891.178.088.284.073.389-.047.105-.121.447-.52.566-.697.119-.177.239-.148.403-.087s1.042.491 1.222.581c.179.089.299.133.343.209.044.076.044.441-.1.846z"/></svg></button>`;
        b.innerHTML += `<tr><td><button class="btn-icon" style="background:orange" onclick='editar(${JSON.stringify(r)})'>✏️</button><button class="btn-icon" style="background:red" onclick="borrarR('${r.ID}','${r.FECHA}','${r.SUCURSAL}')">🗑️</button>${wsBtn}</td><td>${r.SUCURSAL}</td><td>${r.PROVEEDOR}</td><td>${r.FACTURA}</td><td>${r.FECHA}</td><td>${r['TIPO FISCALIZACIÓN']}</td><td>${r.RESPONSABLE}</td><td>${r.INCIDENCIA}</td><td>${r['TIPO DE ERROR']}</td><td>${r.OBSERVACIÓN}</td><td>${r['MONTO $']}</td><td>${f1_html}</td><td>${f2_html}</td></tr>`;
    });
}

function vI(t,f,s,n){ 
    const mes = meses_lista[new Date(f).getUTCMonth()];
    document.getElementById('imgFull').src = `/ver_foto/${t}/${mes}/${s}/${n}`;
    document.getElementById('modalImg').style.display = 'flex';
}

async function bF(id,f,s,t,nombre){ 
    if(confirm(`¿Borrar la foto ${nombre}?`)){ 
        const res = await fetch('/borrar_foto',{
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body:JSON.stringify({id, fecha:f, sucursal:s, tipo:t, nombre:nombre})
        });
        if((await res.json()).status === 'ok') cargar(); 
    } 
}

async function borrarR(id,f,s){ if(confirm("¿Borrar registro?")){ await fetch('/borrar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id,fecha:f,sucursal:s})}); cargar(); } }

function sendWA(r){
    let mensaje = `REPORTE FISCALIZACIÓN%0A%0A`;
    mensaje += ` FECHA: ${r.FECHA}%0A`;
    mensaje += ` TIPO: ${r['TIPO FISCALIZACIÓN']}%0A`;
    mensaje += ` PROVEEDOR: ${r.PROVEEDOR}%0A`;
    mensaje += ` FACTURA: ${r.FACTURA}%0A`;
    mensaje += ` RESPONSABLE: ${r.RESPONSABLE}%0A`;
    let montoNum = parseFloat(r['MONTO $']);
    if(!isNaN(montoNum) && montoNum > 0) {
        mensaje += ` MONTO: $${montoNum.toFixed(2)}%0A`;
    }
    mensaje += ` INCIDENCIA: ${r.INCIDENCIA}%0A`;
    mensaje += ` OBSERVACIONES: ${r.OBSERVACIÓN}`;
    const url = `https://wa.me/584140511731?text=${mensaje}`;
    window.open(url, '_blank');
}

function editar(r){
    const userAct = document.getElementById('current_user').innerText;
    document.getElementById('e_id').value=r.ID; document.getElementById('e_suc').value=r.SUCURSAL; document.getElementById('e_fec').value=r.FECHA;
    document.getElementById('e_pro').value=r.PROVEEDOR; document.getElementById('e_fac').value=r.FACTURA; document.getElementById('e_res').value=r.RESPONSABLE;
    document.getElementById('e_tip').value=r['TIPO FISCALIZACIÓN']; document.getElementById('e_clasm').value=r['CLASIFICACIÓN MONTO'];
    document.getElementById('e_mon').value=r['MONTO $']; 
    document.getElementById('e_cob_ex').value=r['F COBRADA']; 
    document.getElementById('e_inc_ex').value=r['FOTO_INCIDENCIA'];
    document.getElementById('e_inc').value=r.INCIDENCIA; document.getElementById('e_obs').value=r.OBSERVACIÓN;
    
    const campos = ['e_suc', 'e_fec', 'e_pro', 'e_fac', 'e_res', 'e_tip', 'e_inc'];
    campos.forEach(id => {
        const el = document.getElementById(id);
        if(userAct === 'admin'){
            if(el.tagName === 'SELECT') { el.style.pointerEvents = 'auto'; el.style.background = 'white'; }
            else { el.readOnly = false; el.style.background = 'white'; }
        } else {
            if(el.tagName === 'SELECT') { el.style.pointerEvents = 'none'; el.style.background = '#e9ecef'; }
            else { el.readOnly = true; el.style.background = '#e9ecef'; }
        }
    });
    toggleM(); document.getElementById('t1').click();
}
</script></body></html>
"""

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)