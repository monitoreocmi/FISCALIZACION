import os
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, session, redirect, url_for

app = Flask(__name__)
app.secret_key = 'luxor_fuerza_bruta_2026'

# --- CONFIGURACIÓN ---
SUCURSALES_LISTA = ["BARQUISIMETO", "CASTAÑO", "CENTRAL", "CIRCULO MILITAR", "BOSQUE", "GUACARA", "IPSFA", "MORA", "VICTORIA", "ACACIAS", "NAGUANAGUA", "SAN DIEGO", "SAN JUAN", "SANTA RITA", "TUCACAS", "VILLAS DE ARAGUA"]

@app.route('/')
def home():
    hoy = datetime.now().strftime('%Y-%m-%d')
    return render_template_string(HTML_COMPLETO, sucursales=SUCURSALES_LISTA, fecha_hoy=hoy)

@app.route('/analizar_codigo', methods=['POST'])
def analizar_codigo():
    try:
        html = request.json.get('html', '')
        # 1. Guardar copia de seguridad de lo que pegaste para que lo revises
        with open("lo_que_pegaste.txt", "w", encoding="utf-8") as f:
            f.write(html)
            
        soup = BeautifulSoup(html, 'html.parser')
        texto_completo = soup.get_text(separator=' ').upper()
        
        datos = {"SUCURSAL": "", "PROVEEDOR": "NO ENCONTRADO", "FACTURA": "NO ENCONTRADA"}
        
        # BUSCAR SUCURSAL (Fuerza bruta: ¿está el nombre en el texto?)
        for s in SUCURSALES_LISTA:
            if s in texto_completo:
                datos["SUCURSAL"] = s
                break
        
        # BUSCAR PROVEEDOR Y FACTURA 
        # Buscamos etiquetas que suelen estar en las tablas de Luxor
        for td in soup.find_all('td'):
            txt_td = td.get_text().strip().upper()
            if "RIF" in txt_td: # Normalmente el proveedor tiene el RIF cerca
                datos["PROVEEDOR"] = td.find_next('td').get_text().strip() if td.find_next('td') else ""
            if "NRO" in txt_td or "FACTURA" in txt_td:
                datos["FACTURA"] = td.find_next('td').get_text().strip() if td.find_next('td') else ""

        return jsonify({"status": "ok", "data": datos, "debug": texto_completo[:500]})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

HTML_COMPLETO = """
<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
    body { font-family: sans-serif; background: #eceff1; padding: 20px; }
    .box { background: white; max-width: 700px; margin: auto; padding: 25px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
    textarea { width: 100%; height: 120px; border: 2px dashed #0844a4; border-radius: 8px; padding: 10px; box-sizing: border-box; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 20px; }
    input, select { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 5px; }
    button { width: 100%; padding: 12px; background: #0844a4; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; margin-top: 10px; }
    .btn-extraer { background: #2e7d32; margin-bottom: 20px; }
</style>
</head><body>
<div class="box">
    <h2>IMPORTADOR LUXOR (FUERZA BRUTA)</h2>
    <label><b>PASO 1:</b> Pega el código fuente (Ctrl+U en Luxor)</label>
    <textarea id="raw_html" placeholder="Pega aquí..."></textarea>
    <button class="btn-extraer" onclick="procesar()">EXTRAER DATOS AHORA</button>

    <hr>
    <div class="grid">
        <div><label>SUCURSAL</label>
            <select id="f_suc">{% for s in sucursales %}<option>{{s}}</option>{% endfor %}</select>
        </div>
        <div><label>FECHA</label><input type="date" id="f_fec" value="{{fecha_hoy}}"></div>
        <div><label>PROVEEDOR</label><input type="text" id="f_pro"></div>
        <div><label>FACTURA</label><input type="text" id="f_fac"></div>
    </div>
    <button onclick="alert('Datos listos para guardar')">GUARDAR EN EXCEL</button>
</div>

<script>
async function procesar() {
    const html = document.getElementById('raw_html').value;
    if(!html) return alert("El cuadro está vacío");
    
    const response = await fetch('/analizar_codigo', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({html: html})
    });
    
    const r = await response.json();
    if(r.status === 'ok') {
        document.getElementById('f_pro').value = r.data.PROVEEDOR;
        document.getElementById('f_fac').value = r.data.FACTURA;
        if(r.data.SUCURSAL) {
            document.getElementById('f_suc').value = r.data.SUCURSAL;
        }
        console.log("Texto analizado:", r.debug);
    }
}
</script>
</body></html>
"""

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)