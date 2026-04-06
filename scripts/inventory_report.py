"""
ENERGY — Reporte Diario de Inventario
======================================
Lee el inventario desde Index.html, calcula estadísticas
y envía un correo HTML via Microsoft Graph API.

Ejecutado automáticamente por GitHub Actions
Lunes a Viernes a las 5:00 PM Colombia (UTC-5 = 22:00 UTC)
"""

import json
import re
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime


# ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

DAYS_ES    = {'Monday':'Lunes','Tuesday':'Martes','Wednesday':'Miércoles',
               'Thursday':'Jueves','Friday':'Viernes','Saturday':'Sábado','Sunday':'Domingo'}
MONTHS_ES  = {'January':'enero','February':'febrero','March':'marzo','April':'abril',
               'May':'mayo','June':'junio','July':'julio','August':'agosto',
               'September':'septiembre','October':'octubre','November':'noviembre','December':'diciembre'}


# ─── 1. EXTRAE INVENTARIO DESDE Index.html ────────────────────────────────────

def extract_inv_from_html(html_path: str) -> list:
    """Extrae el array INV_RAW del archivo Index.html"""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    match = re.search(r'const INV_RAW\s*=\s*(\[[\s\S]*?\]);', content)
    if not match:
        print("ERROR: No se encontró 'const INV_RAW' en Index.html")
        return []
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError as e:
        print(f"ERROR parseando JSON del inventario: {e}")
        return []


# ─── 2. CALCULA ESTADÍSTICAS ──────────────────────────────────────────────────

def calculate_stats(inv: list) -> dict:
    total        = len(inv)
    con_stock    = sum(1 for p in inv if _stock(p) > 0)
    agotados     = sum(1 for p in inv if _stock(p) == 0)
    entradas     = sum(_num(p, 'ENTRADAS') for p in inv)
    salidas      = sum(_num(p, 'SALIDAS')  for p in inv)
    rotacion     = round(salidas / total, 2) if total else 0
    disponib     = round((con_stock / total) * 100, 1) if total else 0
    return dict(total=total, con_stock=con_stock, agotados=agotados,
                entradas=entradas, salidas=salidas, rotacion=rotacion,
                disponib=disponib)

def _stock(p): return int(p.get('STOCK ACTUAL') or 0)
def _num(p, k): return int(p.get(k) or 0)

# ─── 3. GENERA EL EMAIL HTML ──────────────────────────────────────────────────

def _stock_badge(stock: int) -> str:
    if stock == 0:
        return "<span style='background:#fce8e6;color:#c0392b;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>Agotado</span>"
    elif stock <= 3:
        return f"<span style='background:#fff8e1;color:#b7770d;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>⚠ {stock}</span>"
    else:
        return f"<span style='background:#e6f4ea;color:#1e7e34;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>✓ {stock}</span>"


def generate_html(inv: list, stats: dict, date_str: str) -> str:
    # Ordenar: primero con stock (desc), luego agotados
    inv_sorted = sorted(inv, key=lambda p: (_stock(p) == 0, -_stock(p)))

    # Tabla completa de inventario
    all_rows = ''.join(
        f"<tr>"
        f"<td style='font-family:monospace;font-size:11px;color:#555;'>{p.get('CODIGO PRODUCTO','')}</td>"
        f"<td><b style='font-size:12px;'>{p.get('DESCRIPCION','')}</b></td>"
        f"<td style='color:#666;'>{p.get('MARCA','-')}</td>"
        f"<td style='color:#666;'>{p.get('UBICACIÓN') or p.get('UBICACION','-')}</td>"
        f"<td style='text-align:center;'>{_stock_badge(_stock(p))}</td>"
        f"<td style='text-align:center;color:#555;font-size:11px;'>+{_num(p,'ENTRADAS')} / -{_num(p,'SALIDAS')}</td>"
        f"</tr>"
        for p in inv_sorted
    )

    # Solo agotados para la sección de alerta
    agotados_rows = ''.join(
        f"<tr>"
        f"<td style='font-family:monospace;font-size:11px;'>{p.get('CODIGO PRODUCTO','')}</td>"
        f"<td>{p.get('DESCRIPCION','')}</td>"
        f"<td>{p.get('MARCA','-')}</td>"
        f"<td>{p.get('UBICACIÓN') or p.get('UBICACION','-')}</td>"
        f"</tr>"
        for p in inv if _stock(p) == 0
    )
    agotados_count = stats['agotados']

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4ff;margin:0;padding:16px;}}
  .wrap{{max-width:900px;margin:0 auto;box-shadow:0 8px 32px rgba(0,0,0,.15);border-radius:14px;overflow:hidden;}}
  .hdr{{background:linear-gradient(135deg,#071525 0%,#0F2B5B 60%,#071525 100%);padding:28px;text-align:center;}}
  .hdr h1{{margin:0;font-size:24px;color:#fff;letter-spacing:3px;font-weight:900;}}
  .hdr p{{margin:6px 0 0;color:#8899bb;font-size:13px;}}
  .gold{{height:3px;background:linear-gradient(90deg,#E8A020,transparent);}}
  .body{{background:#fff;padding:24px;}}
  .sec-title{{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;
              color:#0F2B5B;border-bottom:2px solid #E8A020;padding-bottom:6px;margin:22px 0 14px;}}
  .stats{{display:table;width:100%;border-spacing:10px;}}
  .st{{display:table-cell;text-align:center;padding:16px 8px;border-radius:10px;}}
  .st.blue{{background:#e8f0fe;border-left:4px solid #1A3A8F;}}
  .st.green{{background:#e6f4ea;border-left:4px solid #2EAA4A;}}
  .st.red{{background:#fce8e6;border-left:4px solid #e53e3e;}}
  .st .num{{font-size:32px;font-weight:900;color:#0F2B5B;line-height:1;}}
  .st .lbl{{font-size:10px;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;}}
  .kd{{display:table;width:100%;border-spacing:10px;margin:10px 0;}}
  .kd-box{{display:table-cell;text-align:center;padding:14px;border-radius:10px;}}
  .kd-box.ent{{background:#e6f4ea;border:1px solid #2EAA4A;}}
  .kd-box.sal{{background:#fff8e1;border:1px solid #E8A020;}}
  .kd-box .knum{{font-size:26px;font-weight:900;}}
  .kd-box.ent .knum{{color:#2EAA4A;}}
  .kd-box.sal .knum{{color:#E8A020;}}
  .kd-box .klbl{{font-size:10px;color:#666;margin-top:4px;}}
  .meta{{background:#f8f9ff;border-radius:8px;padding:10px 14px;font-size:12px;color:#555;margin:10px 0;}}
  .meta span{{color:#0F2B5B;font-weight:700;}}
  .search-bar{{width:100%;box-sizing:border-box;padding:9px 14px;font-size:13px;border:2px solid #c5d0e8;
              border-radius:8px;margin-bottom:10px;outline:none;}}
  .search-bar:focus{{border-color:#0F2B5B;}}
  .search-tip{{font-size:11px;color:#888;margin-bottom:8px;}}
  table{{width:100%;border-collapse:collapse;font-size:12px;}}
  th{{background:#0F2B5B;color:#fff;padding:8px 10px;text-align:left;font-size:10px;
      text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;}}
  td{{padding:7px 10px;border-bottom:1px solid #eef1f8;vertical-align:middle;}}
  tr:nth-child(even) td{{background:#f8f9ff;}}
  tr.hidden{{display:none;}}
  .footer{{background:#071525;padding:16px;text-align:center;}}
  .footer p{{margin:3px 0;font-size:11px;color:#8899bb;}}
  .footer strong{{color:#E8A020;}}
  .alert-box{{background:#fff3cd;border-left:4px solid #e8a020;border-radius:6px;
              padding:10px 14px;font-size:12px;color:#7d5a00;margin:10px 0;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <h1>⚡ ENERGY</h1>
    <p>Reporte Diario de Inventario — {date_str}</p>
  </div>
  <div class="gold"></div>
  <div class="body">
    <div class="sec-title">📦 Resumen General</div>
    <div class="stats">
      <div class="st blue"><div class="num">{stats['total']}</div><div class="lbl">Total Productos</div></div>
      <div class="st green"><div class="num">{stats['con_stock']}</div><div class="lbl">Con Stock</div></div>
      <div class="st red"><div class="num">{stats['agotados']}</div><div class="lbl">Agotados</div></div>
    </div>
    <div class="sec-title">📊 Movimientos Acumulados</div>
    <div class="kd">
      <div class="kd-box ent"><div class="knum">+{stats['entradas']}</div><div class="klbl">Total Entradas</div></div>
      <div class="kd-box sal"><div class="knum">-{stats['salidas']}</div><div class="klbl">Total Salidas</div></div>
    </div>
    <div class="meta">
      🔄 Rotación: <span>{stats['rotacion']}</span> sal/prod &nbsp;| 
      🎯 Disponibilidad: <span>{stats['disponib']}%</span> &nbsp;| 
      {'\u2705 Inventario saludable' if stats['agotados']/max(stats['total'],1) < 0.1 else '⚠️ +10% agotado — revisar reabastecimiento'}
    </div>
    {('<div class="alert-box">⚠️ <strong>' + str(agotados_count) + ' productos agotados</strong> — ver sección al final.</div>') if agotados_count > 0 else ''}
    <div class="sec-title">🔍 Inventario Completo — {stats['total']} productos</div>
    <p class="search-tip">💡 Escribe abajo para filtrar por código, nombre, marca o ubicación. También puedes usar <strong>Ctrl+F</strong>.</p>
    <input class="search-bar" type="text" id="buscar" placeholder="🔍  Buscar producto, código, marca, ubicación..." oninput="filtrar()" />
    <table id="tablaInv">
      <thead><tr><th>Código</th><th>Descripción</th><th>Marca</th><th>Ubicación</th><th style="text-align:center;">Stock</th><th style="text-align:center;">Ent/Sal</th></tr></thead>
      <tbody id="tbodyInv">{all_rows}</tbody>
    </table>
    <p id="sinResultados" style="display:none;text-align:center;color:#999;padding:20px;">Sin resultados.</p>
    {('<div class="sec-title" style="color:#c0392b;">🔴 Productos Agotados (' + str(agotados_count) + ')</div><table><thead><tr><th>Código</th><th>Descripción</th><th>Marca</th><th>Ubicación</th></tr></thead><tbody>' + agotados_rows + '</tbody></table>') if agotados_count > 0 else '<div class="meta">✅ Sin agotados.</div>'}
  </div>
  <div class="footer">
    <p>⚡ Generado por <strong>ENERGY — Asistente Administrativo</strong></p>
    <p>E&amp;G Energy Group · Reporte automático L-V 5:00 PM Colombia</p>
  </div>
</div>
<script>
function filtrar() {{
  const q = document.getElementById('buscar').value.toLowerCase().trim();
  const filas = document.querySelectorAll('#tbodyInv tr');
  let v = 0;
  filas.forEach(function(f) {{
    if (!q || f.textContent.toLowerCase().includes(q)) {{ f.classList.remove('hidden'); v++; }}
    else {{ f.classList.add('hidden'); }}
  }});
  document.getElementById('sinResultados').style.display = (v===0 && q) ? 'block' : 'none';
}}
</script>
</body>
</html>"""

# ─── 4. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    endpoints = [
        f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token',
        f'https://login.microsoftonline.com/{tenant_id}/oauth2/token',
    ]
    scopes = ['https://graph.microsoft.com/.default', 'https://graph.microsoft.com/']
    for i, url in enumerate(endpoints):
        data = urllib.parse.urlencode({
            'grant_type': 'client_credentials', 'client_id': client_id,
            'client_secret': client_secret,
            'scope' if i == 0 else 'resource': scopes[i]
        }).encode()
        req = urllib.request.Request(url, data=data, method='POST')
        print(f"🔑 Intentando endpoint {i+1}: {url}")
        try:
            with urllib.request.urlopen(req) as resp:
                print(f"✅ Token obtenido con endpoint {i+1}")
                return json.loads(resp.read())['access_token']
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            print(f"⚠️  Endpoint {i+1} falló: {e.code} — {body}")
            if i == len(endpoints) - 1:
                print("ERROR: Todos los endpoints fallaron."); sys.exit(1)
            print("↪️  Intentando siguiente...")


# ─── 5. ENVÍA EL CORREO VIA GRAPH ─────────────────────────────────────────────

def send_email(token: str, sender: str, recipients: list, subject: str, html_body: str):
    payload = json.dumps({
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': r.strip()}} for r in recipients]
        }, 'saveToSentItems': True
    }).encode('utf-8')
    url = f'https://graph.microsoft.com/v1.0/users/{sender}/sendMail'
    req = urllib.request.Request(url, data=payload, method='POST',
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req) as resp:
            print(f"✅ Correo enviado (HTTP {resp.status})")
    except urllib.error.HTTPError as e:
        print(f"ERROR enviando correo: {e.code} — {e.read().decode()}")
        sys.exit(1)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    tenant_id     = os.environ.get('MS_TENANT_ID', '9876dbde-5a7f-4139-8c2d-60a4395fd7d6').strip()
    client_id     = os.environ.get('MS_CLIENT_ID', '0c8bd7f5-a027-4da9-9d11-ccc27682b0ec').strip()
    client_secret = os.environ['MS_CLIENT_SECRET'].strip()
    sender_email  = os.environ['SENDER_EMAIL'].strip()
    recipients    = [r.strip() for r in os.environ['RECIPIENT_EMAILS'].split(',')]

    print(f"🔍 Tenant ID  : '{tenant_id}' (len={len(tenant_id)})")
    print(f"🔍 Client ID  : '{client_id}' (len={len(client_id)})")
    print(f"🔍 Secret len : {len(client_secret)} chars")
    print(f"🔍 Sender     : '{sender_email}'")

    html_path = os.path.join(os.path.dirname(__file__), '..', 'Index.html')
    print(f"📂 Leyendo inventario desde: {html_path}")
    inv = extract_inv_from_html(html_path)
    if not inv: print("ERROR: Inventario vacío."); sys.exit(1)
    print(f"✅ {len(inv)} productos cargados")

    stats = calculate_stats(inv)
    print(f"📊 Stats: Total={stats['total']} | Con stock={stats['con_stock']} | Agotados={stats['agotados']}")

    now = datetime.now()
    day   = DAYS_ES.get(now.strftime('%A'), now.strftime('%A'))
    month = MONTHS_ES.get(now.strftime('%B'), now.strftime('%B'))
    date_str = f"{day} {now.day} de {month} de {now.year}"

    html_body = generate_html(inv, stats, date_str)
    subject   = f"📦 Reporte Inventario E&G — {date_str}"

    print(f"🔑 Obteniendo token Microsoft...")
    token = get_access_token(tenant_id, client_id, client_secret)

    print(f"📧 Enviando correo a: {', '.join(recipients)}")
    send_email(token, sender_email, recipients, subject, html_body)
    print("🎉 Reporte enviado exitosamente.")
