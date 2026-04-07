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


# ─── 1. EXTRAE INVENTARIO DESDE Buscador_Inventario_2026.html ────────────────

def extract_inv_from_html(html_path: str) -> list:
    """Extrae el array STOCK_DATA del archivo Buscador_Inventario_2026.html"""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    match = re.search(r'const STOCK_DATA\s*=\s*(\[[\s\S]*?\]);', content)
    if not match:
        print("ERROR: No se encontró 'const STOCK_DATA' en el HTML")
        return []
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError as e:
        print(f"ERROR parseando JSON del inventario: {e}")
        return []


# ─── 2. CALCULA ESTADÍSTICAS ──────────────────────────────────────────────────

def calculate_stats(inv: list) -> dict:
    total     = len(inv)
    con_stock = sum(1 for p in inv if _stock(p) > 0)
    agotados  = sum(1 for p in inv if _stock(p) == 0)
    entradas  = sum(_num(p, 'ENTRADAS') for p in inv)
    salidas   = sum(_num(p, 'SALIDAS')  for p in inv)
    rotacion  = round(salidas / total, 2) if total else 0
    disponib  = round((con_stock / total) * 100, 1) if total else 0
    return dict(total=total, con_stock=con_stock, agotados=agotados,
                entradas=entradas, salidas=salidas, rotacion=rotacion,
                disponib=disponib)

def _stock(p): return int(p.get('STOCK ACTUAL') or 0)
def _num(p, k): return int(p.get(k) or 0)


# ─── 3. GENERA EL EMAIL HTML ──────────────────────────────────────────────────

def _stock_badge(stock):
    if stock == 0:
        return "<span style='background:#fce8e6;color:#c0392b;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>Agotado</span>"
    elif stock <= 3:
        return "<span style='background:#fff8e1;color:#b7770d;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>&#9888; " + str(stock) + "</span>"
    else:
        return "<span style='background:#e6f4ea;color:#1e7e34;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>&#10003; " + str(stock) + "</span>"


def _cat_label(cat: str) -> str:
    m = {'mecanico': 'Material Mecánico', 'electrico': 'Material Eléctrico',
         'instrumentacion': 'Instrumentación', 'instrumentación': 'Instrumentación'}
    return m.get(str(cat).lower().strip(), 'Otros')

def _cat_color(cat: str) -> tuple:
    """Retorna (fondo_header, color_acento) según categoría."""
    m = {
        'mecanico':       ('#0F2B5B', '#1A3A8F'),
        'electrico':      ('#7B3F00', '#C0641E'),
        'instrumentacion':('#1B5E20', '#2E7D32'),
        'instrumentación':('#1B5E20', '#2E7D32'),
    }
    return m.get(str(cat).lower().strip(), ('#37474F', '#546E7A'))

def _build_familia_block(familia: str, productos: list, accent: str) -> str:
    inv_sorted = sorted(productos, key=lambda p: (_stock(p) == 0, -_stock(p)))
    con_stock = sum(1 for p in productos if _stock(p) > 0)
    agotados  = sum(1 for p in productos if _stock(p) == 0)

    rows = ''.join(
        "<tr>"
        "<td style='font-family:monospace;font-size:11px;color:#555;'>" + str(p.get('CODIGO PRODUCTO','')) + "</td>"
        "<td><b style='font-size:12px;'>" + str(p.get('DESCRIPCION','')) + "</b></td>"
        "<td style='color:#666;'>" + str(p.get('MARCA','-')) + "</td>"
        "<td style='color:#666;'>" + str(p.get('UBICACIÓN') or p.get('UBICACION','-')) + "</td>"
        "<td style='text-align:center;'>" + _stock_badge(_stock(p)) + "</td>"
        "<td style='text-align:center;color:#555;font-size:11px;'>+" + str(_num(p,'ENTRADAS')) + " / -" + str(_num(p,'SALIDAS')) + "</td>"
        "</tr>"
        for p in inv_sorted
    )

    badge_parts = str(len(productos)) + " items"
    if agotados:
        badge_parts += " &nbsp;&#9888; " + str(agotados) + " agotados"

    return (
        "<tr style='background:" + accent + "18;'>"
        "<td colspan='6' style='padding:6px 12px;font-size:11px;font-weight:700;"
        "color:" + accent + ";border-left:3px solid " + accent + ";letter-spacing:.5px;'>"
        "&#128281; " + familia +
        " <span style='font-weight:400;color:#888;font-size:10px;'>— " + badge_parts + "</span>"
        "</td></tr>"
        + rows
    )

def _build_category_section(cat: str, productos: list) -> str:
    label = _cat_label(cat)
    hdr_color, accent = _cat_color(cat)
    con_stock = sum(1 for p in productos if _stock(p) > 0)
    agotados  = sum(1 for p in productos if _stock(p) == 0)

    # Agrupar por FAMILIA dentro de la categoría
    familias = {}
    for p in productos:
        fam = str(p.get('FAMILIA','') or 'Sin familia').strip()
        familias.setdefault(fam, []).append(p)

    # Bloques por familia ordenados por nombre
    familia_blocks = ''
    for fam in sorted(familias.keys()):
        familia_blocks += _build_familia_block(fam, familias[fam], accent)

    return (
        "<div style='margin:18px 0 8px;'>"
        "<div style='background:" + hdr_color + ";color:#fff;padding:10px 14px;border-radius:8px 8px 0 0;"
        "display:flex;justify-content:space-between;align-items:center;'>"
        "<span style='font-size:13px;font-weight:700;letter-spacing:1px;'>" + label + "</span>"
        "<span style='font-size:11px;background:rgba(255,255,255,.15);padding:3px 10px;border-radius:10px;'>"
        + str(len(productos)) + " productos &nbsp;|&nbsp; "
        + str(con_stock) + " con stock &nbsp;|&nbsp; "
        + str(agotados) + " agotados</span>"
        "</div>"
        "<table style='width:100%;border-collapse:collapse;font-size:12px;'>"
        "<thead><tr>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Código</th>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Descripción</th>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Marca</th>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Ubic.</th>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Stock</th>"
        "<th style='background:" + accent + ";color:#fff;padding:7px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Ent/Sal</th>"
        "</tr></thead>"
        "<tbody>" + familia_blocks + "</tbody>"
        "</table></div>"
    )

def generate_html(inv: list, stats: dict, date_str: str) -> str:
    # Agrupar por categoría
    CAT_ORDER = ['mecanico', 'electrico', 'instrumentacion', 'otros']
    grupos = {}
    for p in inv:
        cat = str(p.get('CATEGORIA', '') or '').lower().strip()
        if cat not in ('mecanico','electrico','instrumentacion','instrumentación'):
            cat = 'otros'
        grupos.setdefault(cat, []).append(p)

    # Secciones por categoría
    cat_sections = ''
    for cat in CAT_ORDER:
        if cat in grupos and grupos[cat]:
            cat_sections += _build_category_section(cat, grupos[cat])

    # Filas de productos agotados (todas las categorías)
    agotados_rows = ''.join(
        "<tr>"
        "<td style='font-family:monospace;font-size:11px;'>" + str(p.get('CODIGO PRODUCTO','')) + "</td>"
        "<td>" + str(p.get('DESCRIPCION','')) + "</td>"
        "<td>" + str(p.get('MARCA','-')) + "</td>"
        "<td>" + str(p.get('UBICACIÓN') or p.get('UBICACION','-')) + "</td>"
        "<td style='color:#888;font-size:11px;'>" + _cat_label(p.get('CATEGORIA','')) + "</td>"
        "</tr>"
        for p in inv if _stock(p) == 0
    )
    agotados_count = stats['agotados']

    # Pre-calcular condicionales (no pueden ir dentro de {} en f-strings — Python < 3.12)
    pct_agot   = stats['agotados'] / max(stats['total'], 1)
    health_msg = '&#9989; Inventario saludable' if pct_agot < 0.1 else '&#9888; +10% agotado &mdash; revisar reabastecimiento'

    alert_box = ''
    if agotados_count > 0:
        alert_box = ('<div class="alert-box">&#9888;&#65039; <strong>' + str(agotados_count) +
                     ' productos agotados</strong> &mdash; ver secci&oacute;n al final del reporte.</div>')

    if agotados_count > 0:
        agotados_section = (
            '<div class="sec-title" style="color:#c0392b;">&#128308; Productos Agotados (' + str(agotados_count) + ')</div>'
            '<table><thead><tr><th>C&oacute;digo</th><th>Descripci&oacute;n</th><th>Marca</th><th>Ubicaci&oacute;n</th><th>Categor&iacute;a</th></tr></thead>'
            '<tbody>' + agotados_rows + '</tbody></table>'
        )
    else:
        agotados_section = '<div class="meta">&#9989; No hay productos agotados en este momento.</div>'

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
  .alert-box{{background:#fff3cd;border-left:4px solid #e8a020;border-radius:6px;
              padding:10px 14px;font-size:12px;color:#7d5a00;margin:10px 0;}}
  .search-wrap{{margin-bottom:10px;}}
  .search-bar{{width:100%;box-sizing:border-box;padding:9px 14px;font-size:13px;
               border:2px solid #c5d0e8;border-radius:8px;outline:none;}}
  .search-bar:focus{{border-color:#0F2B5B;}}
  .search-tip{{font-size:11px;color:#888;margin:0 0 8px;}}
  table{{width:100%;border-collapse:collapse;font-size:12px;}}
  th{{background:#0F2B5B;color:#fff;padding:8px 10px;text-align:left;font-size:10px;
      text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;}}
  td{{padding:7px 10px;border-bottom:1px solid #eef1f8;vertical-align:middle;}}
  tr:nth-child(even) td{{background:#f8f9ff;}}
  tr.hidden{{display:none;}}
  #sinResultados{{display:none;text-align:center;color:#999;padding:20px;font-size:13px;}}
  .footer{{background:#071525;padding:16px;text-align:center;}}
  .footer p{{margin:3px 0;font-size:11px;color:#8899bb;}}
  .footer strong{{color:#E8A020;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <h1>&#9889; ENERGY</h1>
    <p>Reporte Diario de Inventario &mdash; {date_str}</p>
  </div>
  <div class="gold"></div>
  <div class="body">

    <div class="sec-title">&#128230; Resumen General</div>
    <div class="stats">
      <div class="st blue"><div class="num">{stats['total']}</div><div class="lbl">Total Productos</div></div>
      <div class="st green"><div class="num">{stats['con_stock']}</div><div class="lbl">Con Stock</div></div>
      <div class="st red"><div class="num">{stats['agotados']}</div><div class="lbl">Agotados</div></div>
    </div>

    <div class="sec-title">&#128202; Movimientos Acumulados</div>
    <div class="kd">
      <div class="kd-box ent"><div class="knum">+{stats['entradas']}</div><div class="klbl">Total Entradas</div></div>
      <div class="kd-box sal"><div class="knum">-{stats['salidas']}</div><div class="klbl">Total Salidas</div></div>
    </div>
    <div class="meta">
      &#128260; Rotaci&oacute;n: <span>{stats['rotacion']}</span> sal/prod &nbsp;|&nbsp;
      &#127919; Disponibilidad: <span>{stats['disponib']}%</span> &nbsp;|&nbsp;
      {health_msg}
    </div>

    {alert_box}

    <div class="sec-title">&#128230; Inventario por Categor&iacute;a &mdash; {stats['total']} productos</div>
    {cat_sections}

    {agotados_section}

  </div>
  <div class="footer">
    <p>&#9889; Generado por <strong>ENERGY &mdash; Asistente Administrativo</strong></p>
    <p>E&amp;G Energy Group &middot; Reporte autom&aacute;tico L&ndash;V 5:00 PM Colombia</p>
  </div>
</div>
</body>
</html>"""


# ─── 4. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    endpoints = [
        'https://login.microsoftonline.com/' + tenant_id + '/oauth2/v2.0/token',
        'https://login.microsoftonline.com/' + tenant_id + '/oauth2/token',
    ]
    scopes = ['https://graph.microsoft.com/.default', 'https://graph.microsoft.com/']
    for i, url in enumerate(endpoints):
        data = urllib.parse.urlencode({
            'grant_type':    'client_credentials',
            'client_id':     client_id,
            'client_secret': client_secret,
            'scope' if i == 0 else 'resource': scopes[i]
        }).encode()
        req = urllib.request.Request(url, data=data, method='POST')
        print("🔑 Intentando endpoint " + str(i+1) + ": " + url)
        try:
            with urllib.request.urlopen(req) as resp:
                print("✅ Token obtenido con endpoint " + str(i+1))
                return json.loads(resp.read())['access_token']
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            print("⚠️  Endpoint " + str(i+1) + " falló: " + str(e.code) + " — " + body)
            if i == len(endpoints) - 1:
                print("ERROR: Todos los endpoints fallaron.")
                sys.exit(1)
            print("↪️  Intentando siguiente...")


# ─── 5. ENVÍA EL CORREO VIA GRAPH ─────────────────────────────────────────────

def send_email(token: str, sender: str, recipients: list, subject: str, html_body: str):
    payload = json.dumps({
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': r.strip()}} for r in recipients]
        },
        'saveToSentItems': True
    }).encode('utf-8')
    url = 'https://graph.microsoft.com/v1.0/users/' + sender + '/sendMail'
    req = urllib.request.Request(url, data=payload, method='POST',
        headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req) as resp:
            print("✅ Correo enviado (HTTP " + str(resp.status) + ")")
    except urllib.error.HTTPError as e:
        print("ERROR enviando correo: " + str(e.code) + " — " + e.read().decode())
        sys.exit(1)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    tenant_id     = os.environ.get('MS_TENANT_ID', '9876dbde-5a7f-4139-8c2d-60a4395fd7d6').strip()
    client_id     = os.environ.get('MS_CLIENT_ID', '0c8bd7f5-a027-4da9-9d11-ccc27682b0ec').strip()
    client_secret = os.environ['MS_CLIENT_SECRET'].strip()
    sender_email  = os.environ['SENDER_EMAIL'].strip()
    recipients    = [r.strip() for r in os.environ['RECIPIENT_EMAILS'].split(',')]

    print("🔍 Tenant ID  : '" + tenant_id + "' (len=" + str(len(tenant_id)) + ")")
    print("🔍 Client ID  : '" + client_id + "' (len=" + str(len(client_id)) + ")")
    print("🔍 Secret len : " + str(len(client_secret)) + " chars")
    print("🔍 Sender     : '" + sender_email + "'")

    html_path = os.path.join(os.path.dirname(__file__), '..', 'Buscador_Inventario_2026.html')
    print("📂 Leyendo inventario desde: " + html_path)
    inv = extract_inv_from_html(html_path)
    if not inv:
        print("ERROR: Inventario vacío.")
        sys.exit(1)
    print("✅ " + str(len(inv)) + " productos cargados")

    stats = calculate_stats(inv)
    print("📊 Stats: Total=" + str(stats['total']) + " | Con stock=" + str(stats['con_stock']) + " | Agotados=" + str(stats['agotados']))

    now = datetime.now()
    day      = DAYS_ES.get(now.strftime('%A'), now.strftime('%A'))
    month    = MONTHS_ES.get(now.strftime('%B'), now.strftime('%B'))
    date_str = day + " " + str(now.day) + " de " + month + " de " + str(now.year)

    html_body = generate_html(inv, stats, date_str)
    subject   = "📦 Reporte Inventario E&G — " + date_str

    print("🔑 Obteniendo token Microsoft...")
    token = get_access_token(tenant_id, client_id, client_secret)

    print("📧 Enviando correo a: " + ', '.join(recipients))
    send_email(token, sender_email, recipients, subject, html_body)
    print("🎉 Reporte enviado exitosamente.")
