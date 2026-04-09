"""
ENERGY — Reporte Diario de Órdenes de Pedido
=============================================
Lee el informe HTML de OP activas desde la carpeta Energy_bot
y lo envía por correo vía Microsoft Graph API.

Ejecutado automáticamente por Cowork
Lunes a Viernes a las 5:00 PM Colombia (UTC-5 = 22:00 UTC)
"""

import json
import os
import sys
import glob
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime


DAYS_ES   = {'Monday':'Lunes','Tuesday':'Martes','Wednesday':'Miércoles',
              'Thursday':'Jueves','Friday':'Viernes','Saturday':'Sábado','Sunday':'Domingo'}
MONTHS_ES = {'January':'enero','February':'febrero','March':'marzo','April':'abril',
              'May':'mayo','June':'junio','July':'julio','August':'agosto',
              'September':'septiembre','October':'octubre','November':'noviembre','December':'diciembre'}


# ─── 1. LEER CREDENCIALES ─────────────────────────────────────────────────────

def load_config() -> dict:
    """Carga credenciales desde cowork_config.json (solo disponible en entorno local Cowork).
    En GitHub Actions las credenciales vienen de variables de entorno (secrets).
    """
    config_path = os.path.join(os.path.dirname(__file__), 'cowork_config.json')
    if not os.path.exists(config_path):
        return {}
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ─── 2. ENCONTRAR EL INFORME OP MÁS RECIENTE ─────────────────────────────────

def find_op_html(base_dir: str) -> str:
    """Busca el archivo Informe_OP_Activas*.html más reciente en la carpeta."""
    pattern = os.path.join(base_dir, 'Informe_OP_Activas*.html')
    files = glob.glob(pattern)
    if not files:
        print("ERROR: No se encontró ningún archivo Informe_OP_Activas*.html")
        sys.exit(1)
    latest = max(files, key=os.path.getmtime)
    print(f"📄 Usando informe: {os.path.basename(latest)}")
    return latest


# ─── 3. LEER Y PREPARAR EL HTML ───────────────────────────────────────────────

def load_op_html(html_path: str) -> str:
    with open(html_path, 'r', encoding='utf-8') as f:
        return f.read()


def extract_style_and_body(html_content: str):
    """Extrae el bloque <style> y el contenido del <body> del HTML del informe."""
    import re
    style_match = re.search(r'<style[^>]*>(.*?)</style>', html_content, re.DOTALL | re.IGNORECASE)
    style_block = style_match.group(1) if style_match else ''
    body_match = re.search(r'<body[^>]*>(.*?)</body>', html_content, re.DOTALL | re.IGNORECASE)
    body_content = body_match.group(1) if body_match else html_content
    return style_block, body_content


def wrap_for_email(html_content: str, date_str: str) -> str:
    """Envuelve el HTML del informe en un contenedor de email con estilos en <head>."""
    style_block, body_content = extract_style_and_body(html_content)
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
{style_block}
</style>
</head>
<body style="margin:0;padding:0;background:#f4f6fb;">
<div style="max-width:860px;margin:0 auto;padding:16px;">
  <div style="background:#0F2B5B;padding:18px 28px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;">
    <div>
      <div style="font-size:20px;font-weight:900;color:#fff;letter-spacing:1px;">⚡ ENERGY</div>
      <div style="font-size:11px;color:#93c5fd;margin-top:2px;text-transform:uppercase;letter-spacing:1px;">Reporte Diario — Órdenes de Pedido</div>
    </div>
    <div style="font-size:12px;color:#93c5fd;text-align:right;">{date_str}</div>
  </div>
  <div style="height:3px;background:linear-gradient(90deg,#E8A020,transparent);"></div>
  {body_content}
  <div style="background:#071525;padding:14px;border-radius:0 0 12px 12px;text-align:center;margin-top:0;">
    <p style="margin:3px 0;font-size:11px;color:#8899bb;">⚡ Generado por <strong style="color:#E8A020;">ENERGY — Asistente Administrativo</strong></p>
    <p style="margin:3px 0;font-size:11px;color:#8899bb;">E&G Energy Group · Reporte automático L–V 5:00 PM Colombia</p>
  </div>
</div>
</body>
</html>"""


# ─── 4. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    data = urllib.parse.urlencode({
        'grant_type':    'client_credentials',
        'client_id':     client_id,
        'client_secret': client_secret,
        'scope':         'https://graph.microsoft.com/.default'
    }).encode()
    req = urllib.request.Request(url, data=data, method='POST')
    print("🔑 Obteniendo token Microsoft Graph...")
    try:
        with urllib.request.urlopen(req) as resp:
            print("✅ Token obtenido correctamente")
            return json.loads(resp.read())['access_token']
    except urllib.error.HTTPError as e:
        print(f"ERROR obteniendo token: {e.code} — {e.read().decode()}")
        sys.exit(1)


# ─── 5. ENVIAR CORREO VIA GRAPH ────────────────────────────────────────────────

def send_email(token: str, sender: str, recipients: list, subject: str, html_body: str):
    payload = json.dumps({
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': r.strip()}} for r in recipients]
        },
        'saveToSentItems': True
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
    cfg = load_config()

    tenant_id     = os.environ.get('MS_TENANT_ID',     cfg.get('ms_tenant_id',     '')).strip()
    client_id     = os.environ.get('MS_CLIENT_ID',     cfg.get('ms_client_id',     '')).strip()
    client_secret = os.environ.get('MS_CLIENT_SECRET', cfg.get('ms_client_secret', '')).strip()
    sender_email  = os.environ.get('SENDER_EMAIL',     cfg.get('sender_email',     '')).strip()
    recipients    = [r.strip() for r in os.environ.get(
                        'RECIPIENT_EMAILS', cfg.get('recipient_emails', '')).split(',')]

    print(f"🔍 Sender     : {sender_email}")
    print(f"🔍 Destinatarios: {', '.join(recipients)}")

    base_dir = os.path.join(os.path.dirname(__file__), '..')
    html_path = find_op_html(base_dir)
    html_content = load_op_html(html_path)
    print(f"✅ Informe OP cargado ({len(html_content):,} caracteres)")

    now      = datetime.now()
    day      = DAYS_ES.get(now.strftime('%A'), now.strftime('%A'))
    month    = MONTHS_ES.get(now.strftime('%B'), now.strftime('%B'))
    date_str = f"{day} {now.day} de {month} de {now.year}"
    subject  = f"📋 Informe OP Activas E&G — {date_str}"

    email_html = wrap_for_email(html_content, date_str)

    token = get_access_token(tenant_id, client_id, client_secret)
    print(f"📧 Enviando a: {', '.join(recipients)}")
    send_email(token, sender_email, recipients, subject, email_html)
    print("🎉 Informe OP enviado exitosamente.")
