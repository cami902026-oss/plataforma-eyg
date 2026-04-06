"""
ENERGY — Sincronización de Inventario desde OneDrive
=====================================================
Descarga el archivo Excel desde OneDrive vía Microsoft Graph API,
parsea el inventario y actualiza el array INV_RAW en Index.html.

Ejecutado automáticamente por GitHub Actions
Lunes a Viernes a las 3:00 PM Colombia (UTC-5 = 20:00 UTC)
"""

import json
import os
import re
import sys
import io
import urllib.request
import urllib.parse
import urllib.error
import openpyxl

# ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

# Ruta del archivo Excel en OneDrive (relativa a la raíz del drive)
ONEDRIVE_FILE_PATH = 'LOGISTICA/Inventario definitivo_2026.xlsx'

# Email del usuario cuyo OneDrive contiene el archivo
ONEDRIVE_USER_EMAIL = 'andrea.bernal@eygenergygroup.com'

# Mapeo de nombres de columna Excel → nombres internos INV_RAW
# (normalizado a minúsculas para comparación flexible)
COLUMN_MAP = {
    'codigo producto':   'CODIGO PRODUCTO',
    'código producto':   'CODIGO PRODUCTO',
    'codigo':            'CODIGO PRODUCTO',
    'descripcion':       'DESCRIPCION',
    'descripción':       'DESCRIPCION',
    'marca':             'MARCA',
    'ubicacion':         'UBICACIÓN',
    'ubicación':         'UBICACIÓN',
    'stock actual':      'STOCK ACTUAL',
    'stock':             'STOCK ACTUAL',
    'entradas':          'ENTRADAS',
    'salidas':           'SALIDAS',
}


# ─── 1. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id, client_id, client_secret):
    url = ('https://login.microsoftonline.com/' + tenant_id +
           '/oauth2/v2.0/token')
    data = urllib.parse.urlencode({
        'grant_type':    'client_credentials',
        'client_id':     client_id,
        'client_secret': client_secret,
        'scope':         'https://graph.microsoft.com/.default'
    }).encode()
    req = urllib.request.Request(url, data=data, method='POST')
    try:
        with urllib.request.urlopen(req) as resp:
            print('✅ Token obtenido')
            return json.loads(resp.read())['access_token']
    except urllib.error.HTTPError as e:
        print('ERROR auth: ' + str(e.code) + ' — ' + e.read().decode())
        sys.exit(1)


# ─── 2. DESCARGA EXCEL DESDE ONEDRIVE ─────────────────────────────────────────

def download_excel(token, user_email, file_path):
    """Descarga el archivo Excel del OneDrive del usuario vía Graph API."""
    encoded = urllib.parse.quote(file_path, safe='/')
    url = ('https://graph.microsoft.com/v1.0/users/' + user_email +
           '/drive/root:/' + encoded + ':/content')
    req = urllib.request.Request(
        url, headers={'Authorization': 'Bearer ' + token})
    print('📥 Descargando: ' + url)
    try:
        with urllib.request.urlopen(req) as resp:
            data = resp.read()
            print('✅ Excel descargado (' + str(len(data)) + ' bytes)')
            return data
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print('ERROR descargando Excel: ' + str(e.code) + ' — ' + body)
        if e.code == 403:
            print('⚠️  El app necesita permiso Files.Read.All en Azure AD.')
        elif e.code == 404:
            print('⚠️  Archivo no encontrado. Ruta: ' + file_path)
        sys.exit(1)


# ─── 3. PARSEA EL EXCEL ───────────────────────────────────────────────────────

def parse_excel(excel_bytes):
    """Lee el Excel y retorna lista de dicts con los productos."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    print('📋 Hoja activa: ' + str(ws.title))

    # Leer encabezados de la primera fila
    raw_headers = [str(c.value or '').strip() for c in ws[1]]
    headers = []
    for h in raw_headers:
        mapped = COLUMN_MAP.get(h.lower(), h.upper() if h else '')
        headers.append(mapped)

    print('📌 Columnas detectadas: ' + str(headers))

    # Verificar columnas mínimas requeridas
    required = {'CODIGO PRODUCTO', 'DESCRIPCION', 'STOCK ACTUAL'}
    found = set(headers)
    missing = required - found
    if missing:
        print('⚠️  Columnas no encontradas: ' + str(missing))
        print('   Columnas disponibles: ' + str(headers))

    # Parsear filas
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or str(v).strip() == '' for v in row):
            continue  # saltar filas vacías
        record = {}
        for i, val in enumerate(row):
            if i < len(headers) and headers[i]:
                record[headers[i]] = str(val).strip() if val is not None else ''
        # Convertir campos numéricos a entero
        for field in ['STOCK ACTUAL', 'ENTRADAS', 'SALIDAS']:
            if field in record:
                try:
                    record[field] = int(float(record[field])) if record[field] else 0
                except (ValueError, TypeError):
                    record[field] = 0
        records.append(record)

    print('✅ ' + str(len(records)) + ' productos parseados')
    return records


# ─── 4. ACTUALIZA Index.html ──────────────────────────────────────────────────

def update_index_html(records, html_path):
    """Reemplaza el array INV_RAW en Index.html con los datos frescos."""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    new_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    pattern = r'const INV_RAW\s*=\s*\[[\s\S]*?\];'
    replacement = 'const INV_RAW = ' + new_json + ';'

    if not re.search(pattern, content):
        print('ERROR: No se encontró "const INV_RAW" en Index.html')
        sys.exit(1)

    new_content = re.sub(pattern, replacement, content)

    if new_content == content:
        print('ℹ️  Sin cambios en el inventario.')
    else:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print('✅ Index.html actualizado con ' + str(len(records)) + ' productos')


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    tenant_id     = os.environ.get('MS_TENANT_ID', '').strip()
    client_id     = os.environ.get('MS_CLIENT_ID', '').strip()
    client_secret = os.environ['MS_CLIENT_SECRET'].strip()
    sender_email  = os.environ['SENDER_EMAIL'].strip()

    print('🔍 Usuario OneDrive: ' + ONEDRIVE_USER_EMAIL)
    print('📂 Archivo:          ' + ONEDRIVE_FILE_PATH)

    print('🔑 Obteniendo token Microsoft...')
    token = get_access_token(tenant_id, client_id, client_secret)

    excel_bytes = download_excel(token, ONEDRIVE_USER_EMAIL, ONEDRIVE_FILE_PATH)

    records = parse_excel(excel_bytes)
    if not records:
        print('ERROR: El Excel está vacío o no tiene datos.')
        sys.exit(1)

    html_path = os.path.join(os.path.dirname(__file__), '..', 'Index.html')
    update_index_html(records, html_path)

    print('🎉 Sincronización completada exitosamente.')
