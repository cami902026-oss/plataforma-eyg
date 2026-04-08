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
ONEDRIVE_FILE_PATH = 'LOGISTICA/Inventario definitivo_2026.xlsm'

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
    'familia':           'FAMILIA',
    'categoria':         'CATEGORIA',
    'categoría':         'CATEGORIA',
    'lote':              'LOTE',
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


# ─── 2. DIAGNÓSTICO: LISTA CONTENIDO DEL DRIVE ────────────────────────────────

def list_drive_folder(token, user_email, folder_path='root/children'):
    """Lista los archivos/carpetas en el OneDrive del usuario para diagnóstico."""
    url = 'https://graph.microsoft.com/v1.0/users/' + user_email + '/drive/' + folder_path
    req = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + token})
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read())
            items = data.get('value', [])
            print('📁 Contenido de /' + folder_path + ':')
            for item in items:
                tipo = '📁' if 'folder' in item else '📄'
                print('   ' + tipo + ' ' + item.get('name', '?'))
            return items
    except urllib.error.HTTPError as e:
        print('⚠️  No se pudo listar ' + folder_path + ': ' + str(e.code) + ' ' + e.read().decode())
        return []


# ─── 3. DESCARGA EXCEL DESDE ONEDRIVE ─────────────────────────────────────────

def download_excel(token, user_email, file_path):
    """Descarga el archivo Excel del OneDrive del usuario vi�a Graph API."""
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

    # Buscar la fila de encabezados (puede no estar en la fila 1)
    # Buscamos en las primeras 15 filas la que tenga más columnas reconocibles
    header_row_idx = 1
    best_score = 0
    for row_idx in range(1, 16):
        row_vals = [str(c.value or '').strip().lower() for c in ws[row_idx]]
        score = sum(1 for v in row_vals if v in COLUMN_MAP)
        non_empty = sum(1 for v in row_vals if v)
        print('   Fila ' + str(row_idx) + ': ' + str([str(c.value or '').strip() for c in ws[row_idx]][:8]) + ' (score=' + str(score) + ')')
        if score > best_score or (score == best_score and non_empty > 2 and score > 0):
            best_score = score
            header_row_idx = row_idx

    print('📌 Fila de encabezados detectada: ' + str(header_row_idx) + ' (score=' + str(best_score) + ')')
    print('📌 Hojas disponibles: ' + str(wb.sheetnames) + ' | Hoja activa: ' + str(ws.title))
    raw_headers = [str(c.value or '').strip() for c in ws[header_row_idx]]
    print('📌 Encabezados RAW del Excel: ' + str(raw_headers))
    headers = []
    for h in raw_headers:
        mapped = COLUMN_MAP.get(h.lower(), h.upper() if h else '')
        headers.append(mapped)

    print('📌 Columnas mapeadas: ' + str(headers))

    # Verificar columnas mínimas requeridas
    required = {'CODIGO PRODUCTO', 'DESCRIPCION', 'STOCK ACTUAL'}
    found = set(headers)
    missing = required - found
    if missing:
        print('⚠️  ALERTA — Columnas no encontradas: ' + str(missing))
        print('   Encabezados del Excel que no coincidieron: ' +
              str([h for h in raw_headers if COLUMN_MAP.get(h.lower()) is None and h]))
    else:
        print('✅ Todas las columnas requeridas encontradas')

    # Parsear filas (empezar después de la fila de encabezados)
    records = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
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


# ─── 4. ACTUALIZA ARCHIVOS HTML ──────────────────────────────────────────────

def update_html_array(records, html_path, array_name):
    """Reemplaza un array JS en un archivo HTML con los datos frescos."""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    new_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    pattern = r'const ' + array_name + r'\s*=\s*\[[\s\S]*?\];'
    replacement = 'const ' + array_name + ' = ' + new_json + ';'

    if not re.search(pattern, content):
        print('⚠️  No se encontró "const ' + array_name + '" en ' + html_path)
        return False

    new_content = re.sub(pattern, replacement, content)
    if new_content == content:
        print('ℹ️  Sin cambios en ' + array_name)
    else:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print('✅ ' + html_path + ' → ' + array_name + ' actualizado con ' + str(len(records)) + ' productos')
    return True

def update_index_html(records, html_path):
    update_html_array(records, html_path, 'INV_RAW')

def update_buscador_html(records, html_path):
    update_html_array(records, html_path, 'STOCK_DATA')


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

    # Diagnóstico: listar raíz y carpeta LOGISTICA
    root_items = list_drive_folder(token, ONEDRIVE_USER_EMAIL)
    logistica_items = [i for i in root_items if i.get('name','').upper() == 'LOGISTICA']
    if logistica_items:
        folder_id = logistica_items[0]['id']
        list_drive_folder(token, ONEDRIVE_USER_EMAIL, 'items/' + folder_id + '/children')
    else:
        print('⚠️  Carpeta LOGISTICA no encontrada en la raíz. Buscando variantes...')
        for item in root_items:
            if 'logis' in item.get('name','').lower():
                print('   Encontrado: ' + item.get('name',''))

    excel_bytes = download_excel(token, ONEDRIVE_USER_EMAIL, ONEDRIVE_FILE_PATH)

    records = parse_excel(excel_bytes)
    if not records:
        print('ERROR: El Excel está vacío o no tiene datos.')
        sys.exit(1)

    base = os.path.dirname(__file__) + '/..'
    update_index_html(records, os.path.join(base, 'Index.html'))
    update_buscador_html(records, os.path.join(base, 'Buscador_Inventario_2026.html'))

    print('🎉 Sincronización completada exitosamente.')
