# -*- coding: utf-8 -*-
"""Correos del PORTAL DE VINCULACIÓN (22-oct-2026).

Tres avisos, por Microsoft Graph desde info@eygenergygroup.com (el mismo camino
de los informes y del aviso de OP):

  TIPO=invitacion NIT=...          → al cliente/proveedor: link del portal + código
  TIPO=revision   NIT=... DOC_ID=… → al cliente/proveedor: documento aprobado o
                                     qué debe corregir
  TIPO=barrido (o sin TIPO)        → a E&G: cuentas con documentos nuevos por
                                     revisar (cada 15 min; anti-duplicado en
                                     `avisos_enviados`, tipo portal_nuevos)

Lo dispara la plataforma por el proxy (workflow_dispatch con inputs) o el cron.
Necesita SUPABASE_KEY con permiso de escritura (SUPABASE_SECRET): lee
portal_credenciales (sin política pública) y marca avisos.
"""
import os
import sys
import json
import datetime
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '')
PORTAL = 'https://cami902026-oss.github.io/plataforma-eyg/portal.html'
PLATAFORMA = 'https://cami902026-oss.github.io/plataforma-eyg/Index.html'
REMITENTE = os.environ.get('REMITENTE_OP', 'info@eygenergygroup.com').strip()
# A quién le llega «hay documentos nuevos por revisar»
INTERNOS = ['andrea.bernal@eygenergygroup.com',
            'asistente.administrativo@eygenergygroup.com']
DOCS = {
    'rut':      'RUT (Registro Único Tributario)',
    'camara':   'Certificado de Cámara de Comercio (máximo 30 días de expedido)',
    'cedula':   'Cédula del representante legal (ambas caras)',
    'bancaria': 'Certificación bancaria',
}
TIPO_AVISO = 'portal_nuevos'


def sb(path, method='GET', body=None, prefer=None):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY,
         'Content-Type': 'application/json', 'User-Agent': 'energy-portal-aviso'}
    if prefer:
        h['Prefer'] = prefer
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path, method=method, headers=h,
                               data=json.dumps(body).encode() if body is not None else None)
    with urllib.request.urlopen(r) as resp:
        t = resp.read().decode()
        return json.loads(t) if t.strip() else None


def graph_token():
    d = urllib.parse.urlencode({
        'grant_type': 'client_credentials',
        'client_id': os.environ['MS_CLIENT_ID'],
        'client_secret': os.environ['MS_CLIENT_SECRET'],
        'scope': 'https://graph.microsoft.com/.default'}).encode()
    r = urllib.request.Request(
        'https://login.microsoftonline.com/' + os.environ['MS_TENANT_ID'] + '/oauth2/v2.0/token',
        data=d, method='POST')
    with urllib.request.urlopen(r) as resp:
        return json.loads(resp.read())['access_token']


def enviar(token, asunto, html, para, copia=None):
    if os.environ.get('TEST_OUT'):
        p = os.path.join(os.environ.get('TEMP', '.'), 'portal_aviso_prueba.html')
        with open(p, 'w', encoding='utf-8') as f:
            f.write('<!-- ' + asunto + ' -> ' + ', '.join(para) + ' -->\n' + html)
        print('  (prueba) correo guardado en', p)
        return 202
    msg = {'subject': asunto,
           'body': {'contentType': 'HTML', 'content': html},
           'toRecipients': [{'emailAddress': {'address': a}} for a in para],
           'ccRecipients': [{'emailAddress': {'address': a}} for a in (copia or [])]}
    payload = json.dumps({'message': msg, 'saveToSentItems': True}).encode()
    r = urllib.request.Request(
        'https://graph.microsoft.com/v1.0/users/' + REMITENTE + '/sendMail',
        data=payload, method='POST',
        headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r) as resp:
        return resp.status


def esc(s):
    return (str(s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def fecha(d):
    d = str(d or '')[:10]
    return '/'.join(reversed(d.split('-'))) if d else ''


def plantilla(titulo, cuerpo, boton_txt, boton_url):
    return ('<div style="font-family:Segoe UI,Arial,sans-serif;max-width:620px;margin:0 auto;color:#1b2a4a">'
            '<div style="background:#0F2B5B;padding:18px 24px;border-bottom:3px solid #E8A020">'
            '<div style="color:#fff;font-weight:800;font-size:18px">E&amp;G Energy Group SAS</div>'
            '<div style="color:#E8A020;font-size:12px;letter-spacing:2px">PORTAL DE VINCULACIÓN</div></div>'
            '<div style="padding:22px 24px;background:#fff;border:1px solid #dfe5f0;border-top:0">'
            '<h2 style="margin:0 0 12px;font-size:18px;color:#0F2B5B">' + titulo + '</h2>'
            + cuerpo +
            '<p style="margin:22px 0 8px"><a href="' + boton_url + '" style="background:#2EAA4A;color:#fff;text-decoration:none;'
            'padding:12px 22px;border-radius:8px;font-weight:700;display:inline-block">' + boton_txt + '</a></p>'
            '<p style="font-size:12px;color:#6b7a99">Si el botón no abre, copia este enlace en el navegador:<br>'
            '<a href="' + boton_url + '" style="color:#1A3A8F">' + boton_url + '</a></p>'
            '</div><div style="padding:12px 24px;font-size:11.5px;color:#8899bb">E&amp;G Energy Group SAS · '
            '<a href="mailto:info@eygenergygroup.com" style="color:#8899bb">info@eygenergygroup.com</a></div></div>')


def cuenta(nit):
    r = sb('portal_cuentas?nit=eq.' + urllib.parse.quote(nit) + '&limit=1') or []
    return r[0] if r else None


# ─── 1) Invitación ────────────────────────────────────────────────────────────

def invitacion(token, nit):
    c = cuenta(nit)
    if not c:
        print('! no existe la cuenta', nit)
        return 0
    cr = (sb('portal_credenciales?nit=eq.' + urllib.parse.quote(nit) + '&select=codigo_invitacion&limit=1') or [{}])[0]
    cod = cr.get('codigo_invitacion') or ''
    if not cod:
        print('! la cuenta', nit, 'no tiene código vigente (ya creó su clave): no se envía invitación')
        return 0
    link = PORTAL + '?nit=' + urllib.parse.quote(nit) + '&c=' + urllib.parse.quote(cod)
    que = 'cliente' if c.get('tipo') != 'proveedor' else 'proveedor'
    lista = ''.join('<li style="margin:4px 0">' + esc(v) + '</li>' for v in DOCS.values())
    cuerpo = ('<p>Estimados señores <b>' + esc(c['razon_social']) + '</b>,</p>'
              '<p>Para completar su vinculación como ' + que + ' de E&amp;G Energy Group SAS, los invitamos a cargar sus '
              'documentos en nuestro portal. Es rápido: entran con el NIT, crean una clave y suben los archivos.</p>'
              '<table style="font-size:14px;border-collapse:collapse;margin:14px 0">'
              '<tr><td style="padding:6px 12px 6px 0;color:#6b7a99">NIT</td><td style="padding:6px 0"><b>' + esc(nit) + '</b></td></tr>'
              '<tr><td style="padding:6px 12px 6px 0;color:#6b7a99">Código de invitación</td><td style="padding:6px 0"><b style="font-size:20px;letter-spacing:4px;color:#0F2B5B">' + esc(cod) + '</b></td></tr>'
              '</table>'
              '<p style="margin:0 0 6px"><b>Documentos que necesitamos</b> (PDF, JPG o PNG):</p><ul style="margin:0 0 10px 18px;padding:0">' + lista + '</ul>'
              '<p style="font-size:13px;color:#4a5a7a">1. Abra el portal con el botón. 2. Cree su clave con el código de arriba. '
              '3. Suba cada documento. Le avisaremos por este correo cuando queden aprobados o si algo necesita corrección. '
              'Después podrá entrar siempre con su NIT y su clave para actualizar documentos vencidos.</p>')
    html = plantilla('Vinculación con E&G Energy Group', cuerpo, 'Abrir el portal y crear mi clave', link)
    st = enviar(token, 'Vinculación con E&G Energy Group · documentos de ' + c['razon_social'], html, [c['correo']])
    print('  invitación enviada a', c['correo'], '→', st)
    try:
        sb('portal_cuentas?nit=eq.' + urllib.parse.quote(nit), 'PATCH',
           {'invitacion_enviada_at': datetime.datetime.utcnow().isoformat() + 'Z'}, 'return=minimal')
    except Exception as e:
        print('  ! no se pudo marcar invitacion_enviada_at:', e)
    return 1


# ─── 2) Resultado de la revisión ─────────────────────────────────────────────

def revision(token, nit, doc_id):
    c = cuenta(nit)
    d = (sb('portal_documentos?id=eq.' + str(int(doc_id)) + '&nit=eq.' + urllib.parse.quote(nit) + '&limit=1') or [None])[0]
    if not c or not d:
        print('! no existe la cuenta o el documento', nit, doc_id)
        return 0
    nombre = DOCS.get(d['tipo_doc'], d['tipo_doc'])
    if d['estado'] == 'aprobado':
        docs = sb('portal_documentos?nit=eq.' + urllib.parse.quote(nit) + '&estado=eq.aprobado&select=tipo_doc') or []
        aprobados = set(x['tipo_doc'] for x in docs)
        faltan = [v for k, v in DOCS.items() if k not in aprobados]
        if faltan:
            titulo = 'Documento aprobado'
            cuerpo = ('<p>Estimados señores <b>' + esc(c['razon_social']) + '</b>,</p>'
                      '<p>Revisamos y aprobamos su <b>' + esc(nombre) + '</b> (' + esc(d['nombre_archivo']) + ').</p>'
                      '<p>Para terminar la vinculación aún falta:</p><ul style="margin:0 0 10px 18px">'
                      + ''.join('<li>' + esc(f) + '</li>' for f in faltan) + '</ul>')
            boton = 'Subir lo que falta'
        else:
            titulo = '¡Vinculación completa!'
            cuerpo = ('<p>Estimados señores <b>' + esc(c['razon_social']) + '</b>,</p>'
                      '<p>Aprobamos su <b>' + esc(nombre) + '</b> y con él quedan <b>todos sus documentos aprobados</b>. '
                      'Su vinculación con E&amp;G Energy Group SAS está completa. Gracias.</p>'
                      '<p style="font-size:13px;color:#4a5a7a">Recuerde que el certificado de cámara de comercio vence a los 30 días: '
                      'cuando lo necesitemos actualizado, puede subirlo entrando al portal con su NIT y su clave.</p>')
            boton = 'Ver mis documentos'
        asunto = 'E&G Energy Group · ' + titulo + ' (' + c['razon_social'] + ')'
    else:
        titulo = 'Necesitamos una corrección'
        cuerpo = ('<p>Estimados señores <b>' + esc(c['razon_social']) + '</b>,</p>'
                  '<p>Revisamos su <b>' + esc(nombre) + '</b> (' + esc(d['nombre_archivo']) + ') y necesitamos que lo vuelvan a cargar:</p>'
                  '<div style="background:#fff4f4;border:1px solid #f3c2c2;border-radius:8px;padding:12px 14px;margin:10px 0;color:#9b1c1c">'
                  '✏️ ' + esc(d.get('observacion') or 'Por favor cargue un archivo legible y vigente.') + '</div>'
                  '<p>Entre al portal con su NIT y su clave y suba el archivo corregido en la casilla correspondiente.</p>')
        boton = 'Subir el documento corregido'
        asunto = 'E&G Energy Group · corrección en ' + nombre + ' (' + c['razon_social'] + ')'
    html = plantilla(titulo, cuerpo, boton, PORTAL)
    st = enviar(token, asunto, html, [c['correo']])
    print('  revisión enviada a', c['correo'], '→', st)
    return 1


# ─── 3) Barrido: documentos nuevos para E&G ──────────────────────────────────

def barrido(token):
    pend = sb('portal_documentos?estado=eq.pendiente&select=id,nit,tipo_doc,nombre_archivo,subido_at&order=subido_at.asc&limit=1000') or []
    if not pend:
        print('sin documentos pendientes')
        return 0
    por_nit = {}
    for d in pend:
        por_nit.setdefault(d['nit'], []).append(d)
    try:
        ya = set(str(x['clave']) for x in (sb('avisos_enviados?tipo=eq.' + TIPO_AVISO + '&select=clave&limit=5000') or []))
    except Exception as e:
        print('! no se pudo leer avisos_enviados, no se avisa nada:', e)
        return 0
    n = 0
    for nit, docs in por_nit.items():
        # La clave es el conjunto de pendientes: si suben uno más, es otro aviso.
        clave = nit + ':' + ','.join(str(d['id']) for d in sorted(docs, key=lambda x: x['id']))
        if clave in ya:
            continue
        c = cuenta(nit)
        if not c:
            continue
        # Marcar ANTES de enviar (índice único tipo+clave = sin duplicados entre corridas)
        try:
            sb('avisos_enviados', 'POST', {'tipo': TIPO_AVISO, 'clave': clave, 'detalle': c['razon_social']})
        except Exception as e:
            print('  ! no se pudo marcar', clave, e)
            continue
        filas = ''.join('<li>' + esc(DOCS.get(d['tipo_doc'], d['tipo_doc'])) + ' · ' + esc(d['nombre_archivo']) + ' · ' + fecha(d['subido_at']) + '</li>' for d in docs)
        cuerpo = ('<p><b>' + esc(c['razon_social']) + '</b> (NIT ' + esc(nit) + ', ' + ('proveedor' if c.get('tipo') == 'proveedor' else 'cliente') + ') '
                  'subió ' + str(len(docs)) + ' documento(s) al portal y están <b>por revisar</b>:</p><ul style="margin:0 0 10px 18px">' + filas + '</ul>'
                  '<p style="font-size:13px;color:#4a5a7a">En la plataforma → 🔗 Vinculación → abrir la cuenta → ✅ Aprobar o ✏️ Pedir corrección.</p>')
        html = plantilla('Documentos por revisar', cuerpo, 'Abrir Vinculación en la plataforma', PLATAFORMA)
        try:
            st = enviar(token, 'Portal · ' + c['razon_social'] + ' subió ' + str(len(docs)) + ' documento(s) por revisar', html, INTERNOS)
            print('  aviso interno de', nit, '→', st)
            n += 1
        except Exception as e:
            print('  ! falló el envío de', nit, e)
            try:
                sb('avisos_enviados?tipo=eq.' + TIPO_AVISO + '&clave=eq.' + urllib.parse.quote(clave), 'DELETE')
            except Exception:
                pass
    return n


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    if not SB_KEY:
        raise SystemExit('Falta SUPABASE_KEY')
    tipo = (os.environ.get('TIPO') or 'barrido').strip().lower()
    nit = (os.environ.get('NIT') or '').split('-')[0].strip()
    doc_id = (os.environ.get('DOC_ID') or '').strip()
    token = None if os.environ.get('TEST_OUT') else graph_token()
    print('tipo =', tipo, '| nit =', nit or '-', '| doc =', doc_id or '-')
    if tipo == 'invitacion' and nit:
        invitacion(token, nit)
    elif tipo == 'revision' and nit and doc_id:
        revision(token, nit, doc_id)
    else:
        print('enviados:', barrido(token))
