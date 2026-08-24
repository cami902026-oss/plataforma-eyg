# -*- coding: utf-8 -*-
"""
AVISO DE OP PENDIENTE DE APROBACIÓN — E&G Energy Group
=======================================================
Busca las OP que están esperando el visto bueno de gerencia y avisa por correo.

POR QUÉ EXISTE
La plataforma ya manda el aviso al vuelo a un Apps Script, pero ese despliegue
quedó atascado sirviendo código viejo. Esto no depende de Apps Script: usa
Microsoft Graph, el mismo camino que ya envía los informes, y sale de verdad
desde info@eygenergygroup.com (no hace falta configurar alias en Gmail).

NO DUPLICA AVISOS
Cada aviso enviado queda como evento `aviso_enviado` en `op_eventos`, y solo se
notifica lo que no lo tenga. Si algún día el Apps Script vuelve a funcionar y
manda el aviso al instante, la plataforma escribe ese mismo evento y este script
se salta esa OP. Los dos caminos conviven sin pisarse.

Prueba local sin enviar:
    TEST_OUT=1 PYTHONIOENCODING=utf-8 python scripts/aviso_op_pendiente.py
"""

import os
import json
import datetime
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '')
PARA = ['gerenciageneral@eygenergygroup.com']
COPIA = ['andrea.bernal@eygenergygroup.com']
PLATAFORMA = 'https://cami902026-oss.github.io/plataforma-eyg/Index.html'


def sb(path, method='GET', body=None, prefer=None):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY,
         'Content-Type': 'application/json', 'User-Agent': 'energy-aviso-op'}
    if prefer:
        h['Prefer'] = prefer
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path, method=method, headers=h,
                               data=json.dumps(body).encode() if body is not None else None)
    with urllib.request.urlopen(r) as resp:
        t = resp.read().decode()
        return json.loads(t) if t.strip() else None


def money(n):
    try:
        return '$' + format(int(round(float(n or 0))), ',d').replace(',', '.')
    except Exception:
        return '$0'


# ─── Microsoft Graph — el mismo camino de los informes ───────────────────────

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


def enviar(token, remitente, asunto, html):
    msg = {'subject': asunto,
           'body': {'contentType': 'HTML', 'content': html},
           'toRecipients': [{'emailAddress': {'address': a}} for a in PARA],
           'ccRecipients': [{'emailAddress': {'address': a}} for a in COPIA]}
    payload = json.dumps({'message': msg, 'saveToSentItems': True}).encode()
    r = urllib.request.Request(
        'https://graph.microsoft.com/v1.0/users/' + remitente + '/sendMail',
        data=payload, method='POST',
        headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r) as resp:
        return resp.status


# ─── Cuerpo del correo ───────────────────────────────────────────────────────
# Lo que el jefe necesita para decidir sin abrir la plataforma: qué es, cuánto
# cuesta, cuánto se vende y qué queda. El asunto arranca por la acción.

def html_aviso(op, items, dias):
    bod = [i for i in items if i.get('origen') == 'BODEGA']
    com = [i for i in items if i.get('origen') == 'COMPRA']
    venta = sum(float(i.get('v_total') or 0) for i in items)
    costo = sum(float(i.get('costo_unit') or 0) * float(i.get('cantidad') or 0) for i in items)
    margen = ((venta - costo) / venta * 100) if venta > 0 else None

    filas = ''.join(
        '<tr>'
        '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;">' + str(i.get('descripcion') or '')[:70] + '</td>'
        '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;">'
        + str(i.get('cantidad') or '') + ' ' + str(i.get('udm') or '') + '</td>'
        '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:12px;white-space:nowrap;">'
        + ('Bodega' if i.get('origen') == 'BODEGA' else (str(i.get('proveedor') or 'sin proveedor'))) + '</td>'
        '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;">'
        + (money(i.get('costo_unit')) if i.get('costo_unit') is not None else '—') + '</td>'
        '</tr>' for i in items[:25])
    if len(items) > 25:
        filas += ('<tr><td colspan="4" style="padding:7px 10px;font-size:12px;color:#6B7A90;">'
                  'y ' + str(len(items) - 25) + ' ítem(s) más</td></tr>')

    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:760px;">'
        '<div style="background:#1A3A8F;padding:18px 22px;border-radius:3px 3px 0 0;">'
        '<div style="font-size:11px;color:#B9CBF0;letter-spacing:2px;font-weight:bold;">E&amp;G ENERGY GROUP</div>'
        '<div style="font-size:22px;color:#fff;font-weight:bold;padding-top:5px;">' + str(op.get('numero')) + ' espera tu aprobación</div>'
        '<div style="font-size:13px;color:#D6E1F7;padding-top:5px;">'
        + str(op.get('cliente') or '—') + ' · cotización ' + str(op.get('cotizacion_id') or '—')
        + (' · O.C. ' + str(op.get('oc_cliente')) if op.get('oc_cliente') else '') + '</div></div>'
        '<div style="border:1px solid #D5DCE6;border-top:none;padding:20px 22px;">'
        '<p style="font-size:14px;color:#0F1B2D;margin:0 0 16px;">'
        'Hasta que la apruebes <b>no se compra nada ni baja a bodega</b>.'
        + (' Lleva <b>' + str(dias) + ' día(s)</b> esperando.' if dias else '') + '</p>'
        '<table cellpadding="0" cellspacing="0" style="margin-bottom:16px;">'
        '<tr>'
        '<td style="padding-right:26px;"><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Venta</div>'
        '<div style="font-size:19px;font-weight:bold;color:#0F1B2D;">' + money(venta) + '</div></td>'
        '<td style="padding-right:26px;"><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Costo</div>'
        '<div style="font-size:19px;font-weight:bold;color:#0F1B2D;">' + money(costo) + '</div></td>'
        + ('<td><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Margen</div>'
           '<div style="font-size:19px;font-weight:bold;color:' + ('#A32B1E' if margen < 10 else '#1F7A38') + ';">'
           + ('%.1f%%' % margen) + '</div></td>' if margen is not None else '')
        + '</tr></table>'
        '<p style="font-size:13px;color:#3D4C63;margin:0 0 12px;">'
        + str(len(items)) + ' ítems · <b>' + str(len(bod)) + '</b> de bodega · <b>' + str(len(com)) + '</b> por comprar</p>'
        '<table cellpadding="0" cellspacing="0" width="100%" style="border:1px solid #D5DCE6;border-radius:3px;">'
        '<tr style="background:#F6F8FB;">'
        '<th align="left" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Material</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Cant.</th>'
        '<th align="left" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">De dónde</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Costo</th></tr>'
        + filas + '</table>'
        '<p style="margin:20px 0 0;"><a href="' + PLATAFORMA + '" '
        'style="background:#1A3A8F;color:#fff;text-decoration:none;padding:11px 22px;border-radius:4px;'
        'font-size:14px;font-weight:bold;display:inline-block;">Abrir y aprobar</a></p>'
        '<p style="font-size:11.5px;color:#6B7A90;margin-top:16px;">'
        'Órdenes de Pedido → filtro “Por aprobar”. Aviso automático; no requiere respuesta.</p>'
        '</div></div>')


if __name__ == '__main__':
    prueba = bool(os.environ.get('TEST_OUT'))
    if not SB_KEY:
        raise SystemExit('Falta SUPABASE_KEY')

    print('Buscando OP pendientes de aprobación...')
    ops = sb('ops?deleted=is.false&estado=eq.pendiente_aprobacion'
             '&select=id,numero,cliente,cotizacion_id,oc_cliente,created_by,updated_at'
             '&order=updated_at.asc&limit=50') or []
    print('   %d en espera' % len(ops))
    if not ops:
        raise SystemExit(0)

    # Las que ya se avisaron no se vuelven a avisar.
    ids = ','.join(str(o['id']) for o in ops)
    ev = sb('op_eventos?op_id=in.(' + ids + ')&evento=eq.aviso_enviado&select=op_id') or []
    ya = set(e['op_id'] for e in ev)
    pend = [o for o in ops if o['id'] not in ya]
    print('   %d sin avisar' % len(pend))
    if not pend:
        raise SystemExit(0)

    token = None if prueba else graph_token()
    remitente = os.environ.get('REMITENTE_OP', 'info@eygenergygroup.com').strip()
    hoy = datetime.date.today()

    for op in pend:
        items = sb('op_items?op_id=eq.%s&select=descripcion,cantidad,udm,origen,proveedor,'
                   'costo_unit,v_total&order=item.asc&limit=500' % op['id']) or []
        dias = None
        try:
            f = str(op.get('updated_at') or '')[:10]
            dias = (hoy - datetime.date(*map(int, f.split('-')))).days
        except Exception:
            pass
        asunto = 'APROBAR %s — %s' % (op['numero'], op.get('cliente') or 'sin cliente')
        html = html_aviso(op, items, dias)

        if prueba:
            p = os.path.join(os.environ.get('TEMP', '.'), 'aviso_%s.html' % op['numero'])
            open(p, 'w', encoding='utf-8').write(html)
            print('   [prueba] %s -> %s' % (op['numero'], p))
            continue

        # La marca se pone ANTES de enviar. Si la base no deja escribirla, se
        # aborta sin mandar nada: un aviso que no sale se arregla; uno que sale
        # cada 15 minutos durante días quema el canal y nadie vuelve a leerlo.
        try:
            sb('op_eventos', 'POST', {'op_id': op['id'], 'evento': 'aviso_enviado',
                                      'detalle': 'Correo a ' + ', '.join(PARA), 'usuario': 'sistema'},
               'return=minimal')
        except Exception as e:
            print('   NO SE PUDO MARCAR %s: %s' % (op['numero'], e))
            print('   Se aborta sin enviar para no repetir el aviso cada 15 minutos.')
            print('   Causa probable: SUPABASE_KEY es la clave de solo lectura.')
            print('   Solución: agregar el secret SUPABASE_SECRET en GitHub (Settings →')
            print('   Secrets and variables → Actions) con la clave secreta del proyecto.')
            raise SystemExit(1)

        try:
            enviar(token, remitente, asunto, html)
            print('   enviado: %s' % op['numero'])
        except Exception as e:
            # El aviso queda marcado pero no salió: se deja constancia para poder
            # reenviarlo a mano, en vez de reintentar en bucle.
            print('   FALLO EL ENVÍO de %s: %s' % (op['numero'], e))
            try:
                sb('op_eventos', 'POST', {'op_id': op['id'], 'evento': 'aviso_fallo',
                                          'detalle': str(e)[:300], 'usuario': 'sistema'}, 'return=minimal')
            except Exception:
                pass
