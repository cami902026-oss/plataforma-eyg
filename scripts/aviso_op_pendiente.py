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
# A cada etapa le toca una gente distinta. Mandarle todo a todos es la forma
# más rápida de que dejen de leer los avisos.
COMPRAS = ['asistente.administrativo@eygenergygroup.com',   # Alexandra
           'andrea.bernal@eygenergygroup.com']
GERENCIA = ['gerenciageneral@eygenergygroup.com']
BODEGA = ['bodega@eygenergygroup.com']                       # Yesid

ETAPAS = {
    # estado                (evento que marca,   para,      copia)
    'en_compras':           ('aviso_compras',    COMPRAS,   []),
    'pendiente_aprobacion': ('aviso_enviado',    GERENCIA,  ['andrea.bernal@eygenergygroup.com']),
    'aprobada':             ('aviso_aprobada',   None,      None),   # se decide por ítems
}
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


def enviar(token, remitente, asunto, html, para, copia):
    msg = {'subject': asunto,
           'body': {'contentType': 'HTML', 'content': html},
           'toRecipients': [{'emailAddress': {'address': a}} for a in para],
           'ccRecipients': [{'emailAddress': {'address': a}} for a in (copia or [])]}
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

TEXTOS = {
    'en_compras': (
        'espera precios de compras',
        'Consigue el costo definitivo y mándala a gerencia. '
        'Él aprueba viendo lo que de verdad se va a pagar.',
        'Abrir y cotizar'),
    'pendiente_aprobacion': (
        'espera tu aprobación',
        'Hasta que la apruebes <b>no se compra nada ni baja a bodega</b>.',
        'Abrir y aprobar'),
    'aprobada_compras': (
        'aprobada — hay que comprar',
        'Gerencia ya aprobó. El plan de compras está listo para emitir las órdenes.',
        'Abrir el plan'),
    'aprobada_bodega': (
        'aprobada — hay que alistar',
        'Gerencia ya aprobó. Estos ítems salen de bodega.',
        'Ver qué alistar'),
}


def html_aviso(op, items, dias, clave='pendiente_aprobacion'):
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

    titulo, bajada, boton = TEXTOS.get(clave, TEXTOS['pendiente_aprobacion'])
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:760px;">'
        '<div style="background:#1A3A8F;padding:18px 22px;border-radius:3px 3px 0 0;">'
        '<div style="font-size:11px;color:#B9CBF0;letter-spacing:2px;font-weight:bold;">E&amp;G ENERGY GROUP</div>'
        '<div style="font-size:22px;color:#fff;font-weight:bold;padding-top:5px;">'
        + str(op.get('numero')) + ' ' + titulo + '</div>'
        '<div style="font-size:13px;color:#D6E1F7;padding-top:5px;">'
        + str(op.get('cliente') or '—') + ' · cotización ' + str(op.get('cotizacion_id') or '—')
        + (' · O.C. ' + str(op.get('oc_cliente')) if op.get('oc_cliente') else '') + '</div></div>'
        '<div style="border:1px solid #D5DCE6;border-top:none;padding:20px 22px;">'
        '<p style="font-size:14px;color:#0F1B2D;margin:0 0 16px;">' + bajada
        # Solo se dice si de verdad lleva esperando. `updated_at` viene en UTC y
        # la fecha local puede ir detrás: sin este filtro salía "-1 día(s)".
        + (' Lleva <b>' + str(dias) + ' día(s)</b> esperando.' if (dias and dias > 0) else '') + '</p>'
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
        'font-size:14px;font-weight:bold;display:inline-block;">' + boton + '</a></p>'
        '<p style="font-size:11.5px;color:#6B7A90;margin-top:16px;">'
        'Órdenes de Pedido. Aviso automático; no requiere respuesta.</p>'
        '</div></div>')


if __name__ == '__main__':
    prueba = bool(os.environ.get('TEST_OUT'))
    if not SB_KEY:
        raise SystemExit('Falta SUPABASE_KEY')

    token = None if prueba else graph_token()
    remitente = os.environ.get('REMITENTE_OP', 'info@eygenergygroup.com').strip()
    hoy = datetime.date.today()
    total = 0

    for estado, (evento, para_fijo, copia_fija) in ETAPAS.items():
        ops = sb('ops?deleted=is.false&estado=eq.' + estado
                 + '&select=id,numero,cliente,cotizacion_id,oc_cliente,created_by,updated_at'
                 + '&order=updated_at.asc&limit=50') or []
        if not ops:
            continue
        ids = ','.join(str(o['id']) for o in ops)
        ev = sb('op_eventos?op_id=in.(' + ids + ')&evento=eq.' + evento + '&select=op_id') or []
        ya = set(e['op_id'] for e in ev)
        pend = [o for o in ops if o['id'] not in ya]
        print('%-22s %d en esa etapa, %d sin avisar' % (estado, len(ops), len(pend)))

        for op in pend:
            items = sb('op_items?op_id=eq.%s&select=descripcion,cantidad,udm,origen,proveedor,'
                       'costo_unit,v_total&order=item.asc&limit=500' % op['id']) or []
            dias = None
            try:
                f = str(op.get('updated_at') or '')[:10]
                dias = (hoy - datetime.date(*map(int, f.split('-')))).days
            except Exception:
                pass

            # Al aprobar se parte en dos encargos: compras y bodega. Cada uno
            # recibe SOLO lo suyo — un correo con material que no te toca es
            # ruido, y el ruido es lo que hace que dejen de leerlos.
            if estado == 'aprobada':
                envios = []
                com = [i for i in items if i.get('origen') == 'COMPRA']
                bod = [i for i in items if i.get('origen') == 'BODEGA']
                if com:
                    envios.append(('aprobada_compras', com, COMPRAS, []))
                if bod:
                    envios.append(('aprobada_bodega', bod, BODEGA, ['andrea.bernal@eygenergygroup.com']))
            else:
                envios = [(estado, items, para_fijo, copia_fija)]
            if not envios:
                continue

            # La marca se pone ANTES de enviar. Un aviso que no sale se arregla;
            # uno que sale cada 15 minutos quema el canal.
            try:
                sb('op_eventos', 'POST', {'op_id': op['id'], 'evento': evento,
                                          'detalle': estado, 'usuario': 'sistema'}, 'return=minimal')
            except Exception as e:
                print('   NO SE PUDO MARCAR %s: %s' % (op['numero'], e))
                print('   Se aborta sin enviar para no repetir el aviso cada 15 minutos.')
                print('   Causa probable: SUPABASE_KEY es la clave de solo lectura.')
                print('   Solución: secret SUPABASE_SECRET en GitHub (Settings → Secrets → Actions).')
                raise SystemExit(1)

            for clave, its, para, copia in envios:
                asunto = {'en_compras': 'COTIZAR %s — %s',
                          'pendiente_aprobacion': 'APROBAR %s — %s',
                          'aprobada_compras': 'COMPRAR %s — %s',
                          'aprobada_bodega': 'ALISTAR %s — %s'}.get(clave, '%s — %s') % (
                              op['numero'], op.get('cliente') or 'sin cliente')
                html = html_aviso(op, its, dias, clave)
                if prueba:
                    p = os.path.join(os.environ.get('TEMP', '.'),
                                     'aviso_%s_%s.html' % (op['numero'], clave))
                    open(p, 'w', encoding='utf-8').write(html)
                    print('   [prueba] %-22s -> %s' % (clave, p))
                    continue
                try:
                    enviar(token, remitente, asunto, html, para, copia)
                    total += 1
                    print('   enviado %-22s %s -> %s' % (clave, op['numero'], ', '.join(para)))
                except Exception as e:
                    print('   FALLO EL ENVIO de %s (%s): %s' % (op['numero'], clave, e))
                    try:
                        sb('op_eventos', 'POST', {'op_id': op['id'], 'evento': 'aviso_fallo',
                                                  'detalle': (clave + ': ' + str(e))[:300],
                                                  'usuario': 'sistema'}, 'return=minimal')
                    except Exception:
                        pass

    print('Correos enviados: %d' % total)
