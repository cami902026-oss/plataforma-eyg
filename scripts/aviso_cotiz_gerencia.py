# -*- coding: utf-8 -*-
"""
AVISO DE COTIZACIÓN EN REVISIÓN DE GERENCIA — E&G Energy Group
===============================================================
Cuando una cotización pasa al estado "Revisión gerencia", Alberto tiene que
mirarla ANTES de que salga al cliente — sobre todo la utilidad. Hasta hoy ese
aviso no existía: la cotización se quedaba esperando a que alguien se acordara
de avisarle por WhatsApp.

CÓMO SALE
Por Microsoft Graph desde info@eygenergygroup.com, el mismo camino de los
informes y del aviso de OP pendiente. No depende de Apps Script.

CUÁNDO SALE
· Al instante: al guardar la cotización, la plataforma le pide al proxy que
  dispare el workflow (30 s de gracia para que la cotización llegue a Supabase).
· De respaldo: el mismo workflow corre cada 15 minutos, así que si el disparo
  falla el aviso sale igual en el siguiente ciclo.

NO DUPLICA
`data/aviso_gerencia_estado.json` guarda las cotizaciones ya avisadas. El
archivo nace SEMBRADO con las 10 que ya estaban en revisión el 25-ago-2026: si
no, al prender esto le habrían entrado 10 correos de golpe a Alberto por
cotizaciones viejas. Si una cotización sale de revisión y vuelve a entrar,
vuelve a avisar (se le quita de la lista al salir).

Prueba local sin enviar:
    TEST_OUT=1 PYTHONIOENCODING=utf-8 python scripts/aviso_cotiz_gerencia.py
"""

import os
import json
import datetime
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '')

ESTADO_REVISION = 'Revisión gerencia'
ESTADO_FILE = 'data/aviso_gerencia_estado.json'

# A quién le llega. Andrea va en copia MIENTRAS SE PRUEBA: cuando diga que ya,
# se borra esta línea y el aviso queda solo para gerencia.
GERENCIA = ['gerenciageneral@eygenergygroup.com']
COPIA_PRUEBA = ['andrea.bernal@eygenergygroup.com']

PLATAFORMA = 'https://cami902026-oss.github.io/plataforma-eyg/Index.html'
MAX_POR_CORRIDA = 10          # tope de cortesía: si algo se desmadra, no inunda


def sb(path):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY,
         'Content-Type': 'application/json', 'User-Agent': 'energy-aviso-cotiz'}
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path, method='GET', headers=h)
    with urllib.request.urlopen(r) as resp:
        t = resp.read().decode()
        return json.loads(t) if t.strip() else None


def money(n):
    try:
        return '$' + format(int(round(float(n or 0))), ',d').replace(',', '.')
    except Exception:
        return '$0'


def num(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


# ─── Microsoft Graph ─────────────────────────────────────────────────────────

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


# ─── El correo ───────────────────────────────────────────────────────────────
# Alberto la abre para mirar la UTILIDAD. Entonces la utilidad va arriba, grande,
# y el detalle por ítem debajo. Sin adjuntos: es informativo, la cotización se
# revisa en la plataforma.
#
# El costo sale de `precio_proveedor` de cada ítem. Cuando faltan precios se
# dice cuántos faltan en vez de mostrar un margen inflado: un margen calculado
# sobre la mitad de los costos miente, y aquí se decide plata con ese número.

def html_aviso(cot, items):
    # Dos limpiezas antes de sumar nada:
    #  · las ALTERNATIVAS no suman (la plataforma hace lo mismo en _cotizSubtotalSinAlts);
    #  · las líneas fantasma —sin valor de venta— no pueden aportar costo. LM1996
    #    tiene el ítem 1 dos veces: una con venta y otra en ceros, y contar las dos
    #    daba un margen de -40% sobre una cotización que va bien.
    reales = [i for i in items if not i.get('alt_de') and num(i.get('v_total')) > 0]
    if not reales:
        reales = [i for i in items if not i.get('alt_de')]
    items = reales
    venta = sum(num(i.get('v_total')) for i in items) or num(cot.get('subtotal')) or num(cot.get('total'))
    con_costo = [i for i in items if i.get('precio_proveedor') not in (None, '', 0)]
    costo = sum(num(i.get('precio_proveedor')) * num(i.get('qty')) for i in items)
    sin_costo = len(items) - len(con_costo)
    completo = (len(items) > 0 and sin_costo == 0)
    utilidad = venta - costo
    margen = (utilidad / venta * 100) if (venta > 0 and costo > 0) else None

    filas = ''
    for i in items[:30]:
        v_tot = num(i.get('v_total'))
        c_tot = num(i.get('precio_proveedor')) * num(i.get('qty'))
        m = ((v_tot - c_tot) / v_tot * 100) if (v_tot > 0 and c_tot > 0) else None
        color = '#6B7A90' if m is None else ('#A32B1E' if m < 10 else '#1F7A38')
        filas += (
            '<tr>'
            '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;">'
            + str(i.get('descripcion') or '')[:70] + '</td>'
            '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;">'
            + str(i.get('qty') or '') + ' ' + str(i.get('udm') or '') + '</td>'
            '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;">'
            + money(v_tot) + '</td>'
            '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;">'
            + (money(c_tot) if c_tot else '<span style="color:#A32B1E;">sin costo</span>') + '</td>'
            '<td style="padding:7px 10px;border-bottom:1px solid #E4E9F0;font-size:13px;text-align:right;white-space:nowrap;color:'
            + color + ';">' + ('%.0f%%' % m if m is not None else '—') + '</td>'
            '</tr>')
    if len(items) > 30:
        filas += ('<tr><td colspan="5" style="padding:7px 10px;font-size:12px;color:#6B7A90;">'
                  'y ' + str(len(items) - 30) + ' ítem(s) más</td></tr>')

    quien = str(cot.get('vendedor') or cot.get('realizada_por') or '—')
    aviso_costos = ''
    if sin_costo:
        aviso_costos = ('<p style="font-size:12.5px;color:#A32B1E;margin:0 0 14px;">'
                        '⚠️ <b>' + str(sin_costo) + ' de ' + str(len(items)) + ' ítems no tienen precio de proveedor '
                        'registrado.</b> La utilidad de abajo solo cuenta los que sí lo tienen.</p>')

    bloque_margen = ''
    if margen is not None:
        bloque_margen = (
            '<td><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">'
            + ('Margen' if completo else 'Margen parcial') + '</div>'
            '<div style="font-size:19px;font-weight:bold;color:'
            + ('#A32B1E' if margen < 10 else '#1F7A38') + ';">' + ('%.1f%%' % margen) + '</div></td>')

    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:760px;">'
        '<div style="background:#1A3A8F;padding:18px 22px;border-radius:3px 3px 0 0;">'
        '<div style="font-size:11px;color:#B9CBF0;letter-spacing:2px;font-weight:bold;">E&amp;G ENERGY GROUP</div>'
        '<div style="font-size:22px;color:#fff;font-weight:bold;padding-top:5px;">'
        + str(cot.get('id')) + ' espera tu revisión</div>'
        '<div style="font-size:13px;color:#D6E1F7;padding-top:5px;">'
        + str(cot.get('cliente') or '—') + ' · la hizo ' + quien
        + (' · vence ' + str(cot.get('fecha_venc')) if cot.get('fecha_venc') else '') + '</div></div>'
        '<div style="border:1px solid #D5DCE6;border-top:none;padding:20px 22px;">'
        '<p style="font-size:14px;color:#0F1B2D;margin:0 0 16px;">'
        'La pasaron a <b>revisión de gerencia</b>. Mira la utilidad antes de que salga al cliente.</p>'
        + aviso_costos +
        '<table cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr>'
        '<td style="padding-right:26px;"><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Venta</div>'
        '<div style="font-size:19px;font-weight:bold;color:#0F1B2D;">' + money(venta) + '</div></td>'
        '<td style="padding-right:26px;"><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Costo</div>'
        '<div style="font-size:19px;font-weight:bold;color:#0F1B2D;">' + money(costo) + '</div></td>'
        '<td style="padding-right:26px;"><div style="font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Utilidad</div>'
        '<div style="font-size:19px;font-weight:bold;color:#0F1B2D;">' + money(utilidad) + '</div></td>'
        + bloque_margen +
        '</tr></table>'
        '<p style="font-size:13px;color:#3D4C63;margin:0 0 12px;">' + str(len(items)) + ' ítems</p>'
        '<table cellpadding="0" cellspacing="0" width="100%" style="border:1px solid #D5DCE6;border-radius:3px;">'
        '<tr style="background:#F6F8FB;">'
        '<th align="left" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Material</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Cant.</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Venta</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Costo</th>'
        '<th align="right" style="padding:8px 10px;font-size:10px;color:#6B7A90;text-transform:uppercase;letter-spacing:1px;">Margen</th></tr>'
        + filas + '</table>'
        '<p style="margin:20px 0 0;"><a href="' + PLATAFORMA + '" '
        'style="background:#1A3A8F;color:#fff;text-decoration:none;padding:11px 22px;border-radius:4px;'
        'font-size:14px;font-weight:bold;display:inline-block;">Abrir la cotización</a></p>'
        '<p style="font-size:11.5px;color:#6B7A90;margin-top:16px;">'
        'Cotizaciones. Aviso automático; no requiere respuesta.</p>'
        '</div></div>')


def cargar_estado():
    try:
        with open(ESTADO_FILE, encoding='utf-8') as f:
            d = json.load(f)
            return [str(x) for x in (d.get('avisadas') or [])]
    except Exception:
        return []


def guardar_estado(ids):
    with open(ESTADO_FILE, 'w', encoding='utf-8') as f:
        json.dump({'avisadas': sorted(set(ids)),
                   'actualizado': datetime.datetime.utcnow().isoformat(timespec='seconds') + 'Z'},
                  f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    prueba = bool(os.environ.get('TEST_OUT'))
    if not SB_KEY:
        raise SystemExit('Falta SUPABASE_KEY')

    q = ('cotizaciones?estado=eq.' + urllib.parse.quote(ESTADO_REVISION)
         + '&deleted=not.is.true'
         + '&select=id,cliente,contacto,total,subtotal,fecha,fecha_venc,vendedor,realizada_por,updated_at'
         + '&order=updated_at.desc&limit=50')
    try:
        cots = sb(q) or []
    except Exception as e:
        # Algunas filas viejas no traen `deleted`; sin el filtro igual sirve.
        print('Reintento sin filtro deleted:', e)
        cots = sb(q.replace('&deleted=not.is.true', '')) or []

    en_revision = [str(c.get('id')) for c in cots]
    avisadas = cargar_estado()
    # Las que salieron de revisión se olvidan: si vuelven a entrar, vuelve a avisar.
    avisadas = [i for i in avisadas if i in en_revision]
    nuevas = [c for c in cots if str(c.get('id')) not in avisadas]

    print('En revisión de gerencia:', len(cots), '| ya avisadas:', len(avisadas), '| nuevas:', len(nuevas))
    if not nuevas:
        if not prueba:
            guardar_estado(avisadas)
        raise SystemExit(0)

    token = None if prueba else graph_token()
    remitente = os.environ.get('REMITENTE_OP', 'info@eygenergygroup.com').strip()
    enviados = 0

    for c in nuevas[:MAX_POR_CORRIDA]:
        cid = str(c.get('id'))
        try:
            items = sb('cotizacion_items?cotizacion_id=eq.' + urllib.parse.quote(cid)
                       + '&select=item,descripcion,udm,qty,v_unit,v_total,precio_proveedor,marca,proveedor,alt_de'
                       + '&order=item.asc&limit=300') or []
        except Exception as e:
            print('  ! no se pudieron leer los ítems de', cid, e)
            items = []

        html = html_aviso(c, items)
        asunto = ('👔 ' + cid + ' espera tu revisión — ' + str(c.get('cliente') or '')
                  + ' · ' + money(c.get('total')))
        if prueba:
            print('--- (prueba, no se envía)', asunto)
            print(html[:400], '...')
        else:
            try:
                enviar(token, remitente, asunto, html, GERENCIA, COPIA_PRUEBA)
                print('  ✓ enviado', cid)
            except Exception as e:
                print('  ! falló el envío de', cid, e)
                continue
        avisadas.append(cid)
        enviados += 1

    if not prueba:
        guardar_estado(avisadas)
    print('Avisos enviados:', enviados)
