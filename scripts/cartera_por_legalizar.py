"""
CARTERA POR LEGALIZAR — E&G Energy Group
=========================================
Informe diario (7 a. m.) para gerencia: lo que el cliente YA nos compró
(cotización Adjudicada) y todavía no se ha facturado. Pedido 27-sep-2026.

De dónde sale cada número
  · Cotizaciones en estado «Adjudicada» (no borradas) en Supabase.
  · Valor = lo ADJUDICADO (valor_adjudicado). Si nunca se registró, se usa el
    subtotal cotizado y la fila dice «valor cotizado»: puede estar inflado.
  · Estado del pedido según su(s) OP: despachada / en proceso / sin OP.
  · Días = desde la adjudicación (o desde la fecha de la cotización si la
    adjudicación no tiene fecha).
  · N° de O.C. del cliente: de la OP o del módulo Procesos O.C. (ordenes.json).

Solo LEE. Prueba local sin enviar:
    TEST_OUT=1 PYTHONIOENCODING=utf-8 python scripts/cartera_por_legalizar.py
"""

import os
import json
import datetime
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '') or 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM'
REMITENTE = os.environ.get('REMITENTE', 'info@eygenergygroup.com')
ANDREA = 'andrea.bernal@eygenergygroup.com'
GERENCIA = 'gerenciageneral@eygenergygroup.com'
# El primer informe va SOLO a Andrea para que lo revise. Cuando ella dé el visto
# bueno se pasa a False y desde ahí va a gerencia con copia a Andrea.
SOLO_ANDREA = True
PLATAFORMA = 'https://cami902026-oss.github.io/plataforma-eyg/Index.html'


def sb(path):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY, 'User-Agent': 'energy-cartera-legalizar'}
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path, headers=h)
    with urllib.request.urlopen(r, timeout=60) as resp:
        t = resp.read().decode()
        return json.loads(t) if t.strip() else []


def todo(path):
    out, off = [], 0
    sep = '&' if '?' in path else '?'
    while True:
        d = sb(path + sep + 'limit=1000&offset=' + str(off)) or []
        out.extend(d)
        if len(d) < 1000 or off > 100000:
            return out
        off += 1000


def money(n):
    return '$' + format(int(round(float(n or 0))), ',d').replace(',', '.')


def fecha(s):
    try:
        return datetime.date.fromisoformat(str(s)[:10])
    except Exception:
        return None


def ocs_del_modulo():
    """cotizacionId → N° de O.C. del cliente, desde el módulo Procesos O.C."""
    m = {}
    try:
        with open('ordenes.json', encoding='utf-8') as f:
            d = json.load(f)
        for o in (d if isinstance(d, list) else []):
            if o.get('deleted') or o.get('estado') == 'cancelado':
                continue
            c, n = o.get('cotizacionId'), o.get('num')
            if c and n:
                m.setdefault(c, []).append(str(n))
    except Exception:
        pass
    return m


def datos():
    hoy = datetime.date.today()
    cots = todo('cotizaciones?estado=eq.Adjudicada&deleted=is.false'
                '&select=id,cliente,fecha,subtotal,valor_adjudicado,adjudicada_at,adjudicacion_at,vendedor&order=id')
    ops = todo('ops?estado=neq.anulada&select=numero,cotizacion_id,estado,oc_cliente,enviada_at')
    por_cot = {}
    for o in ops:
        por_cot.setdefault(o.get('cotizacion_id'), []).append(o)
    oc_mod = ocs_del_modulo()
    filas = []
    for c in cots:
        va = c.get('valor_adjudicado')
        valor = float(va) if va is not None else float(c.get('subtotal') or 0)
        sus = por_cot.get(c['id'], [])
        if not sus:
            est = 'Sin OP'
        elif any(o.get('estado') in ('despachada', 'cerrada') for o in sus):
            est = 'Despachada'
        else:
            est = 'En proceso'
        f = fecha(c.get('adjudicada_at')) or fecha(c.get('adjudicacion_at')) or fecha(c.get('fecha'))
        dias = (hoy - f).days if f else None
        ocs = [o.get('oc_cliente') for o in sus if o.get('oc_cliente')] + oc_mod.get(c['id'], [])
        filas.append({
            'cot': c['id'], 'cliente': c.get('cliente') or '—', 'valor': valor,
            'estimado': va is None, 'estado': est, 'dias': dias,
            'ops': ', '.join(o['numero'] for o in sus) or '—',
            'oc': ', '.join(dict.fromkeys(ocs)) or '—', 'vendedor': c.get('vendedor') or '—'})
    return filas


def html(filas):
    hoy = datetime.date.today().strftime('%d/%m/%Y')
    total = sum(f['valor'] for f in filas)
    orden = ['Despachada', 'En proceso', 'Sin OP']
    nota = {'Despachada': 'Ya se entregó: es lo primero que hay que facturar.',
            'En proceso': 'Tiene OP en curso.',
            'Sin OP': 'Adjudicada pero sin OP creada: confirmar si sigue viva o si ya se facturó y falta cambiar el estado.'}
    tramos = [('Más de 60 días', lambda d: d is not None and d > 60),
              ('30 a 60 días', lambda d: d is not None and 30 < d <= 60),
              ('Menos de 30 días', lambda d: d is None or d <= 30)]
    td = 'style="padding:5px 8px;border-bottom:1px solid #e5e7eb"'
    tdr = 'style="padding:5px 8px;border-bottom:1px solid #e5e7eb;text-align:right;white-space:nowrap"'
    p = ['<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#1f2937">',
         '<h2 style="color:#1a3a8f;margin:0 0 2px">💼 Cartera por legalizar · %s</h2>' % hoy,
         '<p style="margin:0 0 12px;color:#4b5563">Lo adjudicado por el cliente que todavía <b>no se ha facturado</b> '
         '(valores antes de IVA).</p>',
         '<table style="border-collapse:collapse;margin-bottom:14px"><tr>']
    for et in orden:
        fs = [f for f in filas if f['estado'] == et]
        p.append('<td style="padding:8px 14px;border:1px solid #e5e7eb;border-radius:6px">'
                 '<div style="font-size:11px;color:#6b7280">%s</div><div style="font-size:17px;font-weight:700">%s</div>'
                 '<div style="font-size:11px;color:#6b7280">%d cotización(es)</div></td>'
                 % (et, money(sum(f['valor'] for f in fs)), len(fs)))
    p.append('<td style="padding:8px 14px;border:2px solid #1a3a8f"><div style="font-size:11px;color:#6b7280">TOTAL</div>'
             '<div style="font-size:19px;font-weight:800;color:#1a3a8f">%s</div><div style="font-size:11px;color:#6b7280">%d</div></td></tr></table>'
             % (money(total), len(filas)))
    p.append('<p style="margin:0 0 14px;font-size:12px;color:#4b5563">Antigüedad: '
             + ' · '.join('%s <b>%s</b>' % (t, money(sum(f['valor'] for f in filas if fn(f['dias'])))) for t, fn in tramos)
             + '</p>')
    for et in orden:
        fs = sorted([f for f in filas if f['estado'] == et], key=lambda f: -(f['dias'] or 0))
        if not fs:
            continue
        p.append('<h3 style="color:#b45309;margin:14px 0 2px">%s — %s</h3>' % (et, money(sum(f['valor'] for f in fs))))
        p.append('<p style="margin:0 0 6px;font-size:12px;color:#6b7280">%s</p>' % nota[et])
        p.append('<table style="border-collapse:collapse;font-size:12.5px;width:100%%"><tr style="background:#f3f4f6">'
                 '<th %s>Cotización</th><th %s>Cliente</th><th %s>O.C. cliente</th><th %s>OP</th><th %s>Días</th><th %s>Valor</th></tr>'
                 % (td, td, td, td, tdr, tdr))
        for f in fs:
            p.append('<tr><td %s><b>%s</b></td><td %s>%s</td><td %s>%s</td><td %s>%s</td><td %s>%s</td><td %s>%s%s</td></tr>'
                     % (td, f['cot'], td, f['cliente'], td, f['oc'], td, f['ops'],
                        tdr, f['dias'] if f['dias'] is not None else '—', tdr, money(f['valor']),
                        '<br><span style="font-size:10.5px;color:#b45309">valor cotizado</span>' if f['estimado'] else ''))
        p.append('</table>')
    n_est = sum(1 for f in filas if f['estimado'])
    if n_est:
        p.append('<p style="margin-top:12px;font-size:11.5px;color:#b45309">%d cotización(es) dicen «valor cotizado»: '
                 'no tienen registrado cuánto adjudicó el cliente, así que el valor puede ser mayor al real.</p>' % n_est)
    p.append('<p style="margin-top:14px;font-size:11.5px;color:#6b7280">Si una ya se facturó, marcarla '
             '«Facturada» con su número en la plataforma y sale de este informe. '
             '<a href="%s">Abrir la plataforma</a></p></div>' % PLATAFORMA)
    return ''.join(p)


def graph_token():
    d = urllib.parse.urlencode({
        'grant_type': 'client_credentials', 'client_id': os.environ['MS_CLIENT_ID'],
        'client_secret': os.environ['MS_CLIENT_SECRET'],
        'scope': 'https://graph.microsoft.com/.default'}).encode()
    r = urllib.request.Request('https://login.microsoftonline.com/' + os.environ['MS_TENANT_ID']
                               + '/oauth2/v2.0/token', data=d, method='POST')
    with urllib.request.urlopen(r) as resp:
        return json.loads(resp.read())['access_token']


def enviar(asunto, cuerpo, para, copia):
    msg = {'subject': asunto, 'body': {'contentType': 'HTML', 'content': cuerpo},
           'toRecipients': [{'emailAddress': {'address': a}} for a in para],
           'ccRecipients': [{'emailAddress': {'address': a}} for a in copia]}
    r = urllib.request.Request(
        'https://graph.microsoft.com/v1.0/users/' + REMITENTE + '/sendMail',
        data=json.dumps({'message': msg, 'saveToSentItems': True}).encode(), method='POST',
        headers={'Authorization': 'Bearer ' + graph_token(), 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r) as resp:
        return resp.status


def main():
    filas = datos()
    total = sum(f['valor'] for f in filas)
    print('Adjudicadas sin facturar: %d · %s' % (len(filas), money(total)))
    if not filas:
        print('Nada por legalizar: no se manda correo.')
        return
    asunto = '💼 Cartera por legalizar: %s en %d cotización(es)' % (money(total), len(filas))
    cuerpo = html(filas)
    if os.environ.get('TEST_OUT'):
        with open('cartera_legalizar_prueba.html', 'w', encoding='utf-8') as f:
            f.write(cuerpo)
        print('TEST_OUT: no se envió; vista previa en cartera_legalizar_prueba.html')
        return
    if SOLO_ANDREA:
        enviar('[REVISIÓN] ' + asunto, cuerpo, [ANDREA], [])
        print('Enviado SOLO a Andrea (modo revisión)')
    else:
        enviar(asunto, cuerpo, [GERENCIA], [ANDREA])
        print('Enviado a gerencia con copia a Andrea')


if __name__ == '__main__':
    main()
