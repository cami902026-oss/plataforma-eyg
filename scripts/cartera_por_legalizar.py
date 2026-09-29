"""
CARTERA POR LEGALIZAR — E&G Energy Group
=========================================
Informe diario (7 a. m.) para gerencia: lo que el cliente YA nos compró y
todavía no se ha facturado. Pedido 27-sep-2026; rehecho POR ORDEN DE PEDIDO
el 29-sep-2026 y cruzado contra Cartera.

Secciones
  1. ENTREGADO SIN FACTURAR (el número que importa): una fila por OP sin cerrar
     con algo entregado y SIN factura en Cartera. Valor = lo ENTREGADO, no lo
     cotizado: cada remisión de la OP se valora al precio unitario de la
     cotización (item_cot → cotizacion_items), con la cantidad topada a lo
     cotizado; si la remisión no trae ítem, se usa lo despachado en op_items.
     Se toma el mayor de los dos. Días = desde la primera remisión.
  2. FACTURADO EN CARTERA, OP SIN CERRAR: la factura ya existe en Cartera
     (mismo N° de cotización) pero la OP no se cerró. No suma: es trámite.
  3. OP EN CURSO SIN ENTREGAS: pedido vivo, falta despachar. Informativo.
  4. ADJUDICADAS SIN OP: se cruzan con Cartera por N° de cotización, por la
     O.C. del cliente (módulo Procesos O.C.) o por cliente + monto (±2 %).
     Con coincidencia → «posiblemente ya facturada»; sin nada → confirmar.

Solo LEE. Prueba local sin enviar:
    TEST_OUT=1 PYTHONIOENCODING=utf-8 python scripts/cartera_por_legalizar.py
"""

import os
import re
import json
import datetime
import unicodedata
import collections
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '') or 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM'
REMITENTE = os.environ.get('REMITENTE', 'info@eygenergygroup.com')
ANDREA = 'andrea.bernal@eygenergygroup.com'
GERENCIA = 'gerenciageneral@eygenergygroup.com'
# 29-sep-2026: aprobado por Andrea → va a gerencia con copia a ella.
SOLO_ANDREA = False
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


def num(x):
    try:
        return float(x or 0)
    except Exception:
        return 0.0


def clave(s):
    """LM2092 / lm-2092 / 'LM 2092' → LM2092 (para cruzar textos escritos a mano)."""
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


_SOCIETARIA = r'\b(S ?A ?S|S ?A|LTDA|E ?U|BIC|SUCURSAL COLOMBIA(NA)?)\b'


def cliente_norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().upper()
    s = re.sub(r'[^A-Z0-9 ]', ' ', s)
    s = re.sub(_SOCIETARIA, ' ', s)
    return ' '.join(s.split())


def mismo_cliente(a, b):
    """Mismo criterio que Cartera en la plataforma: exacto, normalizado o uno
    contiene al otro (4+ letras)."""
    a, b = cliente_norm(a), cliente_norm(b)
    if not a or not b:
        return False
    return a == b or (len(min(a, b, key=len)) >= 4 and (a in b or b in a))


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


def cot_ids(o):
    """Todas las cotizaciones de una OP (igual que opCotizIds en la plataforma):
    la principal + extra.cotiz_ids + extra.cotiz_extra. Ej.: OP-2026-0064 lleva
    LM1784 y LM1946, y se facturó junta en la 656."""
    out = []
    ex = o.get('extra') or {}
    for v in [o.get('cotizacion_id')] + list(ex.get('cotiz_ids') or []) + [x and x.get('id') for x in (ex.get('cotiz_extra') or [])]:
        t = str(v or '').strip()
        if t and t not in out:
            out.append(t)
    return out


def _cot_items(ids):
    """{cotizacion_id: {item(str): (v_unit, qty)}} — sin alternativas."""
    out = collections.defaultdict(dict)
    ids = sorted(set(i for i in ids if i))
    for k in range(0, len(ids), 40):
        lote = ','.join('"%s"' % i.replace('"', '') for i in ids[k:k + 40])
        for x in todo('cotizacion_items?cotizacion_id=in.(' + urllib.parse.quote(lote, safe=',"()')
                      + ')&select=cotizacion_id,item,v_unit,qty,alt_de'):
            if x.get('alt_de'):
                continue
            out[x['cotizacion_id']][str(x.get('item'))] = (num(x.get('v_unit')), num(x.get('qty')))
    return out


def datos():
    hoy = datetime.date.today()
    ops = [o for o in todo('ops?select=id,numero,cotizacion_id,cliente,estado,oc_cliente,factura,deleted,valor_venta,extra')
           if not o.get('deleted') and o.get('estado') != 'anulada']   # borrador vivo = OP en curso
    abiertas = [o for o in ops if o.get('estado') != 'cerrada' and not o.get('factura')]
    ids_op = set(o['id'] for o in abiertas)
    op_items = collections.defaultdict(list)
    for x in todo('op_items?select=op_id,cantidad,v_unit,v_total,estado,despachada'):
        if x.get('op_id') in ids_op:
            op_items[x['op_id']].append(x)
    nums = set(o['numero'] for o in abiertas)
    rems = collections.defaultdict(list)
    for r in todo('remisiones?op_numero=not.is.null&select=remision,fecha,op_numero,item_cot,cantidad,cotizacion_id'):
        if r.get('op_numero') in nums:
            rems[r['op_numero']].append(r)
    precios = _cot_items([c for o in abiertas for c in cot_ids(o)])

    cartera = todo('cartera_facturas?select=numero,cliente_nombre,cotizacion_id,oc,oc_num,monto_antes_iva,fecha_facturacion')
    fac_por_cot = collections.defaultdict(list)
    fac_por_oc = collections.defaultdict(list)
    for f in cartera:
        for c in re.split(r'[,;/ ]+', str(f.get('cotizacion_id') or '')):
            if clave(c):
                fac_por_cot[clave(c)].append(f)
        for oc in (f.get('oc'), f.get('oc_num')):
            if len(clave(oc)) >= 4:
                fac_por_oc[clave(oc)].append(f)

    entregado, facturado, en_curso = [], [], []
    for o in abiertas:
        its = op_items.get(o['id'], [])
        valor_op = sum(num(x.get('v_total')) or num(x.get('cantidad')) * num(x.get('v_unit')) for x in its) \
            or num(o.get('valor_venta'))
        # Entregado según la OP (op_items). «despachado» sin cantidad = todo.
        ent_op = 0.0
        for x in its:
            q = num(x.get('cantidad'))
            d = q if (x.get('estado') == 'despachado' and not x.get('despachada')) else min(num(x.get('despachada')), q)
            ent_op += d * num(x.get('v_unit'))
        # Entregado según las remisiones, al precio de la cotización.
        rs = rems.get(o['numero'], [])
        ids = cot_ids(o)
        q_item = collections.defaultdict(float)
        for r in rs:
            cid = r.get('cotizacion_id') if r.get('cotizacion_id') in ids else (ids[0] if ids else None)
            if r.get('item_cot') is not None and str(r['item_cot']) in precios.get(cid, {}):
                q_item[(cid, str(r['item_cot']))] += num(r.get('cantidad'))
        ent_rem = 0.0
        for (cid, i), q in q_item.items():
            vu, qc = precios[cid][i]
            ent_rem += min(q, qc or q) * vu
        ent = min(max(ent_op, ent_rem), valor_op) if valor_op else max(ent_op, ent_rem)
        fechas = [fecha(r.get('fecha')) for r in rs if fecha(r.get('fecha'))]
        f0 = min(fechas) if fechas else None
        fila = {
            'op': o['numero'], 'cot': ' + '.join(ids) or '—', 'cliente': o.get('cliente') or '—',
            'oc': o.get('oc_cliente') or '—', 'estado': o.get('estado'),
            'rem': ', '.join(sorted(set(str(r['remision']) for r in rs))) or '—',
            'valor_op': valor_op, 'entregado': ent, 'parcial': ent < valor_op * 0.995,
            'dias': (hoy - f0).days if f0 else None}
        facs = list({f['numero']: f for c in ids for f in fac_por_cot.get(clave(c), [])}.values())
        if facs:
            fila['facturas'] = ', '.join(sorted(set(str(f['numero']) for f in facs)))
            fila['monto_fac'] = sum(num(f.get('monto_antes_iva')) for f in facs)
            facturado.append(fila)
        elif ent > 0:
            entregado.append(fila)
        else:
            en_curso.append(fila)

    # Adjudicadas sin OP, cruzadas con Cartera.
    con_op = set(c for o in ops for c in cot_ids(o))
    oc_mod = ocs_del_modulo()
    sin_op = []
    for c in todo('cotizaciones?estado=eq.Adjudicada&deleted=is.false'
                  '&select=id,cliente,fecha,subtotal,valor_adjudicado,adjudicada_at,adjudicacion_at&order=id'):
        if c['id'] in con_op:
            continue
        va = c.get('valor_adjudicado')
        valor = num(va) if va is not None else num(c.get('subtotal'))
        pistas = []
        for f in fac_por_cot.get(clave(c['id']), []):
            pistas.append('Fact. %s (misma cotización, %s)' % (f['numero'], money(f.get('monto_antes_iva'))))
        for oc in oc_mod.get(c['id'], []):
            for f in fac_por_oc.get(clave(oc), []):
                pistas.append('Fact. %s (misma O.C. %s, %s)' % (f['numero'], oc, money(f.get('monto_antes_iva'))))
        if not pistas and valor > 0:
            for f in cartera:
                m = num(f.get('monto_antes_iva'))
                if m and abs(m - valor) / valor < 0.02 and mismo_cliente(f.get('cliente_nombre'), c.get('cliente')):
                    pistas.append('Fact. %s (mismo cliente y monto, %s)' % (f['numero'], f.get('fecha_facturacion') or ''))
        f = fecha(c.get('adjudicada_at')) or fecha(c.get('adjudicacion_at')) or fecha(c.get('fecha'))
        sin_op.append({'cot': c['id'], 'cliente': c.get('cliente') or '—', 'valor': valor, 'estimado': va is None,
                       'dias': (hoy - f).days if f else None, 'pistas': list(dict.fromkeys(pistas))})
    return entregado, facturado, en_curso, sin_op


TD = 'style="padding:5px 8px;border-bottom:1px solid #e5e7eb;vertical-align:top"'
TDR = 'style="padding:5px 8px;border-bottom:1px solid #e5e7eb;text-align:right;white-space:nowrap;vertical-align:top"'


def _tabla(cab, filas):
    p = ['<table style="border-collapse:collapse;font-size:12.5px;width:100%"><tr style="background:#f3f4f6">']
    for c in cab:
        p.append('<th %s>%s</th>' % (TDR if c.startswith('>') else TD, c.lstrip('>')))
    p.append('</tr>')
    for f in filas:
        p.append('<tr>' + ''.join('<td %s>%s</td>' % (TDR if cab[i].startswith('>') else TD, v)
                                  for i, v in enumerate(f)) + '</tr>')
    p.append('</table>')
    return ''.join(p)


def _caja(titulo, valor, n, fuerte=False):
    borde = '2px solid #1a3a8f' if fuerte else '1px solid #e5e7eb'
    color = '#1a3a8f' if fuerte else '#1f2937'
    return ('<td style="padding:8px 14px;border:%s">'
            '<div style="font-size:11px;color:#6b7280">%s</div>'
            '<div style="font-size:18px;font-weight:800;color:%s">%s</div>'
            '<div style="font-size:11px;color:#6b7280">%s</div></td>' % (borde, titulo, color, money(valor), n))


def html(entregado, facturado, en_curso, sin_op):
    hoy = datetime.date.today().strftime('%d/%m/%Y')
    t_ent = sum(f['entregado'] for f in entregado)
    t_fac = sum(f['entregado'] for f in facturado)
    t_cur = sum(f['valor_op'] for f in en_curso) + sum(f['valor_op'] - f['entregado'] for f in entregado)
    sin_rastro = [f for f in sin_op if not f['pistas']]
    con_pista = [f for f in sin_op if f['pistas']]
    p = ['<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#1f2937">',
         '<h2 style="color:#1a3a8f;margin:0 0 2px">💼 Cartera por legalizar · %s</h2>' % hoy,
         '<p style="margin:0 0 12px;color:#4b5563">Por <b>orden de pedido</b>: lo que ya se le entregó al cliente '
         'y no se ha facturado, cruzado con Cartera. Valores antes de IVA.</p>',
         '<table style="border-collapse:collapse;margin-bottom:16px"><tr>',
         _caja('ENTREGADO SIN FACTURAR', t_ent, '%d OP' % len(entregado), True),
         _caja('Facturado, falta cerrar OP', t_fac, '%d OP' % len(facturado)),
         _caja('Pedido pendiente de entregar', t_cur, '%d OP' % (len(en_curso) + sum(1 for f in entregado if f['parcial']))),
         _caja('Adjudicado sin OP (por confirmar)', sum(f['valor'] for f in sin_rastro), '%d cotiz.' % len(sin_rastro)),
         '</tr></table>']

    if entregado:
        p.append('<h3 style="color:#b91c1c;margin:14px 0 2px">1 · Entregado sin facturar — %s</h3>' % money(t_ent))
        p.append('<p style="margin:0 0 6px;font-size:12px;color:#6b7280">Ya salió de bodega y no hay factura en Cartera: '
                 'es lo primero que hay que facturar. Ordenado por antigüedad.</p>')
        filas = []
        for f in sorted(entregado, key=lambda f: -(f['dias'] or 0)):
            valor = '<b>%s</b>' % money(f['entregado'])
            if f['parcial']:
                valor += '<br><span style="font-size:10.5px;color:#b45309">parcial de %s</span>' % money(f['valor_op'])
            filas.append(['<b>%s</b>' % f['op'], f['cot'], f['cliente'], f['oc'], f['rem'],
                          f['dias'] if f['dias'] is not None else '—', valor])
        p.append(_tabla(['OP', 'Cotización', 'Cliente', 'O.C. cliente', 'Remisiones', '>Días', '>Entregado'], filas))

    if facturado:
        p.append('<h3 style="color:#047857;margin:18px 0 2px">2 · Ya facturado en Cartera, falta cerrar la OP — %s</h3>' % money(t_fac))
        p.append('<p style="margin:0 0 6px;font-size:12px;color:#6b7280">La factura existe en Cartera con el mismo N° de '
                 'cotización. No es plata pendiente: hay que <b>cerrar la OP con ese N° de factura</b> para que salga de aquí.</p>')
        filas = []
        for f in sorted(facturado, key=lambda f: f['op']):
            dif = f['monto_fac'] - f['entregado']
            nota = ''
            if abs(dif) > max(1000, f['entregado'] * 0.01):
                nota = '<br><span style="font-size:10.5px;color:#b45309">factura %s %s que lo entregado</span>' % (
                    money(abs(dif)), 'más' if dif > 0 else 'menos')
            filas.append(['<b>%s</b>' % f['op'], f['cot'], f['cliente'], 'Fact. <b>%s</b>' % f['facturas'],
                          money(f['monto_fac']) + nota, money(f['entregado'])])
        p.append(_tabla(['OP', 'Cotización', 'Cliente', 'Cartera', '>Facturado', '>Entregado'], filas))

    if en_curso:
        p.append('<h3 style="color:#1a3a8f;margin:18px 0 2px">3 · OP en curso sin entregas — %s</h3>'
                 % money(sum(f['valor_op'] for f in en_curso)))
        p.append('<p style="margin:0 0 6px;font-size:12px;color:#6b7280">Pedido vivo, todavía no se ha despachado nada.</p>')
        filas = [['<b>%s</b>' % f['op'], f['cot'], f['cliente'], f['oc'], f['estado'].replace('_', ' '), money(f['valor_op'])]
                 for f in sorted(en_curso, key=lambda f: f['op'])]
        p.append(_tabla(['OP', 'Cotización', 'Cliente', 'O.C. cliente', 'Estado', '>Valor OP'], filas))

    if sin_op:
        p.append('<h3 style="color:#6b7280;margin:18px 0 2px">4 · Adjudicadas sin OP</h3>')
        p.append('<p style="margin:0 0 6px;font-size:12px;color:#6b7280">No tienen orden de pedido. Las que tienen una pista '
                 'en Cartera probablemente ya se facturaron y solo falta pasarlas a «Facturada». Las demás hay que '
                 'confirmarlas con el comercial: ¿sigue vivo el pedido?</p>')
        filas = []
        for f in sorted(sin_op, key=lambda f: (bool(f['pistas']), -(f['valor']))):
            valor = money(f['valor']) + ('<br><span style="font-size:10.5px;color:#b45309">valor cotizado</span>'
                                         if f['estimado'] else '')
            pista = '<br>'.join(f['pistas']) if f['pistas'] else '<span style="color:#b91c1c">Sin rastro en Cartera</span>'
            filas.append(['<b>%s</b>' % f['cot'], f['cliente'], f['dias'] if f['dias'] is not None else '—', valor, pista])
        p.append(_tabla(['Cotización', 'Cliente', '>Días', '>Valor', 'En Cartera'], filas))
        if con_pista:
            p.append('<p style="margin:6px 0 0;font-size:11.5px;color:#6b7280">%d con pista en Cartera por %s; '
                     '%d sin rastro por %s.</p>' % (len(con_pista), money(sum(f['valor'] for f in con_pista)),
                                                   len(sin_rastro), money(sum(f['valor'] for f in sin_rastro))))

    p.append('<p style="margin-top:14px;font-size:11.5px;color:#6b7280">Cómo sale de este informe: facturar y registrar '
             'la factura en Cartera con el N° de cotización, y cerrar la OP con ese N° de factura. '
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
    entregado, facturado, en_curso, sin_op = datos()
    total = sum(f['entregado'] for f in entregado)
    print('Entregado sin facturar: %d OP · %s | facturado sin cerrar OP: %d | en curso: %d | sin OP: %d'
          % (len(entregado), money(total), len(facturado), len(en_curso), len(sin_op)))
    if not (entregado or facturado or en_curso or sin_op):
        print('Nada por legalizar: no se manda correo.')
        return
    asunto = '💼 Cartera por legalizar: %s entregado sin facturar en %d OP' % (money(total), len(entregado))
    cuerpo = html(entregado, facturado, en_curso, sin_op)
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
