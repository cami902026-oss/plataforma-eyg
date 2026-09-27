"""
VIGILANTE DE LA PLATAFORMA — E&G Energy Group
==============================================
Revisa la base todos los días a las 7 a. m. y le escribe a Andrea SOLO si
encuentra algo que hay que corregir. Si no hay nada, no manda nada.

POR QUÉ EXISTE (27-sep-2026)
Cada regla sale de un error real que se descubrió tarde, cuando ya dolía:
  · OP-2026-0112: la tubería de 4" tenía el código de producto "60", que no
    existe → al despachar no descontó nada del inventario.
  · CC-OP-2026-0112: el plan de compras desapareció de la base sin que nadie
    lo borrara a propósito.
  · OP-2026-0107 / LM2145: 31 ítems adjudicados con precio de venta $0 → el
    centro de costos mostraba $3 M vendidos contra $50 M de costo.
  · OP-2026-0054: el valor de la OP decía $298.000 y sus ítems sumaban $6,1 M.
  · OP-2026-0112: 60 m despachados y la línea quedó «pendiente».

CÓMO NO SE VUELVE PAISAJE
  · Sin hallazgos → no hay correo.
  · Lo que ya se avisó no se repite a diario: se recuerda en una sola línea
    («siguen abiertos N») y el detalle solo va para lo NUEVO. El estado vive en
    el caché de GitHub Actions (vigilante_estado.json), no en el repositorio.
  · No usa Claude: son reglas fijas contra la base. Cuesta $0.

Solo LEE la base (key publicable). No corrige nada: dice qué, dónde y qué hacer.

Prueba local sin enviar:
    TEST_OUT=1 PYTHONIOENCODING=utf-8 python scripts/vigilante.py
"""

import os
import json
import hashlib
import datetime
import urllib.request
import urllib.parse

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_KEY', '') or 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM'
PARA = ['andrea.bernal@eygenergygroup.com']
REMITENTE = os.environ.get('REMITENTE', 'info@eygenergygroup.com')
ESTADO = os.environ.get('VIGILANTE_ESTADO', 'vigilante_estado.json')
PLATAFORMA = 'https://cami902026-oss.github.io/plataforma-eyg/Index.html'
OP_CERRADAS = ('anulada', 'cerrada')


def sb(path):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY, 'User-Agent': 'energy-vigilante'}
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path, headers=h)
    with urllib.request.urlopen(r, timeout=60) as resp:
        t = resp.read().decode()
        return json.loads(t) if t.strip() else []


def todo(path):
    """PostgREST corta en 1.000 filas: se pagina siempre."""
    out, off = [], 0
    sep = '&' if '?' in path else '?'
    while True:
        d = sb(path + sep + 'limit=1000&offset=' + str(off)) or []
        out.extend(d)
        if len(d) < 1000 or off > 100000:
            return out
        off += 1000


def money(n):
    try:
        return '$' + format(int(round(float(n or 0))), ',d').replace(',', '.')
    except Exception:
        return '$0'


def corto(t, n=70):
    t = ' '.join(str(t or '').split())
    return t if len(t) <= n else t[:n - 1] + '…'


# ─── Reglas ──────────────────────────────────────────────────────────────────
# Cada hallazgo: (clave estable, título de la regla, texto, qué hacer)

def revisar():
    ops = todo('ops?select=id,numero,cotizacion_id,cliente,estado,valor_venta,created_at&order=id')
    vivas = {o['id']: o for o in ops if (o.get('estado') or '') not in OP_CERRADAS}
    items = todo('op_items?select=id,op_id,item,descripcion,cantidad,v_unit,v_total,origen,'
                 'producto_codigo,estado,despachada,costo_unit&order=id')
    prods = {p['codigo'] for p in todo('productos?select=codigo')}
    plan = todo('plan_compras?select=cc,op_numero,costo_unit,proveedor,descripcion,cantidad')
    eventos = todo('op_eventos?select=op_id,evento,detalle,at&evento=in.(plan_generado,plan_eliminado,plan_restaurado,plan_completado)&order=id')

    H = []
    por_op = {}
    for it in items:
        por_op.setdefault(it['op_id'], []).append(it)

    for oid, o in vivas.items():
        num, its = o['numero'], por_op.get(oid, [])
        cli = corto(o.get('cliente'), 30)

        # 1) Código de producto que no existe (el caso del "60")
        for it in its:
            c = str(it.get('producto_codigo') or '').strip()
            if it.get('origen') == 'BODEGA' and c and not c.startswith('OP-') and c not in prods:
                H.append(('codigo|%s|%s' % (it['id'], c), 'Código de producto que no existe',
                          '%s (%s) · ítem %s «%s»: código «%s» no está en el inventario'
                          % (num, cli, it.get('item'), corto(it.get('descripcion'), 50), c),
                          'Al despachar NO va a descontar. Poner el código correcto en la OP.'))

        # 2) Sale de bodega pero sin producto
        for it in its:
            if it.get('origen') == 'BODEGA' and not str(it.get('producto_codigo') or '').strip() \
                    and it.get('estado') not in ('despachado',):
                H.append(('sincodigo|%s' % it['id'], 'Sale de bodega sin producto amarrado',
                          '%s (%s) · ítem %s «%s»' % (num, cli, it.get('item'), corto(it.get('descripcion'), 50)),
                          'Decir de qué código sale antes de despachar.'))

        # 3) Ítems con precio de venta $0 (lo adjudicado sin precio)
        ceros = [it for it in its if not float(it.get('v_unit') or 0) and float(it.get('cantidad') or 0) > 0]
        if ceros:
            H.append(('venta0|%s|%d' % (num, len(ceros)), 'Ítems vendidos en $0',
                      '%s (%s, %s): %d de %d ítems sin precio de venta'
                      % (num, cli, o.get('cotizacion_id') or '—', len(ceros), len(its)),
                      'El centro de costos y el margen salen falsos. Revisar la cotización / O.C. del cliente.'))

        # 4) El valor de la OP no cuadra con sus ítems
        suma = sum(float(it.get('v_total') or 0) for it in its)
        vv = float(o.get('valor_venta') or 0)
        if its and suma > 0 and abs(vv - suma) > max(1000, suma * 0.01):
            H.append(('valor|%s|%d|%d' % (num, round(vv), round(suma)), 'Valor de la OP no cuadra',
                      '%s (%s): la OP dice %s y sus ítems suman %s' % (num, cli, money(vv), money(suma)),
                      'Corregir el valor de la OP (no se recalcula solo).'))

        # 5) Despacho a medias: entregado completo pero sigue «pendiente»/«comprado».
        #    Una línea por OP (renglón por renglón era ruido).
        abiertos = [it for it in its
                    if float(it.get('cantidad') or 0) > 0
                    and float(it.get('despachada') or 0) >= float(it.get('cantidad') or 0)
                    and it.get('estado') != 'despachado']
        if abiertos:
            H.append(('desp|%s|%d' % (num, len(abiertos)), 'Ítems entregados completos pero sin cerrar',
                      '%s (%s, estado de la OP: %s): %d ítem(s) ya salieron completos y siguen «%s»'
                      % (num, cli, o.get('estado'), len(abiertos),
                         '/'.join(sorted({str(it.get('estado')) for it in abiertos}))),
                      'Marcarlos despachados: con eso la OP puede cerrar.'))
        elif o.get('estado') == 'despachada' and any(it.get('estado') != 'despachado' for it in its):
            n = sum(1 for it in its if it.get('estado') != 'despachado')
            H.append(('opdesp|%s|%d' % (num, n), 'OP «despachada» con ítems pendientes',
                      '%s (%s): %d ítem(s) no están despachados' % (num, cli, n),
                      'Ligar la remisión que falta o corregir el estado.'))

        # 6) Costo mayor que la venta (el patrón de LM1945-1)
        for it in its:
            v, c = float(it.get('v_unit') or 0), float(it.get('costo_unit') or 0)
            if v > 0 and c > v * 1.02:
                H.append(('costo|%s|%d|%d' % (it['id'], round(c), round(v)), 'Costo mayor que la venta',
                          '%s (%s) · ítem %s «%s»: costo %s vs venta %s'
                          % (num, cli, it.get('item'), corto(it.get('descripcion'), 40), money(c), money(v)),
                          'Revisar el costo del proveedor o el precio (pierde plata).'))

    # 7) Planes de compra que desaparecieron (el caso CC-OP-2026-0112)
    ccs = {r.get('cc') for r in plan}
    ult = {}
    for e in eventos:
        cc = str(e.get('detalle') or '').split(' ')[0]
        if cc.startswith('CC-'):
            ult[cc] = e
    ops_por_id = {o['id']: o for o in ops}
    for cc, e in ult.items():
        o = ops_por_id.get(e.get('op_id')) or {}
        if e.get('evento') != 'plan_eliminado' and cc not in ccs and (o.get('estado') or '') not in OP_CERRADAS:
            H.append(('plan|%s' % cc, 'Plan de compras desaparecido',
                      '%s (%s): se generó el %s y hoy no tiene ninguna línea en la base'
                      % (cc, corto(o.get('cliente'), 30), str(e.get('at') or '')[:10]),
                      'Restaurarlo desde el respaldo diario (backups/plan_compras.json).'))

    # 8) Plan de OP viva con costos de relleno o sin proveedor
    for r in plan:
        num = r.get('op_numero')
        o = next((x for x in vivas.values() if x['numero'] == num), None) if num else None
        if not o:
            continue
        c = r.get('costo_unit')
        prov = str(r.get('proveedor') or '').strip().upper()
        if (c is not None and 0 < float(c) < 200) or (not prov) or prov == 'POR DEFINIR':
            H.append(('plancosto|%s|%s|%s' % (r.get('cc'), corto(r.get('descripcion'), 40), c), 'Plan con costo o proveedor sin definir',
                      '%s · «%s» x%s: costo %s, proveedor %s'
                      % (r.get('cc'), corto(r.get('descripcion'), 45), r.get('cantidad'),
                         money(c) if c is not None else 'vacío', prov or 'vacío'),
                      'El margen del centro de costos no es real hasta corregirlo.'))
    return H


# ─── Memoria de lo ya avisado ────────────────────────────────────────────────
def cargar_estado():
    try:
        with open(ESTADO, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def guardar_estado(st):
    with open(ESTADO, 'w', encoding='utf-8') as f:
        json.dump(st, f, ensure_ascii=False)


def clave(h):
    return hashlib.sha1(h[0].encode()).hexdigest()[:16]


def armar_html(nuevos, siguen, total):
    hoy = datetime.date.today().strftime('%d/%m/%Y')
    grupos = {}
    for h in nuevos:
        grupos.setdefault(h[1], []).append(h)
    partes = ['<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1f2937">',
              '<h2 style="color:#1a3a8f;margin:0 0 4px">🔎 Vigilante de la plataforma · %s</h2>' % hoy,
              '<p style="margin:0 0 14px;color:#4b5563">%d cosa(s) nueva(s) para corregir%s.</p>'
              % (len(nuevos), (' · siguen abiertas %d de días anteriores' % siguen) if siguen else '')]
    for titulo, hs in grupos.items():
        partes.append('<h3 style="color:#b45309;margin:16px 0 6px">%s (%d)</h3>' % (titulo, len(hs)))
        partes.append('<p style="margin:0 0 6px;color:#6b7280;font-size:12px">Qué hacer: %s</p>' % hs[0][3])
        partes.append('<ul style="margin:0 0 6px 18px;padding:0">')
        for h in hs[:25]:
            partes.append('<li style="margin:2px 0">%s</li>' % h[2])
        if len(hs) > 25:
            partes.append('<li>… y %d más</li>' % (len(hs) - 25))
        partes.append('</ul>')
    partes.append('<p style="margin-top:18px;font-size:12px;color:#6b7280">Si no hay nada que corregir, '
                  'este correo no llega. <a href="%s">Abrir la plataforma</a></p></div>' % PLATAFORMA)
    return ''.join(partes)


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


def enviar(asunto, html):
    msg = {'subject': asunto, 'body': {'contentType': 'HTML', 'content': html},
           'toRecipients': [{'emailAddress': {'address': a}} for a in PARA]}
    payload = json.dumps({'message': msg, 'saveToSentItems': True}).encode()
    r = urllib.request.Request(
        'https://graph.microsoft.com/v1.0/users/' + REMITENTE + '/sendMail',
        data=payload, method='POST',
        headers={'Authorization': 'Bearer ' + graph_token(), 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r) as resp:
        return resp.status


def main():
    hallazgos = revisar()
    st = cargar_estado()
    hoy = datetime.date.today().isoformat()
    actuales = {clave(h): h for h in hallazgos}
    nuevos = [h for k, h in actuales.items() if k not in st]
    siguen = len([k for k in actuales if k in st])
    # Se olvida lo que ya se corrigió; lo nuevo queda anotado con la fecha del aviso.
    st = {k: st.get(k, hoy) for k in actuales}
    print('Hallazgos: %d · nuevos: %d · siguen abiertos: %d' % (len(hallazgos), len(nuevos), siguen))
    for h in nuevos:
        print(' +', h[1], '|', h[2])
    if not nuevos:
        print('Nada nuevo: no se manda correo.')
        guardar_estado(st)
        return
    asunto = '🔎 Plataforma: %d cosa(s) nueva(s) para corregir' % len(nuevos)
    html = armar_html(nuevos, siguen, len(hallazgos))
    if os.environ.get('TEST_OUT'):
        with open('vigilante_prueba.html', 'w', encoding='utf-8') as f:
            f.write(html)
        print('TEST_OUT: correo NO enviado; vista previa en vigilante_prueba.html')
        return
    enviar(asunto, html)
    guardar_estado(st)      # solo se da por avisado si el correo salió
    print('Correo enviado a', ', '.join(PARA))


if __name__ == '__main__':
    main()
