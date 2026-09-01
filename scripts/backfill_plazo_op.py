# -*- coding: utf-8 -*-
"""
BACKFILL de `ops.fecha_comprometida` — E&G Energy Group (v239)
==============================================================
Las OP creadas antes de la v239 no tienen fecha comprometida, así que caen al
plazo corto por etapa. Las que llevan material de FABRICACIÓN se pintarían
atrasadas sin serlo. Esto se la calcula, una sola vez, con la misma regla que
usa la plataforma al crear una OP nueva:

    fecha comprometida = fecha de creación de la OP + el tiempo de entrega
                         de su línea MÁS LENTA

El tiempo sale de `cotizacion_items.tiempo_entrega`, amarrado por
`op_items.cotizacion_item_uid`. El intérprete es el mismo `_opDiasEntrega` del
Index.html, portado tal cual (probado contra los 25 valores distintos que hay).

NO PISA NADA: solo escribe donde `fecha_comprometida` está vacía, y solo en OP
que no estén cerradas ni anuladas — ponerle plazo a algo que ya terminó no
cambia nada y ensucia el dato.

Requiere que la migración `migracion_plazo_op_2026-09-01.sql` esté corrida.

    python scripts/backfill_plazo_op.py            # dry run
    python scripts/backfill_plazo_op.py --write    # aplica
"""

import os
import re
import sys
import json
import math
import datetime
import urllib.request

SB_URL = os.environ.get('SUPABASE_URL', 'https://juprjevxkcitqpsnemto.supabase.co').rstrip('/')
SB_READ = os.environ.get('SUPABASE_KEY', 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM')
PROXY = os.environ.get('ENERGY_PROXY', '')
PROXY_SECRET = os.environ.get('ENERGY_PROXY_SECRET', '')
SB_WRITE_V = 129

ARCHIVADAS = ('cerrada', 'anulada')
WRITE = '--write' in sys.argv


def sbr(path):
    r = urllib.request.Request(SB_URL + '/rest/v1/' + path,
                               headers={'apikey': SB_READ, 'Authorization': 'Bearer ' + SB_READ,
                                        'User-Agent': 'energy-backfill-plazo'})
    with urllib.request.urlopen(r, timeout=60) as resp:
        t = resp.read().decode()
    return json.loads(t) if t.strip() else []


def sbw(path, method, body):
    """Escribe por el energy-proxy: la key del HTML es de solo lectura (RLS)."""
    if not PROXY or not PROXY_SECRET:
        raise SystemExit('Falta ENERGY_PROXY y/o ENERGY_PROXY_SECRET en el entorno.')
    payload = {'secret': PROXY_SECRET,
               'sb': {'path': path, 'method': method, 'body': json.dumps(body),
                      'prefer': 'return=minimal', 'v': SB_WRITE_V}}
    r = urllib.request.Request(PROXY, method='POST',
                               headers={'Content-Type': 'text/plain;charset=utf-8'},
                               data=json.dumps(payload).encode())
    with urllib.request.urlopen(r, timeout=60) as resp:
        return json.loads(resp.read().decode())


def dias_entrega(txt):
    """Puerto exacto de _opDiasEntrega (Index.html). Devuelve días CALENDARIO."""
    t = (txt or '').upper()
    for a, b in (('ÁÀÄÂ', 'A'), ('ÉÈËÊ', 'E'), ('ÍÌÏÎ', 'I'), ('ÓÒÖÔ', 'O'), ('ÚÙÜÛ', 'U')):
        for ch in a:
            t = t.replace(ch, b)
    t = t.strip()
    if not t:
        return None
    if re.search(r'INMEDIAT|ENTREGA YA|STOCK|DISPONIBLE', t):
        return 0
    nums = [float(n.replace(',', '.')) for n in re.findall(r'\d+(?:[.,]\d+)?', t)]
    if not nums:
        return None                      # "A CONVENIR" y parecidos
    n = max(nums)                        # el tope del rango: "1-2 DIAS" son 2
    if 'SEMANA' in t:
        d = n * 7
    elif re.search(r'\bMES(ES)?\b', t):
        d = n * 30
    else:
        d = n
    if 'HABIL' in t and not re.search(r'SEMANA|MES', t):
        d = math.ceil(d * 7 / 5)         # 10 hábiles son 2 semanas
    return int(round(d))


def main():
    try:
        vivas = sbr('ops?deleted=is.false&select=id,numero,cliente,estado,created_at,'
                    'fecha_comprometida&order=id')
    except urllib.error.HTTPError as e:
        if e.code == 400:
            raise SystemExit(
                'La tabla `ops` todavia no tiene la columna `fecha_comprometida`.\n'
                'Corre primero scripts/migracion_plazo_op_2026-09-01.sql en el SQL\n'
                'editor de Supabase y vuelve a intentar.')
        raise
    ops = [o for o in vivas
           if o['estado'] not in ARCHIVADAS and not o.get('fecha_comprometida')]
    if not ops:
        print('Nada que hacer: todas las OP vivas ya tienen fecha comprometida.')
        return

    items = sbr('op_items?op_id=in.(%s)&select=op_id,cotizacion_item_uid,descripcion&limit=5000'
                % ','.join(str(o['id']) for o in ops))
    uids = sorted({i['cotizacion_item_uid'] for i in items if i.get('cotizacion_item_uid')})
    te = {}
    for k in range(0, len(uids), 60):
        lote = ','.join('"%s"' % u for u in uids[k:k + 60])
        for r in sbr('cotizacion_items?uid=in.(%s)&select=uid,tiempo_entrega' % lote):
            te[r['uid']] = r.get('tiempo_entrega')

    print('%-14s %-14s %-20s %-9s %-11s %s' %
          ('OP', 'cliente', 'estado', 'plazo', 'fecha', 'la manda'))
    tocadas = 0
    for o in ops:
        mios = [i for i in items if i['op_id'] == o['id']]
        peor, quien = None, ''
        for i in mios:
            d = dias_entrega(te.get(i.get('cotizacion_item_uid')))
            if d is not None and (peor is None or d > peor):
                peor, quien = d, (i.get('descripcion') or '')[:34]
        if peor is None:
            print('%-14s %-14s %-20s %-9s %-11s %s' %
                  (o['numero'], (o['cliente'] or '')[:14], o['estado'], '—', '—',
                   'sin tiempo de entrega, se deja sin fecha'))
            continue
        f = (datetime.datetime.fromisoformat(o['created_at'].replace('Z', '+00:00')).date()
             + datetime.timedelta(days=peor))
        print('%-14s %-14s %-20s %-9s %-11s %s' %
              (o['numero'], (o['cliente'] or '')[:14], o['estado'],
               '%d d' % peor, f.isoformat(), quien))
        if WRITE:
            r = sbw('ops?id=eq.%d' % o['id'], 'PATCH', {'fecha_comprometida': f.isoformat()})
            if not (200 <= (r.get('status') or 0) < 300):
                print('    ERROR: %s %s' % (r.get('status'), r.get('body')))
                continue
        tocadas += 1

    print()
    print(('APLICADO a %d OP.' if WRITE else 'DRY-RUN: se tocarían %d OP. '
           'Corre con --write para aplicar.') % tocadas)


if __name__ == '__main__':
    main()
