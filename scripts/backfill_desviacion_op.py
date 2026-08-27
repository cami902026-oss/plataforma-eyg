# -*- coding: utf-8 -*-
"""
Rellena la DESVIACIÓN TÉCNICA que se perdió en el camino cotización → OP → plan.

  python backfill_desviacion_op.py            → simulación (no escribe)
  python backfill_desviacion_op.py --write    → aplica

Por qué: hasta la v223 `op_items` no tenía dónde guardarla y `opGenerarPlan`
escribía `nota: null` en las líneas de compra. Resultado: 14 de 23 líneas de los
planes nacidos de una OP salieron sin la desviación, y esa es la nota que la O.C.
imprime en "Referencia, Especificaciones o datos de ingeniería" (E6-FC-01).

Qué hace, en dos pasos:
  1. op_items.desviacion_tecnica  <- cotizacion_items.desviacion_tecnica
     Cruce EXACTO por `cotizacion_item_uid`. Si el ítem no lo tiene, cruza por
     descripción normalizada dentro de la MISMA cotización. Nada de adivinar.
  2. plan_compras.nota            <- la desviación, anteponiéndola a lo que haya.

Regla de oro: NUNCA se pisa una nota escrita por una persona.
  · nota vacía             -> queda la desviación
  · nota "Sale de bodega…" -> queda "<desviación> · Sale de bodega…"
  · cualquier otro texto   -> NO SE TOCA, solo se reporta

Requisito: correr antes scripts/migracion_desviacion_op_2026-08-26.sql
Escribe por el energy-proxy (la RLS de Supabase está cerrada para la key pública).
"""
import json, io, os, sys, urllib.request

RAIZ   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG    = json.load(io.open(os.path.join(RAIZ, 'data', 'config.json'), encoding='utf-8'))
PROXY  = CFG.get('ghProxyUrl') or CFG.get('proxyUrl')
SECRET = 'eyg_prx_c70613f89a19d97c73d8800612029462'
SB_V   = 129
SB_URL = 'https://juprjevxkcitqpsnemto.supabase.co/rest/v1'
SB_KEY = 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM'
WRITE  = '--write' in sys.argv
sys.stdout.reconfigure(encoding='utf-8')


def leer(path):
    req = urllib.request.Request(SB_URL + '/' + path,
                                 headers={'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode('utf-8'))


def escribir(path, method, body):
    payload = json.dumps({'secret': SECRET, 'sb': {'path': path, 'method': method,
                          'body': json.dumps(body), 'prefer': 'return=minimal', 'v': SB_V}})
    req = urllib.request.Request(PROXY, data=payload.encode('utf-8'),
                                 headers={'Content-Type': 'text/plain;charset=utf-8',
                                          'User-Agent': 'EYG-script/1.0'})   # no-navegador a propósito
    with urllib.request.urlopen(req, timeout=45) as r:
        j = json.loads(r.read().decode('utf-8'))
    return j.get('status'), (j.get('body') or '')


norm = lambda s: ' '.join(str(s or '').upper().split())

# ── Datos ────────────────────────────────────────────────────────────────
ops   = {o['id']: o for o in leer('ops?select=id,numero,cotizacion_id,cliente,deleted&limit=1000')
         if not o.get('deleted')}
items = leer('op_items?select=id,op_id,item,descripcion,cotizacion_item_uid,desviacion_tecnica&limit=3000')
items = [i for i in items if i['op_id'] in ops]

cots = sorted({ops[i['op_id']].get('cotizacion_id') for i in items if ops[i['op_id']].get('cotizacion_id')})
ci = []
for j in range(0, len(cots), 40):                      # de a tandas: el filtro in.() tiene límite
    lote = ','.join('"%s"' % c for c in cots[j:j + 40])
    ci += leer('cotizacion_items?select=cotizacion_id,uid,item,descripcion,desviacion_tecnica'
               '&cotizacion_id=in.(%s)&limit=3000' % lote)

por_uid, por_desc = {}, {}
for c in ci:
    dv = (c.get('desviacion_tecnica') or '').strip()
    if not dv:
        continue
    if c.get('uid'):
        por_uid[c['uid']] = dv
    por_desc.setdefault((c['cotizacion_id'], norm(c['descripcion'])), dv)

# ── Paso 1: op_items ─────────────────────────────────────────────────────
print('=' * 74)
print('PASO 1 — desviación técnica de la cotización -> op_items')
print('=' * 74)
p1 = 0
for it in items:
    if (it.get('desviacion_tecnica') or '').strip():
        continue                                        # ya la tiene: no se pisa
    cot = ops[it['op_id']].get('cotizacion_id')
    dv = por_uid.get(it.get('cotizacion_item_uid')) or por_desc.get((cot, norm(it['descripcion'])))
    if not dv:
        continue
    p1 += 1
    print('  %-14s it %-3s %-34s -> %s' % (ops[it['op_id']]['numero'], it['item'],
                                           str(it['descripcion'])[:34], dv[:44]))
    it['desviacion_tecnica'] = dv                       # para el paso 2
    if WRITE:
        st, tx = escribir('op_items?id=eq.%d' % it['id'], 'PATCH', {'desviacion_tecnica': dv})
        if st not in (200, 204):
            print('     [!] %s %s' % (st, tx[:140]))
print('  -> %d linea(s) de OP' % p1)

# ── Paso 2: plan_compras ─────────────────────────────────────────────────
print()
print('=' * 74)
print('PASO 2 — desviación técnica -> nota del Plan de Compras (la que imprime la O.C.)')
print('=' * 74)
planes = leer('plan_compras?select=id,cc,cotizacion,item,descripcion,nota&cc=like.CC-OP*&limit=1000')
desv_op = {}
for it in items:
    dv = (it.get('desviacion_tecnica') or '').strip()
    if dv:
        desv_op[('CC-' + ops[it['op_id']]['numero'], norm(it['descripcion']))] = dv

p2 = respetadas = 0
for r in planes:
    dv = desv_op.get((r['cc'], norm(r['descripcion'])))
    if not dv:
        continue
    nota = (r.get('nota') or '').strip()
    if dv in nota:
        continue                                        # ya está
    if nota and not nota.startswith('Sale de bodega'):
        respetadas += 1
        print('  [=] %-17s it %-3s nota escrita a mano, NO se toca: %r' % (r['cc'], r['item'], nota[:40]))
        continue
    nueva = ' · '.join(x for x in (dv, nota) if x)
    p2 += 1
    print('  %-17s it %-3s %-30s -> %s' % (r['cc'], r['item'], str(r['descripcion'])[:30], nueva[:44]))
    if WRITE:
        st, tx = escribir('plan_compras?id=eq.%d' % r['id'], 'PATCH', {'nota': nueva})
        if st not in (200, 204):
            print('     [!] %s %s' % (st, tx[:140]))
print('  -> %d linea(s) de plan · %d respetada(s) por tener nota humana' % (p2, respetadas))

print()
print('MODO ESCRITURA' if WRITE else 'SIMULACION — nada se escribio. Corre con --write para aplicar.')
