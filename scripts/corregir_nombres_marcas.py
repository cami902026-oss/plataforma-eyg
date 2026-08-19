# -*- coding: utf-8 -*-
"""
Corrige SOLO errores de escritura en `productos`: descripción y marca.
NO toca stock, costo, precio, código, ubicación, familia ni kardex.

  python corregir_nombres_marcas.py            → simulación (no escribe)
  python corregir_nombres_marcas.py --write    → aplica

Respaldo previo: backups/productos_nombres_ANTES_<fecha>.json
Escribe por el energy-proxy (la RLS de Supabase está cerrada para la key pública).
"""
import json, io, os, re, sys, time, urllib.request

RAIZ    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG     = json.load(io.open(os.path.join(RAIZ, 'data', 'config.json'), encoding='utf-8'))
PROXY   = CFG.get('ghProxyUrl') or CFG.get('proxyUrl')
SECRET  = 'eyg_prx_c70613f89a19d97c73d8800612029462'
SB_V    = 129
SB_URL  = 'https://juprjevxkcitqpsnemto.supabase.co/rest/v1'
SB_KEY  = 'sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM'
WRITE   = '--write' in sys.argv
sys.stdout.reconfigure(encoding='utf-8')

def leer(path):
    req = urllib.request.Request(SB_URL + '/' + path, headers={'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode('utf-8'))

def escribir(path, method, body):
    """PATCH vía energy-proxy. Devuelve (status, texto)."""
    payload = json.dumps({'secret': SECRET, 'sb': {'path': path, 'method': method,
                                                   'body': json.dumps(body), 'prefer': 'return=minimal', 'v': SB_V}})
    req = urllib.request.Request(PROXY, data=payload.encode('utf-8'),
                                 headers={'Content-Type': 'text/plain;charset=utf-8',
                                          'User-Agent': 'EYG-script/1.0'})   # no-navegador a propósito
    with urllib.request.urlopen(req, timeout=45) as r:
        j = json.loads(r.read().decode('utf-8'))
    return j.get('status'), (j.get('body') or '')

# ── SOLO errores de escritura (no reglas de formato) ─────────────────────
MARCA_FIX = {'': 'SIN MARCA', '-': 'SIN MARCA', '0': 'SIN MARCA', 'N.A.': 'SIN MARCA', 'NA': 'SIN MARCA',
             'TORNILOS Y PARTES': 'TORNILLOS Y PARTES', 'SWAGELOCK': 'SWAGELOK'}
TIPO_FIX  = {'ESPIROTALICOS': 'ESPIROMETALICO', 'ESPIROTALICO': 'ESPIROMETALICO',
             'COENCTOR': 'CONECTOR', 'CONDULTEA': 'CONDULETA', 'FLANCHECON': 'FLANCHE CON'}

def arregla_desc(d):
    d0 = d or ''
    d1 = re.sub(r'\s+', ' ', d0.strip())
    p = d1.split(' ')
    if p and p[0].upper() in TIPO_FIX:
        p[0] = TIPO_FIX[p[0].upper()]; d1 = ' '.join(p)
    d1 = re.sub(r'\((\s*[\d,\.]+\s*CM)\(', r'(\1)', d1)     # (10 CM( → (10 CM)
    return d1

def arregla_marca(m):
    m0 = (m or '').strip()
    return MARCA_FIX.get(m0.upper(), m0)

prods = leer('productos?select=codigo,descripcion,marca&limit=5000')
cambios = []
for p in prods:
    d0, m0 = p.get('descripcion') or '', p.get('marca') or ''
    d1, m1 = arregla_desc(d0), arregla_marca(m0)
    body = {}
    if d1 != d0: body['descripcion'] = d1
    if m1 != m0: body['marca'] = m1
    if body: cambios.append((p['codigo'], d0, d1, m0, m1, body))

print('%s — %d producto(s) a corregir de %d\n' % ('APLICANDO' if WRITE else 'SIMULACIÓN (no escribe)', len(cambios), len(prods)))
for cod, d0, d1, m0, m1, body in cambios:
    det = []
    if 'descripcion' in body: det.append('desc: "%s"  →  "%s"' % (d0, d1))
    if 'marca' in body:       det.append('marca: "%s"  →  "%s"' % (m0 or '(vacío)', m1))
    print('  %-12s %s' % (cod, '   |   '.join(det)))

if not WRITE:
    print('\nNada escrito. Para aplicar:  python corregir_nombres_marcas.py --write')
    sys.exit(0)

print('\n--- escribiendo ---')
ok = err = 0
for cod, d0, d1, m0, m1, body in cambios:
    try:
        st, txt = escribir('productos?codigo=eq.' + urllib.parse.quote(str(cod)), 'PATCH', body)
        if st and 200 <= st < 300:
            ok += 1; print('  ✓ %-12s %s' % (cod, ','.join(body.keys())))
        else:
            err += 1; print('  ✗ %-12s status=%s %s' % (cod, st, txt[:120]))
    except Exception as e:
        err += 1; print('  ✗ %-12s %s' % (cod, e))
    time.sleep(0.25)
print('\nListos: %d   ·   con error: %d' % (ok, err))
