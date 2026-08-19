# -*- coding: utf-8 -*-
"""
Archiva (activo=false) los códigos REPETIDOS que ya están en stock 0.
NO borra: el kardex queda intacto y el producto se puede revivir.
NO toca stock, costo, precio, ubicación ni ningún otro dato.

  python apagar_duplicados_en_cero.py            → simulación
  python apagar_duplicados_en_cero.py --write    → aplica

CRITERIO (conservador, decidido con el usuario el 19-ago-2026):
  · Solo se apaga un código en 0 si en su grupo hay otro código con stock
    que es EL MISMO producto, DE LA MISMA MARCA y DEL MISMO PROVEEDOR.
  · Si el código en 0 es de otra marca u otro proveedor/origen (CODIFER /
    GRANADA / IMP), NO se apaga: no sobra, está AGOTADO, y esa separación
    por origen se mantiene a propósito para no mezclar costos.
  · Si todo el grupo está en 0, no se toca (no hay a dónde consolidar).
"""
import json, io, os, re, sys, time, collections, urllib.request, urllib.parse

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
    req = urllib.request.Request(SB_URL + '/' + path, headers={'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode('utf-8'))

def escribir(path, method, body):
    payload = json.dumps({'secret': SECRET, 'sb': {'path': path, 'method': method,
                                                   'body': json.dumps(body), 'prefer': 'return=minimal', 'v': SB_V}})
    req = urllib.request.Request(PROXY, data=payload.encode('utf-8'),
                                 headers={'Content-Type': 'text/plain;charset=utf-8', 'User-Agent': 'EYG-script/1.0'})
    with urllib.request.urlopen(req, timeout=45) as r:
        j = json.loads(r.read().decode('utf-8'))
    return j.get('status'), (j.get('body') or '')

# ── misma normalización del diagnóstico ──────────────────────────────────
PAL = {'BRIDA': 'FLANCHE', 'BRIDAS': 'FLANCHE', 'ELBOW': 'CODO', 'BLIND': 'CIEGO', 'WN': 'CUELLO', 'NIPPLE': 'NIPLE'}
def clave(d):
    k = re.sub(r'\s+', ' ', (d or '').upper().strip())
    k = ' '.join(PAL.get(w, w) for w in k.split(' '))
    k = re.sub(r'\bA\s*/\s*C\b', 'AC', k)
    k = re.sub(r'\bSCH\s*[- ]?\s*(\d{2,3})\b', r'S\1', k)
    k = re.sub(r'\b(\d)\.(\d{3})\b', r'\1\2', k)
    k = re.sub(r'\b(\d{3,4})\s*(?:#|LBS|LIBRAS|L)\b', r'\1L', k)
    k = re.sub(r'[\*"\'’]', ' ', k)
    k = re.sub(r'\bP\s*/\s*S\b', 'PS', k)
    k = re.sub(r'[.,]', ' ', k)
    return re.sub(r'\s+', ' ', k).strip()

def norm(x): return re.sub(r'\s+', ' ', (x or '').strip().upper())
def st(p):   return float(p.get('stock_actual') or 0)

prods = leer('productos?select=id,codigo,descripcion,marca,proveedor,ubicacion,stock_actual,familia_id,activo&activo=eq.true&limit=5000')
grupos = collections.defaultdict(list)
for p in prods:
    grupos[(p.get('familia_id'), clave(p.get('descripcion')))].append(p)
grupos = {k: v for k, v in grupos.items() if len(v) > 1}

apagar, agotados, todos_cero = [], [], []
for k, lst in sorted(grupos.items(), key=lambda x: x[0][1]):
    conStock = [p for p in lst if st(p) > 0]
    enCero   = [p for p in lst if st(p) == 0]
    if not enCero:
        continue
    if not conStock:
        todos_cero.append((k[1], lst)); continue
    for p in enCero:
        gemelo = next((q for q in conStock
                       if norm(q.get('marca')) == norm(p.get('marca'))
                       and norm(q.get('proveedor')) == norm(p.get('proveedor'))), None)
        if gemelo:
            apagar.append((p, gemelo, k[1]))
        else:
            agotados.append((p, conStock[0], k[1]))

print('%s\n' % ('APLICANDO' if WRITE else 'SIMULACIÓN (no escribe)'))
print('=== A APAGAR — repetidos de verdad (mismo producto, misma marca, mismo proveedor, en 0) : %d' % len(apagar))
for p, g, key in apagar:
    print('  %-12s %-46s marca %-14s prov %-12s  →  se queda %s (stock %g)' % (
        p['codigo'], (p.get('descripcion') or '')[:46], (p.get('marca') or '-')[:14], (p.get('proveedor') or '-')[:12], g['codigo'], st(g)))
print('\n=== NO SE TOCAN — en 0 pero de otra marca u otro origen (están AGOTADOS, no sobran) : %d' % len(agotados))
for p, g, key in agotados:
    print('  %-12s %-46s marca %-14s prov %-12s  (el que tiene stock es %s, marca %s / %s)' % (
        p['codigo'], (p.get('descripcion') or '')[:46], (p.get('marca') or '-')[:14], (p.get('proveedor') or '-')[:12],
        g['codigo'], (g.get('marca') or '-'), (g.get('proveedor') or '-')))
print('\n=== NO SE TOCAN — grupos donde TODOS están en 0 : %d grupo(s)' % len(todos_cero))
for key, lst in todos_cero:
    print('  %-46s → %s' % (key[:46], ', '.join(p['codigo'] for p in lst)))

if not WRITE:
    print('\nNada escrito. Para aplicar:  python apagar_duplicados_en_cero.py --write')
    sys.exit(0)

print('\n--- archivando (activo=false + nota en la descripción) ---')
ok = err = 0
for p, g, key in apagar:
    desc = p.get('descripcion') or ''
    nueva = desc if desc.startswith('⛔') else ('⛔ NO USAR → ver ' + str(g['codigo']) + ' · ' + desc)
    try:
        stt, txt = escribir('productos?id=eq.' + urllib.parse.quote(str(p['id'])), 'PATCH',
                            {'activo': False, 'descripcion': nueva})
        if stt and 200 <= stt < 300:
            ok += 1; print('  ✓ %-12s archivado → ver %s' % (p['codigo'], g['codigo']))
        else:
            err += 1; print('  ✗ %-12s status=%s %s' % (p['codigo'], stt, txt[:120]))
    except Exception as e:
        err += 1; print('  ✗ %-12s %s' % (p['codigo'], e))
    time.sleep(0.25)
print('\nArchivados: %d   ·   con error: %d' % (ok, err))
