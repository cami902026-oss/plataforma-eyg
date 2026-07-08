# -*- coding: utf-8 -*-
"""
🩺 SALUD DEL SISTEMA ENERGY — Plan de Sostenibilidad (sesión 1, 2026-07-08)
============================================================================
Revisa la salud completa de la plataforma y envía el informe por correo:

1. BACKUPS: frescura y completitud (reusa scripts/verificar_backups.py).
2. TAMAÑOS de data/*.json (lo que viaja al navegador) + tendencia vs mes anterior.
3. SUPABASE: filas por tabla (vigila el límite del plan gratuito).
4. ROBOTS: workflows de GitHub que fallaron en los últimos 30 días.
5. REPO: tamaño del repositorio.

Modos (env MODO o automático por fecha):
  informe  -> SIEMPRE envía el informe completo (día 1 de mes, 7AM).
  chequeo  -> solo envía correo SI HAY ALERTAS (lunes 7AM). Silencio = todo bien.
  prueba   -> informe completo solo a cami902026@gmail.com, asunto [PRUEBA].

Estado previo (para tendencias): data/salud_estado.json (lo commitea el workflow).
"""
import os, sys, json, glob, datetime, urllib.request, urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import verificar_backups as vb   # reutiliza backup_info / sb_count / TABLES

REPO_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
ESTADO     = os.path.join(REPO_DIR, 'data', 'salud_estado.json')
GH_REPO    = 'cami902026-oss/plataforma-eyg'
DEST_REAL  = ['andrea.bernal@eygenergygroup.com', 'gerenciageneral@eygenergygroup.com']
DEST_PRUEBA= ['cami902026@gmail.com']
SENDER     = os.environ.get('EGRESOS_SENDER', 'info@eygenergygroup.com').strip()
NAVY='#0F2B5B'; GOLD='#E8A020'; GREEN='#2EAA4A'; RED='#b91c1c'; AMBER='#b7770d'

TABLAS_SB = ['productos','kardex','cotizaciones','cotizacion_items','remisiones',
             'plan_compras','oc_compras','proveedores','familias']
LIMITE_DATA_MB   = 8     # data/ total (proxy del localStorage del navegador)
LIMITE_ARCHIVO_MB= 2     # un solo json activo no debería pasar de esto


def kb(n): return f'{n/1024:.0f} KB' if n < 1048576 else f'{n/1048576:.1f} MB'


def check_backups(alertas):
    filas = []
    for t in vb.TABLES:
        b, _ = vb.backup_info(t)
        s = vb.sb_count(t)
        estado, nivel = 'OK', 'ok'
        if b is None:
            estado, nivel = 'FALTA BACKUP', 'rojo'; alertas.append(('rojo', f'Backup de {t}: NO EXISTE'))
        elif b == 'ILEGIBLE':
            estado, nivel = 'ILEGIBLE', 'rojo'; alertas.append(('rojo', f'Backup de {t}: archivo ilegible'))
        elif b == 0 and isinstance(s, int) and s > 0:
            estado, nivel = 'VACÍO (¡hay datos!)', 'rojo'; alertas.append(('rojo', f'Backup de {t} vacío pero Supabase tiene {s} filas'))
        elif isinstance(s, int) and isinstance(b, int) and b < s * 0.5 and s > 20:
            estado, nivel = 'POSIBLE TRUNCADO', 'amarillo'; alertas.append(('amarillo', f'Backup de {t} ({b}) muy inferior a Supabase ({s})'))
        filas.append((t, b, s, estado, nivel))
    # frescura
    edad = None
    try:
        with open(os.path.join(vb.BK_DIR, '_resumen.json'), encoding='utf-8') as f:
            res = json.load(f)
        fecha = datetime.datetime.fromisoformat(str(res.get('actualizado', res.get('fecha', '')))[:19])
        edad = (datetime.datetime.utcnow() - fecha).days
    except Exception:
        alertas.append(('amarillo', 'No se pudo leer la fecha del último respaldo'))
    if edad is not None and edad > 2:
        alertas.append(('rojo', f'El respaldo diario tiene {edad} días — la Action nocturna pudo dejar de correr'))
    return filas, edad


def check_data_sizes(alertas):
    prev = {}
    try:
        with open(ESTADO, encoding='utf-8') as f:
            prev = json.load(f).get('tamanos', {})
    except Exception:
        pass
    filas, total, tam_actual = [], 0, {}
    for p in sorted(glob.glob(os.path.join(REPO_DIR, 'data', '*.json'))):
        nombre = os.path.basename(p)
        if nombre.startswith(('diag_', 'salud_estado')): continue
        n = os.path.getsize(p); total += n; tam_actual[nombre] = n
        delta = n - prev.get(nombre, n)
        filas.append((nombre, n, delta))
        if n > LIMITE_ARCHIVO_MB * 1048576:
            alertas.append(('amarillo', f'{nombre} pesa {kb(n)} — candidato a archivado (frena el navegador)'))
    filas.sort(key=lambda x: -x[1])
    if total > LIMITE_DATA_MB * 1048576:
        alertas.append(('rojo', f'data/ total {kb(total)} — el almacenamiento del navegador (5-10MB) está en riesgo'))
    return filas, total, tam_actual


def check_supabase(alertas):
    filas = []
    for t in TABLAS_SB:
        try:
            n = vb.sb_count(t)
        except Exception:
            n = '?'
        filas.append((t, n))
        if n == '?' or (t == 'productos' and isinstance(n, int) and n == 0):
            alertas.append(('rojo', f'Supabase: no se pudo leer la tabla {t}'))
    return filas


def check_workflows(alertas):
    tok = os.environ.get('GITHUB_TOKEN', '').strip()
    if not tok: return []
    desde = (datetime.datetime.utcnow() - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    url = (f'https://api.github.com/repos/{GH_REPO}/actions/runs'
           f'?status=failure&per_page=100&created=%3E{desde}')
    try:
        req = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + tok,
                                                   'Accept': 'application/vnd.github+json'})
        with urllib.request.urlopen(req, timeout=45) as r:
            runs = json.loads(r.read()).get('workflow_runs', [])
    except Exception as e:
        alertas.append(('amarillo', 'No se pudieron consultar los robots: ' + str(e)[:60]))
        return []
    cnt = {}
    for x in runs:
        cnt[x.get('name', '?')] = cnt.get(x.get('name', '?'), 0) + 1
    for nombre, n in sorted(cnt.items(), key=lambda x: -x[1]):
        alertas.append(('amarillo', f'Robot "{nombre}" falló {n} vez/veces en 30 días'))
    return sorted(cnt.items(), key=lambda x: -x[1])


def repo_size(alertas):
    tok = os.environ.get('GITHUB_TOKEN', '').strip()
    if not tok: return None
    try:
        req = urllib.request.Request(f'https://api.github.com/repos/{GH_REPO}',
                                     headers={'Authorization': 'Bearer ' + tok})
        with urllib.request.urlopen(req, timeout=30) as r:
            n = json.loads(r.read()).get('size', 0)   # KB
        if n > 800 * 1024:
            alertas.append(('amarillo', f'El repositorio pesa {n//1024} MB — considerar limpieza de historial'))
        return n
    except Exception:
        return None


def armar_html(sem, alertas, bk_filas, bk_edad, data_filas, data_total, sb_filas, wf_fallidos, repo_kb, hoy):
    color = GREEN if sem == '🟢' else (AMBER if sem == '🟡' else RED)
    titulo = {'🟢': 'TODO EN ORDEN', '🟡': 'CON AVISOS', '🔴': 'REQUIERE ATENCIÓN'}[sem]
    al_html = ''.join(f'<li style="color:{RED if n=="rojo" else AMBER};margin-bottom:4px;">{"🔴" if n=="rojo" else "🟡"} {t}</li>'
                      for n, t in alertas) or '<li style="color:#2e7d32;">Sin alertas — todo dentro de lo normal ✅</li>'
    bk_html = ''.join(f'<tr><td style="padding:5px 10px;border-bottom:1px solid #eee;">{t}</td>'
                      f'<td style="padding:5px 10px;border-bottom:1px solid #eee;text-align:right;">{b if b is not None else "—"}</td>'
                      f'<td style="padding:5px 10px;border-bottom:1px solid #eee;text-align:right;">{s}</td>'
                      f'<td style="padding:5px 10px;border-bottom:1px solid #eee;color:{GREEN if nv=="ok" else (AMBER if nv=="amarillo" else RED)};font-weight:600;">{e}</td></tr>'
                      for t, b, s, e, nv in bk_filas)
    dt_html = ''.join(f'<tr><td style="padding:5px 10px;border-bottom:1px solid #eee;">{n}</td>'
                      f'<td style="padding:5px 10px;border-bottom:1px solid #eee;text-align:right;">{kb(s)}</td>'
                      f'<td style="padding:5px 10px;border-bottom:1px solid #eee;text-align:right;color:{"#888" if abs(d)<1024 else (RED if d>0 else GREEN)};">{"+" if d>0 else ""}{kb(abs(d)) if abs(d)>=1024 else "="}</td></tr>'
                      for n, s, d in data_filas[:10])
    sb_html = ' · '.join(f'<b>{t}</b>: {n}' for t, n in sb_filas)
    wf_html = ('Sin fallos en 30 días ✅' if not wf_fallidos
               else ' · '.join(f'<b>{n}</b>: {c} fallo(s)' for n, c in wf_fallidos))
    return f"""<div style="font-family:Segoe UI,Arial,sans-serif;color:#222;max-width:780px">
  <div style="background:{NAVY};border-radius:12px 12px 0 0;padding:18px 24px;">
    <span style="font-size:22px;font-weight:900;color:#fff;">🩺 Salud del Sistema ENERGY</span>
    <span style="float:right;font-size:20px;font-weight:900;color:{color};background:#fff;border-radius:10px;padding:2px 14px;">{sem} {titulo}</span>
  </div>
  <div style="border:1px solid #dde3ee;border-top:none;border-radius:0 0 12px 12px;padding:20px 24px;">
    <p style="color:#666;margin-top:0;">{hoy.strftime('%d/%m/%Y')} · Informe automático del plan de sostenibilidad</p>
    <h3 style="color:{NAVY};margin-bottom:6px;">⚠️ Alertas</h3>
    <ul style="margin-top:4px;">{al_html}</ul>
    <h3 style="color:{NAVY};margin-bottom:6px;">💾 Respaldos (backup diario 1:37 AM{f' · último hace {bk_edad} día(s)' if bk_edad is not None else ''})</h3>
    <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="background:{NAVY};color:#fff;">
      <th style="padding:6px 10px;text-align:left;">Tabla</th><th style="padding:6px 10px;">Backup</th><th style="padding:6px 10px;">Supabase</th><th style="padding:6px 10px;text-align:left;">Estado</th></tr>{bk_html}</table>
    <h3 style="color:{NAVY};margin-bottom:6px;">📦 Datos que viajan al navegador (data/ = {kb(data_total)} · umbral {LIMITE_DATA_MB} MB)</h3>
    <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="background:{NAVY};color:#fff;">
      <th style="padding:6px 10px;text-align:left;">Archivo</th><th style="padding:6px 10px;">Tamaño</th><th style="padding:6px 10px;">vs. anterior</th></tr>{dt_html}</table>
    <h3 style="color:{NAVY};margin-bottom:6px;">🗄️ Supabase (plan gratuito: 500 MB)</h3>
    <p style="font-size:13px;">{sb_html}</p>
    <h3 style="color:{NAVY};margin-bottom:6px;">🤖 Robots (GitHub Actions, últimos 30 días)</h3>
    <p style="font-size:13px;">{wf_html}</p>
    {f'<p style="font-size:13px;">📁 Repositorio: {repo_kb//1024} MB</p>' if repo_kb else ''}
    <p style="color:#999;font-size:11px;border-top:3px solid {GOLD};padding-top:8px;margin-top:16px;">
      ⚡ ENERGY — Informe mensual (día 1, 7AM) + chequeo silencioso los lunes (solo escribe si hay alertas)</p>
  </div></div>"""


def enviar(destinatarios, asunto, html):
    data = urllib.parse.urlencode({
        'grant_type': 'client_credentials',
        'client_id': os.environ['MS_CLIENT_ID'].strip(),
        'client_secret': os.environ['MS_CLIENT_SECRET'].strip(),
        'scope': 'https://graph.microsoft.com/.default'}).encode()
    url = 'https://login.microsoftonline.com/' + os.environ['MS_TENANT_ID'].strip() + '/oauth2/v2.0/token'
    with urllib.request.urlopen(urllib.request.Request(url, data=data, method='POST')) as r:
        token = json.loads(r.read())['access_token']
    payload = json.dumps({'message': {'subject': asunto,
        'body': {'contentType': 'HTML', 'content': html},
        'toRecipients': [{'emailAddress': {'address': d}} for d in destinatarios]},
        'saveToSentItems': True}).encode('utf-8')
    req = urllib.request.Request(f'https://graph.microsoft.com/v1.0/users/{SENDER}/sendMail',
        data=payload, method='POST',
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as r:
        print('Correo enviado a:', ', '.join(destinatarios))


def main():
    hoy = datetime.datetime.utcnow() - datetime.timedelta(hours=5)
    modo = os.environ.get('MODO', '').strip() or ('informe' if hoy.day == 1 else 'chequeo')
    print('Modo:', modo, '| Fecha Colombia:', hoy.strftime('%Y-%m-%d %H:%M'))

    alertas = []
    bk_filas, bk_edad = check_backups(alertas)
    data_filas, data_total, tam_actual = check_data_sizes(alertas)
    sb_filas = check_supabase(alertas)
    wf_fallidos = check_workflows(alertas)
    repo_kb = repo_size(alertas)

    sem = '🟢' if not alertas else ('🔴' if any(n == 'rojo' for n, _ in alertas) else '🟡')
    print('Semáforo:', sem, '| Alertas:', len(alertas))
    for n, t in alertas: print('  -', n.upper(), t)

    html = armar_html(sem, alertas, bk_filas, bk_edad, data_filas, data_total, sb_filas, wf_fallidos, repo_kb, hoy)
    asunto = f'🩺 Salud del Sistema ENERGY {sem} — {hoy.strftime("%d/%m/%Y")}' + (f' ({len(alertas)} alertas)' if alertas else ' (todo en orden)')

    if modo == 'prueba':
        enviar(DEST_PRUEBA, '[PRUEBA] ' + asunto, html)
    elif modo == 'informe':
        enviar(DEST_REAL, asunto, html)
    elif alertas:   # chequeo: solo molesta si hay algo mal
        enviar(DEST_REAL, asunto, html)
    else:
        print('Chequeo sin alertas: no se envía correo (silencio = todo bien).')

    if modo != 'prueba':
        with open(ESTADO, 'w', encoding='utf-8') as f:
            json.dump({'fecha': hoy.isoformat(), 'tamanos': tam_actual,
                       'alertas': len(alertas), 'semaforo': sem}, f, ensure_ascii=False, indent=2)
        print('Estado guardado en', ESTADO)


if __name__ == '__main__':
    main()
