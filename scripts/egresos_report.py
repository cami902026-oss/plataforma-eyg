# -*- coding: utf-8 -*-
"""
Informe diario de Pagos (Egresos) — E&G Energy Group — VERSIÓN NUBE (GitHub Actions).

Reemplaza a la tarea local Informe_Pagos_EYG (que dependía de que el PC estuviera
encendido a las 6PM). Corre todos los días a las 6PM Colombia:

1. Descarga Comprobante-de-Egresos.xlsm del OneDrive de Andrea vía Microsoft Graph
   (mismo patrón probado de sync_inventory.py).
2. Reporta TODOS los pagos desde el último informe enviado (no solo los de hoy):
   si un día el informe no salió, el siguiente recoge lo pendiente y nada queda
   sin reportar. El estado vive en data/informe_pagos_estado.json (lo commitea
   el workflow).
3. Envía el correo vía Graph sendMail desde SENDER_EMAIL. SIEMPRE envía, incluso
   sin pagos ("No se registraron pagos") — así, si un día NO llega el correo,
   eso significa que algo falló y hay que revisar.

Uso:
    python scripts/egresos_report.py            -> informe real (avanza el estado)
    python scripts/egresos_report.py --prueba   -> solo al correo de prueba, NO avanza estado
"""
import os, sys, json, urllib.request, urllib.parse, datetime, argparse

ONEDRIVE_USER = 'andrea.bernal@eygenergygroup.com'
FILE_PATH     = 'CONTABILIDAD/Comprobante-de-Egresos.xlsm'
HOJA          = 'BD_EGRESO'
ESTADO_FILE   = 'data/informe_pagos_estado.json'
TMP_XLSM      = '_egresos_tmp.xlsm'

DESTINATARIOS = ['gerenciageneral@eygenergygroup.com', 'andrea.bernal@eygenergygroup.com']
CORREO_PRUEBA = ['cami902026@gmail.com']

NAVY = '#0F2B5B'; GREEN = '#2EAA4A'


def hoy_colombia():
    return (datetime.datetime.utcnow() - datetime.timedelta(hours=5)).date()


def pesos(n):
    try:
        return '$' + f'{int(round(float(n))):,}'.replace(',', '.')
    except Exception:
        return str(n)


def get_token():
    data = urllib.parse.urlencode({
        'grant_type':    'client_credentials',
        'client_id':     os.environ['MS_CLIENT_ID'].strip(),
        'client_secret': os.environ['MS_CLIENT_SECRET'].strip(),
        'scope':         'https://graph.microsoft.com/.default'
    }).encode()
    url = ('https://login.microsoftonline.com/' + os.environ['MS_TENANT_ID'].strip()
           + '/oauth2/v2.0/token')
    with urllib.request.urlopen(urllib.request.Request(url, data=data, method='POST')) as r:
        return json.loads(r.read())['access_token']


def descargar_excel(token):
    encoded = urllib.parse.quote(FILE_PATH, safe='/')
    url = ('https://graph.microsoft.com/v1.0/users/' + ONEDRIVE_USER
           + '/drive/root:/' + encoded + ':/content')
    req = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + token})
    with urllib.request.urlopen(req) as r:
        data = r.read()
    with open(TMP_XLSM, 'wb') as f:
        f.write(data)
    print(f'Excel descargado ({len(data)/1024:.1f} KB)')
    return TMP_XLSM


def leer_pagos(ruta, desde, hasta):
    """Pagos con desde < FECHA <= hasta (mismo parser del script local)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    ws = wb[HOJA]
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    if not filas:
        return []
    hdr = [str(c).strip().upper() if c is not None else '' for c in filas[0]]

    def buscar(*claves):
        for i, h in enumerate(hdr):
            for k in claves:
                if k in h:
                    return i
        return None

    i_fecha = buscar('FECHA')
    i_banco = buscar('BANCO')
    i_conc  = buscar('CONCEPTO')
    i_benef = buscar('BENEI', 'BENEF')
    i_valor = next((i for i, h in enumerate(hdr) if h.startswith('VALOR')), None)
    if None in (i_fecha, i_valor):
        raise RuntimeError(f'No encontré columnas FECHA/VALOR. Encabezados: {hdr}')

    pagos = []
    for r in filas[1:]:
        f = r[i_fecha] if i_fecha < len(r) else None
        if not isinstance(f, datetime.datetime):
            continue
        if not (desde < f.date() <= hasta):
            continue
        val = r[i_valor] if i_valor < len(r) else None
        if val in (None, '', 0):
            continue
        try:
            val = float(val)
        except Exception:
            continue
        pagos.append({
            'fecha':        f.date(),
            'beneficiario': (str(r[i_benef]).strip() if i_benef is not None and i_benef < len(r) and r[i_benef] else ''),
            'concepto':     (str(r[i_conc]).strip()  if i_conc  is not None and i_conc  < len(r) and r[i_conc]  else ''),
            'banco':        (str(r[i_banco]).strip() if i_banco is not None and i_banco < len(r) and r[i_banco] else ''),
            'valor':        val,
        })
    pagos.sort(key=lambda p: p['fecha'])
    return pagos


def armar_html(pagos, desde, hasta):
    rango_txt = (hasta.strftime('%d/%m/%Y') if desde + datetime.timedelta(days=1) >= hasta
                 else f'del {(desde + datetime.timedelta(days=1)).strftime("%d/%m/%Y")} al {hasta.strftime("%d/%m/%Y")}')
    if not pagos:
        return f"""
        <div style="font-family:Segoe UI,Arial,sans-serif;color:#222;max-width:680px">
          <h2 style="color:{NAVY};margin-bottom:4px">Resumen de Pagos — E&amp;G Energy Group</h2>
          <p style="color:#666;margin-top:0">{rango_txt}</p>
          <p style="font-size:15px;background:#f4f6fa;border-left:4px solid {NAVY};padding:12px 16px;border-radius:4px">
            No se registraron pagos en este periodo.
          </p>
          <p style="color:#999;font-size:12px">Informe automático 6:00 p.m. · Se envía todos los días: si un día no llega, hay que revisar el robot.</p>
        </div>"""

    total = sum(p['valor'] for p in pagos)
    por_banco = {}
    for p in pagos:
        b = p['banco'] or '(Sin banco)'
        por_banco[b] = por_banco.get(b, 0) + p['valor']
    dias = sorted(set(p['fecha'] for p in pagos))
    multi = len(dias) > 1

    filas_html = ''
    for d in dias:
        del_dia = [p for p in pagos if p['fecha'] == d]
        if multi:
            sub = sum(p['valor'] for p in del_dia)
            filas_html += f"""
        <tr style="background:#e9eef7;font-weight:bold;color:{NAVY}">
          <td style="padding:7px 10px" colspan="3">📅 {d.strftime('%A %d/%m/%Y').capitalize()} — {len(del_dia)} pago(s)</td>
          <td style="padding:7px 10px;text-align:right">{pesos(sub)}</td>
        </tr>"""
        for p in del_dia:
            filas_html += f"""
        <tr>
          <td style="padding:8px 10px;border-bottom:1px solid #eee">{p['beneficiario']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #eee">{p['concepto']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #eee">{p['banco']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap">{pesos(p['valor'])}</td>
        </tr>"""

    banco_html = ' &nbsp;·&nbsp; '.join(f'<b>{b}:</b> {pesos(v)}' for b, v in por_banco.items())
    aviso_multi = ('<p style="font-size:13px;color:#b7770d;background:#fff8e1;border-left:4px solid #E8A020;'
                   'padding:8px 12px;border-radius:4px">⚠️ Este informe incluye días anteriores que no se habían reportado.</p>') if multi else ''

    return f"""
    <div style="font-family:Segoe UI,Arial,sans-serif;color:#222;max-width:760px">
      <h2 style="color:{NAVY};margin-bottom:4px">Resumen de Pagos — E&amp;G Energy Group</h2>
      <p style="color:#666;margin-top:0">{rango_txt}</p>
      {aviso_multi}
      <p style="font-size:16px">
        Se registraron <b>{len(pagos)} pago(s)</b> por un total de
        <b style="color:{GREEN}">{pesos(total)}</b>.
      </p>
      <table style="border-collapse:collapse;width:100%;font-size:14px">
        <thead>
          <tr style="background:{NAVY};color:#fff;text-align:left">
            <th style="padding:9px 10px">Beneficiario</th>
            <th style="padding:9px 10px">Concepto</th>
            <th style="padding:9px 10px">Banco</th>
            <th style="padding:9px 10px;text-align:right">Valor</th>
          </tr>
        </thead>
        <tbody>{filas_html}
          <tr style="background:#f4f6fa;font-weight:bold">
            <td style="padding:9px 10px" colspan="3">TOTAL</td>
            <td style="padding:9px 10px;text-align:right">{pesos(total)}</td>
          </tr>
        </tbody>
      </table>
      <p style="font-size:13px;color:#444;margin-top:14px">Por banco: {banco_html}</p>
      <p style="color:#999;font-size:12px;margin-top:18px">
        Informe automático 6:00 p.m. desde la nube (no depende de ningún PC) · Fuente: Comprobante-de-Egresos.xlsm
      </p>
    </div>"""


def enviar_correo(token, destinatarios, asunto, html):
    sender = os.environ['SENDER_EMAIL'].strip()
    payload = json.dumps({
        'message': {
            'subject': asunto,
            'body': {'contentType': 'HTML', 'content': html},
            'toRecipients': [{'emailAddress': {'address': d}} for d in destinatarios]
        },
        'saveToSentItems': True
    }).encode('utf-8')
    req = urllib.request.Request(
        f'https://graph.microsoft.com/v1.0/users/{sender}/sendMail',
        data=payload, method='POST',
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as r:
        print('Correo enviado a: ' + ', '.join(destinatarios))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--prueba', action='store_true', help='solo al correo de prueba; NO avanza el estado')
    args = ap.parse_args()

    hoy = hoy_colombia()
    try:
        with open(ESTADO_FILE, encoding='utf-8') as f:
            estado = json.load(f)
        desde = datetime.date.fromisoformat(estado.get('ultimaFechaReportada'))
    except Exception:
        desde = hoy - datetime.timedelta(days=1)
    if desde >= hoy:
        desde = hoy - datetime.timedelta(days=1)   # re-ejecución el mismo día: reporta hoy de nuevo

    print(f'Rango a reportar: {desde.isoformat()} (exclusivo) → {hoy.isoformat()} (inclusive)')
    token = get_token()
    ruta = descargar_excel(token)
    pagos = leer_pagos(ruta, desde, hoy)
    total = sum(p['valor'] for p in pagos)
    print(f'{len(pagos)} pago(s), total {pesos(total)}')

    html = armar_html(pagos, desde, hoy)
    asunto = f'Resumen de Pagos E&G — {hoy.strftime("%d/%m/%Y")}' + (f' ({len(pagos)} pagos, {pesos(total)})' if pagos else ' (sin pagos)')
    dest = CORREO_PRUEBA if args.prueba else DESTINATARIOS
    if args.prueba:
        asunto = '[PRUEBA] ' + asunto
    enviar_correo(token, dest, asunto, html)

    if not args.prueba:
        with open(ESTADO_FILE, 'w', encoding='utf-8') as f:
            json.dump({'ultimaFechaReportada': hoy.isoformat(),
                       'ultimoEnvio': datetime.datetime.utcnow().isoformat() + 'Z',
                       'pagosReportados': len(pagos)}, f, ensure_ascii=False, indent=2)
        print('Estado actualizado: ultimaFechaReportada =', hoy.isoformat())


if __name__ == '__main__':
    main()
