"""
ENERGY — Reporte Diario de Órdenes de Pedido
=============================================
Lee las órdenes desde ordenes.json en GitHub,
genera un informe HTML y lo envía por correo
vía Microsoft Graph API.

Ejecutado automáticamente por GitHub Actions
Lunes a Viernes a las 5:00 PM Colombia (UTC-5 = 22:00 UTC)
"""

import json
import os
import sys
import base64
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime


DAYS_ES   = {'Monday':'Lunes','Tuesday':'Martes','Wednesday':'Miércoles',
              'Thursday':'Jueves','Friday':'Viernes','Saturday':'Sábado','Sunday':'Domingo'}
MONTHS_ES = {'January':'enero','February':'febrero','March':'marzo','April':'abril',
              'May':'mayo','June':'junio','July':'julio','August':'agosto',
              'September':'septiembre','October':'octubre','November':'noviembre','December':'diciembre'}

GH_OWNER = 'cami902026-oss'
GH_REPO  = 'plataforma-eyg'

POC_STAGES = [
    {'key': 'compra',   'icon': '🛒', 'label': 'Compra'},
    {'key': 'entrega',  'icon': '🚚', 'label': 'Entrega'},
    {'key': 'cert',     'icon': '📋', 'label': 'Certificado'},
    {'key': 'factura',  'icon': '💰', 'label': 'Facturación'},
]


# ─── 1. LEER CREDENCIALES ─────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = os.path.join(os.path.dirname(__file__), 'cowork_config.json')
    if not os.path.exists(config_path):
        return {}
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ─── 2. DESCARGAR ordenes.json DESDE GITHUB ───────────────────────────────────

def load_ordenes_from_github(gh_token: str) -> list:
    """Descarga ordenes.json del repositorio GitHub via API."""
    url = f'https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/contents/ordenes.json'
    req = urllib.request.Request(url, headers={
        'Authorization': f'Bearer {gh_token}',
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'EnergyBot/1.0'
    })
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read())
            content = base64.b64decode(data['content']).decode('utf-8')
            ordenes = json.loads(content)
            print(f'✅ ordenes.json descargado: {len(ordenes)} órdenes')
            return ordenes
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f'ERROR descargando ordenes.json: {e.code} — {body}')
        if e.code == 404:
            print('⚠️  ordenes.json no existe en el repo aún. Se enviará reporte vacío.')
            return []
        sys.exit(1)
    except Exception as ex:
        print(f'ERROR inesperado: {ex}')
        return []


# ─── 3. GENERAR HTML DEL REPORTE ──────────────────────────────────────────────

def badge_estado(estado: str) -> str:
    colores = {
        'activo':     ('background:#e6f4ea;color:#1e7e34', 'Activo'),
        'completado': ('background:#e8f0fe;color:#1a56db', 'Completado'),
        'cancelado':  ('background:#fce8e6;color:#c0392b', 'Cancelado'),
    }
    style, label = colores.get(estado, ('background:#f0f0f0;color:#555', estado.title()))
    return f"<span style='padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;{style};'>{label}</span>"


def stage_dot(stage: dict, idx: int) -> str:
    s     = stage.get('s', 'pending')
    fecha = stage.get('f', '')
    nota  = stage.get('n', '')
    st    = POC_STAGES[idx]
    # done = s=='done' ó tiene fecha (igual que frontend)
    visually_done = (s == 'done' or bool(fecha.strip()))
    color = '#2EAA4A' if visually_done else ('#E8A020' if s == 'active' else '#8899bb')
    bg    = '#e6f4ea' if visually_done else ('#fff8e1' if s == 'active' else '#1e3a6e')
    fecha_txt = f"<br><span style='font-size:10px;color:#8899bb;'>{'/'.join(reversed(fecha.split('-')))}</span>" if fecha else ''
    nota_txt  = f"<br><span style='font-size:10px;color:#8899bb;' title='{nota}'>💬 {nota[:25]}{'…' if len(nota)>25 else ''}</span>" if nota else ''
    return (
        f"<td style='text-align:center;padding:4px 8px;'>"
        f"<div style='display:inline-block;background:{bg};border-radius:50%;width:28px;height:28px;"
        f"line-height:28px;font-size:14px;color:{color};border:2px solid {color};'>{st['icon']}</div>"
        f"<div style='font-size:10px;color:{color};font-weight:600;margin-top:2px;'>{st['label']}</div>"
        f"{fecha_txt}{nota_txt}</td>"
    )


def is_stage_done(st: dict) -> bool:
    """Una etapa está completa si s=='done' O si tiene fecha registrada (igual que el frontend)."""
    return st.get('s') == 'done' or bool(st.get('f', '').strip())


def format_money(v) -> str:
    """Formatea un valor monetario en COP. Devuelve '—' si es 0 o vacío."""
    try:
        v = float(v or 0)
    except (TypeError, ValueError):
        v = 0
    if v >= 1_000_000:
        return f'${v/1_000_000:.1f}M'
    if v >= 1_000:
        return f'${int(v/1_000)}K'
    if v > 0:
        return '$' + f'{int(v):,}'.replace(',', '.')
    return '—'


def he_icon(o: dict) -> str:
    """Devuelve ícono visual del estado de Hoja de Entrada."""
    he = o.get('hojaEntrada') or {}
    if not he.get('requerida'):
        return '—'
    return '✅' if he.get('fecha') else '⏳'


def he_label(o: dict) -> str:
    he = o.get('hojaEntrada') or {}
    if not he.get('requerida'):
        return 'No requiere'
    if he.get('fecha'):
        f = he.get('fecha', '')
        return f"Recibida {'/'.join(reversed(f.split('-')))}"
    return 'Pendiente'


def fecha_ingreso(o: dict) -> str:
    """Devuelve la fecha de ingreso de la OP (fechaIngreso o createdAt)."""
    f = (o.get('fechaIngreso') or '').strip()
    if f:
        return f
    ca = o.get('createdAt')
    if ca:
        try: return str(ca)[:10]
        except: pass
    return ''


def dias_desde_ingreso(o: dict) -> int:
    """Calcula días transcurridos desde fecha de ingreso. -1 si no hay fecha."""
    from datetime import date
    f = fecha_ingreso(o)
    if not f:
        return -1
    try:
        y, m, d = f[:10].split('-')
        d_ingreso = date(int(y), int(m), int(d))
        return max(0, (date.today() - d_ingreso).days)
    except Exception:
        return -1


def dias_html(o: dict) -> str:
    """HTML para mostrar días transcurridos con color según alerta."""
    d = dias_desde_ingreso(o)
    if d < 0:
        return '<span style="color:#8899bb;">—</span>'
    color = '#86efac' if d <= 15 else ('#fbbf24' if d <= 30 else '#f87171')
    return f'<span style="color:{color};font-weight:700;">{d} d</span>'


def get_etapa_actual(orden: dict) -> int:
    """Devuelve el índice (0-3) de la etapa actual de una orden activa.
    Busca la primera etapa 'active'; si no hay, la primera no-completada."""
    stages = orden.get('stages') or []
    # Asegurar 4 elementos
    while len(stages) < 4:
        stages.append({})
    # Primero: buscar etapa marcada 'active'
    for i, st in enumerate(stages[:4]):
        if st.get('s') == 'active':
            return i
    # Luego: primera etapa que NO esté done (done = s=='done' ó tiene fecha)
    for i, st in enumerate(stages[:4]):
        if not is_stage_done(st):
            return i
    return 3  # todas done → se queda en facturación


def build_resumen_etapas(activos: list) -> str:
    """Genera el bloque HTML de resumen por etapa para órdenes activas.
    Para cada etapa, cuenta las órdenes que AÚN NO la tienen completada."""
    grupos = {0: [], 1: [], 2: [], 3: []}
    for o in activos:
        stages = o.get('stages') or []
        while len(stages) < 4:
            stages.append({})
        for i in range(4):
            # Si esta etapa no está done → la orden aparece en este grupo
            if not is_stage_done(stages[i]):
                grupos[i].append(o.get('num') or o.get('id', '—'))

    filas_html = ''
    for idx, st in enumerate(POC_STAGES):
        items = grupos[idx]
        cantidad = len(items)
        lista_txt = ', '.join(str(x) for x in items) if items else '—'
        color_num = '#E8A020' if cantidad > 0 else '#8899bb'
        filas_html += f"""
        <tr style='border-bottom:1px solid #1e3a6e;'>
          <td style='padding:10px 14px;font-size:18px;text-align:center;width:40px;'>{st['icon']}</td>
          <td style='padding:10px 8px;font-size:13px;font-weight:700;color:#e2e8f0;'>{st['label']}</td>
          <td style='padding:10px 8px;text-align:center;'>
            <span style='font-size:22px;font-weight:900;color:{color_num};'>{cantidad}</span>
            <div style='font-size:10px;color:#8899bb;'>orden{'es' if cantidad != 1 else ''}</div>
          </td>
          <td style='padding:10px 14px;font-size:11px;color:#93c5fd;'>{lista_txt}</td>
        </tr>"""

    return f"""
    <div style='margin-bottom:24px;'>
      <div style='font-size:11px;color:#8899bb;font-weight:700;text-transform:uppercase;
                  letter-spacing:1px;margin-bottom:10px;'>📊 Órdenes Pendientes por Etapa</div>
      <table style='width:100%;border-collapse:collapse;background:#0d1f3c;
                    border:1px solid #1e3a6e;border-radius:10px;overflow:hidden;'>
        <thead>
          <tr style='background:#0F2B5B;'>
            <th style='padding:8px;'></th>
            <th style='padding:8px;text-align:left;font-size:11px;color:#8899bb;
                       text-transform:uppercase;letter-spacing:1px;'>Etapa</th>
            <th style='padding:8px;text-align:center;font-size:11px;color:#8899bb;
                       text-transform:uppercase;letter-spacing:1px;'>Cantidad</th>
            <th style='padding:8px;text-align:left;font-size:11px;color:#8899bb;
                       text-transform:uppercase;letter-spacing:1px;'>Órdenes</th>
          </tr>
        </thead>
        <tbody>{filas_html}</tbody>
      </table>
    </div>"""


def build_report_html(ordenes: list, date_str: str) -> str:
    # Excluir entradas eliminadas (soft-delete / guardias del sistema)
    ordenes = [o for o in ordenes if not o.get('deleted')]
    activos     = [o for o in ordenes if o.get('estado') == 'activo']
    completados = [o for o in ordenes if o.get('estado') == 'completado']
    cancelados  = [o for o in ordenes if o.get('estado') == 'cancelado']

    # Sumar valores totales
    def sum_val(lista):
        total = 0.0
        for o in lista:
            try: total += float(o.get('valor') or 0)
            except: pass
        return total
    valor_activos     = sum_val(activos)
    valor_completados = sum_val(completados)

    # Órdenes que requieren HE y aún no la tienen (incluye completadas prematuramente)
    pendientes_he = [o for o in (activos + completados)
                     if (o.get('hojaEntrada') or {}).get('requerida')
                     and not (o.get('hojaEntrada') or {}).get('fecha')]

    resumen_etapas = build_resumen_etapas(activos) if activos else ''

    resumen = f"""
    <div style='display:flex;gap:12px;flex-wrap:wrap;margin-bottom:24px;'>
      <div style='flex:1;min-width:120px;background:#0d1f3c;border:1px solid #1e3a6e;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:28px;font-weight:900;color:#fff;'>{len(ordenes)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>Total</div>
      </div>
      <div style='flex:1;min-width:120px;background:#0d1f3c;border:1px solid #1e3a6e;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:28px;font-weight:900;color:#2EAA4A;'>{len(activos)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>Activas</div>
      </div>
      <div style='flex:1;min-width:120px;background:#0d1f3c;border:1px solid #1e3a6e;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:28px;font-weight:900;color:#1a56db;'>{len(completados)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>Completadas</div>
      </div>
      <div style='flex:1;min-width:120px;background:#0d1f3c;border:1px solid #1e3a6e;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:28px;font-weight:900;color:#e53e3e;'>{len(cancelados)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>Canceladas</div>
      </div>
      <div style='flex:1;min-width:140px;background:#0d1f3c;border:1px solid #2EAA4A;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:22px;font-weight:900;color:#86efac;'>{format_money(valor_activos)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>💵 Valor Activas</div>
      </div>
      <div style='flex:1;min-width:140px;background:#0d1f3c;border:1px solid #1a56db;border-radius:10px;padding:16px;text-align:center;'>
        <div style='font-size:22px;font-weight:900;color:#93c5fd;'>{format_money(valor_completados)}</div>
        <div style='font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;margin-top:4px;'>💵 Valor Completadas</div>
      </div>
    </div>"""

    # Mostrar órdenes con al menos una etapa pendiente (activas o completadas-prematuras),
    # excluyendo solo las canceladas.
    # (done = s=='done' ó tiene fecha, igual que el frontend)
    def _tiene_pendiente(o):
        stages = o.get('stages') or []
        while len(stages) < 4:
            stages.append({})
        return any(not is_stage_done(stages[i]) for i in range(4))

    pendientes = [o for o in (activos + completados) if _tiene_pendiente(o)]

    if not pendientes:
        cuerpo = "<div style='text-align:center;padding:40px;color:#8899bb;'>✅ Todas las órdenes activas están al día.</div>"
    else:
        filas = ''
        for o in reversed(pendientes):
            stages = o.get('stages', [{},{},{},{}])
            while len(stages) < 4:
                stages.append({})
            dots = ''.join(stage_dot(stages[i], i) for i in range(4))
            fing = fecha_ingreso(o)
            fing_fmt = '/'.join(reversed(fing.split('-'))) if fing else '—'
            filas += f"""
            <tr style='border-bottom:1px solid #1e3a6e;'>
              <td style='padding:12px 10px;'>
                <div style='font-weight:700;color:#fff;font-size:13px;'>{o.get('num','—')}</div>
                <div style='color:#8899bb;font-size:11px;margin-top:2px;'>{o.get('cliente','')}</div>
              </td>
              <td style='padding:12px 10px;color:#d0d9f0;font-size:12px;max-width:200px;'>
                {(o.get('desc') or '')[:80]}{'…' if len(o.get('desc',''))>80 else ''}
              </td>
              <td style='padding:12px 6px;text-align:center;color:#cbd5ff;font-size:11px;'>{fing_fmt}</td>
              <td style='padding:12px 6px;text-align:center;font-size:12px;'>{dias_html(o)}</td>
              <td style='padding:12px 10px;text-align:center;color:#86efac;font-size:12px;font-weight:700;'>{format_money(o.get('valor'))}</td>
              <td style='padding:12px 10px;text-align:center;'>{badge_estado(o.get('estado','activo'))}</td>
              {dots}
              <td style='padding:12px 6px;text-align:center;font-size:14px;' title='{he_label(o)}'>{he_icon(o)}</td>
            </tr>"""

        cuerpo = f"""
        <table style='width:100%;border-collapse:collapse;background:#0d1f3c;border-radius:10px;overflow:hidden;'>
          <thead>
            <tr style='background:#0F2B5B;'>
              <th style='padding:10px;text-align:left;font-size:11px;color:#E8A020;text-transform:uppercase;letter-spacing:1px;' colspan='11'>⚠️ Órdenes con Etapas Pendientes ({len(pendientes)})</th></tr><tr style='background:#0F2B5B;'>
              <th style='padding:10px;text-align:left;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>OP / Proyecto</th>
              <th style='padding:10px;text-align:left;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>Descripción</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>📅 Ingreso</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>⏱ Días</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>💵 Valor</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>Estado</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>🛒 Compra</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>🚚 Entrega</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>📋 Cert.</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>💰 Factura</th>
              <th style='padding:10px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>📋 HE</th>
            </tr>
          </thead>
          <tbody>{filas}</tbody>
        </table>"""

    # Sección de órdenes pendientes de Hoja de Entrada
    if pendientes_he:
        he_rows = ''
        for o in pendientes_he:
            he_rows += f"""
            <tr style='border-bottom:1px solid #1e3a6e;'>
              <td style='padding:10px 14px;'>
                <div style='font-weight:700;color:#fff;font-size:13px;'>{o.get('num','—')}</div>
                <div style='color:#8899bb;font-size:11px;margin-top:2px;'>{o.get('cliente','')}</div>
              </td>
              <td style='padding:10px 14px;color:#d0d9f0;font-size:12px;'>
                {(o.get('desc') or '')[:80]}{'…' if len(o.get('desc',''))>80 else ''}
              </td>
              <td style='padding:10px 14px;text-align:center;color:#86efac;font-size:12px;font-weight:700;'>{format_money(o.get('valor'))}</td>
              <td style='padding:10px 14px;text-align:center;color:#fbbf24;font-size:12px;font-weight:700;'>⏳ Pendiente</td>
            </tr>"""
        seccion_he = f"""
        <div style='margin-top:24px;margin-bottom:24px;'>
          <div style='font-size:11px;color:#E8A020;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;'>📋 Órdenes Pendientes de Hoja de Entrada ({len(pendientes_he)})</div>
          <table style='width:100%;border-collapse:collapse;background:#0d1f3c;border:1px solid #E8A020;border-radius:10px;overflow:hidden;'>
            <thead>
              <tr style='background:#0F2B5B;'>
                <th style='padding:10px 14px;text-align:left;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>OP / Proyecto</th>
                <th style='padding:10px 14px;text-align:left;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>Descripción</th>
                <th style='padding:10px 14px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>💵 Valor</th>
                <th style='padding:10px 14px;text-align:center;font-size:11px;color:#8899bb;text-transform:uppercase;letter-spacing:1px;'>Estado HE</th>
              </tr>
            </thead>
            <tbody>{he_rows}</tbody>
          </table>
        </div>"""
    else:
        seccion_he = ''

    # ── Sección "Histórico completo" — TODAS las OCs no eliminadas (agrupadas por estado) ──
    def _bloque_grupo(titulo, lista, color_titulo):
        if not lista:
            return ''
        filas_hist = ''
        for o in reversed(lista):
            fing = fecha_ingreso(o)
            fing_fmt = '/'.join(reversed(fing.split('-'))) if fing else '—'
            filas_hist += f"""
            <tr style='border-bottom:1px solid #1e3a6e;'>
              <td style='padding:8px 10px;'>
                <div style='font-weight:700;color:#fff;font-size:12px;'>{o.get('num','—')}</div>
                <div style='color:#8899bb;font-size:10px;'>{(str(o.get('cliente','')))[:35]}</div>
              </td>
              <td style='padding:8px 10px;color:#d0d9f0;font-size:11px;max-width:200px;'>{(o.get('desc') or '')[:60]}{'…' if len(o.get('desc',''))>60 else ''}</td>
              <td style='padding:8px 6px;text-align:center;color:#cbd5ff;font-size:11px;'>{fing_fmt}</td>
              <td style='padding:8px 6px;text-align:center;font-size:11px;'>{dias_html(o)}</td>
              <td style='padding:8px 10px;text-align:center;color:#86efac;font-size:11px;font-weight:700;'>{format_money(o.get('valor'))}</td>
              <td style='padding:8px 6px;text-align:center;font-size:13px;' title='{he_label(o)}'>{he_icon(o)}</td>
            </tr>"""
        return f"""
        <div style='margin-top:18px;'>
          <div style='font-size:11px;color:{color_titulo};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;'>{titulo} ({len(lista)})</div>
          <table style='width:100%;border-collapse:collapse;background:#0d1f3c;border:1px solid #1e3a6e;border-radius:8px;overflow:hidden;'>
            <thead>
              <tr style='background:#0F2B5B;'>
                <th style='padding:8px;text-align:left;font-size:10px;color:#8899bb;text-transform:uppercase;'>OP / Cliente</th>
                <th style='padding:8px;text-align:left;font-size:10px;color:#8899bb;text-transform:uppercase;'>Descripción</th>
                <th style='padding:8px;text-align:center;font-size:10px;color:#8899bb;text-transform:uppercase;'>📅 Ingreso</th>
                <th style='padding:8px;text-align:center;font-size:10px;color:#8899bb;text-transform:uppercase;'>⏱ Días</th>
                <th style='padding:8px;text-align:center;font-size:10px;color:#8899bb;text-transform:uppercase;'>💵 Valor</th>
                <th style='padding:8px;text-align:center;font-size:10px;color:#8899bb;text-transform:uppercase;'>📋 HE</th>
              </tr>
            </thead>
            <tbody>{filas_hist}</tbody>
          </table>
        </div>"""

    historico_completo = f"""
    <div style='margin-top:32px;padding-top:20px;border-top:2px solid #1e3a6e;'>
      <div style='font-size:13px;color:#E8A020;font-weight:900;text-transform:uppercase;letter-spacing:2px;margin-bottom:14px;'>📊 Histórico Completo de OPs</div>
      {_bloque_grupo('🟢 Activas', activos, '#2EAA4A')}
      {_bloque_grupo('🔵 Completadas', completados, '#1a56db')}
      {_bloque_grupo('🔴 Canceladas', cancelados, '#e53e3e')}
    </div>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Open Sans',Arial,sans-serif;">
<div style="max-width:900px;margin:0 auto;padding:16px;">
  <div style="background:#0F2B5B;padding:18px 28px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;">
    <div>
      <div style="font-size:20px;font-weight:900;color:#fff;letter-spacing:1px;">⚡ ENERGY</div>
      <div style="font-size:11px;color:#93c5fd;margin-top:2px;text-transform:uppercase;letter-spacing:1px;">Reporte Diario — Órdenes de Pedido</div>
    </div>
    <div style="font-size:12px;color:#93c5fd;text-align:right;">{date_str}</div>
  </div>
  <div style="height:3px;background:linear-gradient(90deg,#E8A020,transparent);"></div>
  <div style="background:#071525;padding:20px;border-radius:0 0 0 0;">
    {resumen}
    {resumen_etapas}
    {seccion_he}
    {cuerpo}
    {historico_completo}
  </div>
  <div style="background:#071525;padding:14px;border-radius:0 0 12px 12px;text-align:center;margin-top:0;border-top:1px solid #1e3a6e;">
    <p style="margin:3px 0;font-size:11px;color:#8899bb;">⚡ Generado por <strong style="color:#E8A020;">ENERGY — Asistente Administrativo</strong></p>
    <p style="margin:3px 0;font-size:11px;color:#8899bb;">E&amp;G Energy Group · Reporte automático L–V 5:00 PM Colombia</p>
  </div>
</div>
</body>
</html>"""


# ─── 4. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    data = urllib.parse.urlencode({
        'grant_type':    'client_credentials',
        'client_id':     client_id,
        'client_secret': client_secret,
        'scope':         'https://graph.microsoft.com/.default'
    }).encode()
    req = urllib.request.Request(url, data=data, method='POST')
    print("🔑 Obteniendo token Microsoft Graph...")
    try:
        with urllib.request.urlopen(req) as resp:
            print("✅ Token obtenido correctamente")
            return json.loads(resp.read())['access_token']
    except urllib.error.HTTPError as e:
        print(f"ERROR obteniendo token: {e.code} — {e.read().decode()}")
        sys.exit(1)


# ─── 5. ENVIAR CORREO VIA GRAPH ────────────────────────────────────────────────

def send_email(token: str, sender: str, recipients: list, subject: str, html_body: str):
    payload = json.dumps({
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': r.strip()}} for r in recipients]
        },
        'saveToSentItems': True
    }).encode('utf-8')
    url = f'https://graph.microsoft.com/v1.0/users/{sender}/sendMail'
    req = urllib.request.Request(url, data=payload, method='POST',
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req) as resp:
            print(f"✅ Correo enviado (HTTP {resp.status})")
    except urllib.error.HTTPError as e:
        print(f"ERROR enviando correo: {e.code} — {e.read().decode()}")
        sys.exit(1)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    cfg = load_config()

    tenant_id     = os.environ.get('MS_TENANT_ID',     cfg.get('ms_tenant_id',     '')).strip()
    client_id     = os.environ.get('MS_CLIENT_ID',     cfg.get('ms_client_id',     '')).strip()
    client_secret = os.environ.get('MS_CLIENT_SECRET', cfg.get('ms_client_secret', '')).strip()
    sender_email  = os.environ.get('SENDER_EMAIL',     cfg.get('sender_email',     '')).strip()
    recipients    = [r.strip() for r in os.environ.get(
                        'RECIPIENT_EMAILS', cfg.get('recipient_emails', '')).split(',')]
    extra_str     = os.environ.get('EXTRA_RECIPIENTS', cfg.get('extra_recipients', '')).strip()
    if extra_str:
        recipients += [r.strip() for r in extra_str.split(',') if r.strip()]
    recipients    = list(dict.fromkeys(r for r in recipients if r))  # eliminar duplicados vacíos
    gh_token      = os.environ.get('GH_TOKEN',         cfg.get('github_token',     '')).strip()

    print(f"🔍 Sender        : {sender_email}")
    print(f"🔍 Destinatarios : {', '.join(recipients)}")

    # Descargar órdenes desde GitHub
    ordenes = load_ordenes_from_github(gh_token)

    now      = datetime.now()
    day      = DAYS_ES.get(now.strftime('%A'), now.strftime('%A'))
    month    = MONTHS_ES.get(now.strftime('%B'), now.strftime('%B'))
    date_str = f"{day} {now.day} de {month} de {now.year}"
    subject  = f"📋 Informe OP Activas E&G — {date_str}"

    print(f"📝 Generando informe para {len(ordenes)} órdenes...")
    email_html = build_report_html(ordenes, date_str)

    token = get_access_token(tenant_id, client_id, client_secret)
    print(f"📧 Enviando a: {', '.join(recipients)}")
    send_email(token, sender_email, recipients, subject, email_html)
    print("🎉 Informe OP enviado exitosamente.")
