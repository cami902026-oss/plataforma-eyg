"""
ENERGY — Reporte Diario de Inventario
======================================
Lee el inventario desde Index.html, calcula estadísticas
y envía un correo HTML via Microsoft Graph API.

Ejecutado automáticamente por GitHub Actions
Lunes a Viernes a las 5:00 PM Colombia (UTC-5 = 22:00 UTC)
"""

import json
import re
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime


# ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

DAYS_ES    = {'Monday':'Lunes','Tuesday':'Martes','Wednesday':'Miércoles',
               'Thursday':'Jueves','Friday':'Viernes','Saturday':'Sábado','Sunday':'Domingo'}
MONTHS_ES  = {'January':'enero','February':'febrero','March':'marzo','April':'abril',
               'May':'mayo','June':'junio','July':'julio','August':'agosto',
               'September':'septiembre','October':'octubre','November':'noviembre','December':'diciembre'}


# ─── 1. EXTRAE INVENTARIO DESDE Buscador_Inventario_2026.html ────────────────

def extract_rot_from_html(html_path: str) -> list:
    """Extrae el array ROT_DATA del archivo Buscador_Inventario_2026.html"""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    match = re.search(r'const ROT_DATA\s*=\s*(\[[\s\S]*?\]);', content)
    if not match:
        return []
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError:
        return []

def extract_inv_from_html(html_path: str) -> list:
    """Extrae el array STOCK_DATA del archivo Buscador_Inventario_2026.html"""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    match = re.search(r'const STOCK_DATA\s*=\s*(\[[\s\S]*?\]);', content)
    if not match:
        print("ERROR: No se encontró 'const STOCK_DATA' en el HTML")
        return []
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError as e:
        print(f"ERROR parseando JSON del inventario: {e}")
        return []


# ─── 2. CALCULA ESTADÍSTICAS ──────────────────────────────────────────────────

def calculate_stats(inv: list) -> dict:
    total     = len(inv)
    con_stock = sum(1 for p in inv if _stock(p) > 0)
    agotados  = sum(1 for p in inv if _stock(p) == 0)
    entradas  = sum(_num(p, 'ENTRADAS') for p in inv)
    salidas   = sum(_num(p, 'SALIDAS')  for p in inv)
    rotacion  = round(salidas / total, 2) if total else 0
    disponib  = round((con_stock / total) * 100, 1) if total else 0
    return dict(total=total, con_stock=con_stock, agotados=agotados,
                entradas=entradas, salidas=salidas, rotacion=rotacion,
                disponib=disponib)

def _stock(p): return int(p.get('STOCK ACTUAL') or 0)
def _num(p, k): return int(p.get(k) or 0)


# ─── 3. GENERA EL EMAIL HTML ──────────────────────────────────────────────────

def _stock_badge(stock):
    if stock == 0:
        return "<span style='background:#fce8e6;color:#c0392b;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>Agotado</span>"
    elif stock <= 3:
        return "<span style='background:#fff8e1;color:#b7770d;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>&#9888; " + str(stock) + "</span>"
    else:
        return "<span style='background:#e6f4ea;color:#1e7e34;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;'>&#10003; " + str(stock) + "</span>"


def _cat_label(cat: str) -> str:
    m = {'mecanico': 'Material Mecánico', 'electrico': 'Material Eléctrico',
         'instrumentacion': 'Instrumentación', 'instrumentación': 'Instrumentación'}
    return m.get(str(cat).lower().strip(), 'Otros')

def _cat_color(cat: str) -> tuple:
    """Retorna (fondo_header, color_acento) según categoría."""
    m = {
        'mecanico':       ('#0F2B5B', '#1A3A8F'),
        'electrico':      ('#7B3F00', '#C0641E'),
        'instrumentacion':('#1B5E20', '#2E7D32'),
        'instrumentación':('#1B5E20', '#2E7D32'),
    }
    return m.get(str(cat).lower().strip(), ('#37474F', '#546E7A'))

def _auto_classify(p: dict) -> tuple:
    """Auto-clasifica un producto en (categoria, familia) por palabras clave.
    Usa los campos CATEGORIA y FAMILIA del Excel si ya existen.
    """
    # Respeta los valores del Excel si están presentes
    cat_excel = str(p.get('CATEGORIA', '') or '').lower().strip()
    fam_excel = str(p.get('FAMILIA',   '') or '').strip()

    desc = str(p.get('DESCRIPCION', '') or '').upper()

    # ── Categoría ─────────────────────────────────────────────────────────────
    if cat_excel in ('mecanico','mecánico','electrico','eléctrico',
                     'instrumentacion','instrumentación'):
        cat = cat_excel.replace('á','a').replace('é','e').replace('ó','o')
    else:
        ELEC_KW  = ['CABLE','BREAKER','INTERRUPTOR','TRANSFORMADOR','MOTOR',
                    'BOBINA','CONTACTOR','RELE','RELÉ','FUSIBLE','TERMINAL',
                    'SWITCH','VARIADOR','CONDUCTOR','TABLERO',
                    'LAMPARA','LÁMPARA','ENCHUFE','UPS',
                    'LIQUID TIGHT','LIQUIDTIGHT','CORAZA','EMT','CONDUIT',
                    'AISLADOR','CAJA DE PASO','CAJA FUNDIDA','CAJA ELECTRICA',
                    'TOMA ELECT','TOMA CORR','TOMACORRIENTE',
                    'CONECTOR CURVO','CURVA EMT','CURVA CONDUIT']
        INSTR_KW = ['SENSOR','TRANSMISOR','MEDIDOR','PRESOSTATO','TERMOSTATO',
                    'MANOMETRO','MANÓMETRO','CAUDALIMETRO','CONTROLADOR',
                    'PLC','HMI','INDICADOR','ALARMA','TRANSDUCTOR','DETECTOR',
                    'TERMOCUPLA','TERMOPAR','ROTAMETRO']
        MECA_KW  = ['CODO','BUSHING','PLATINA','VALVULA','VÁLVULA','NIPLE',
                    'NIPPLE','TEE','UNION','UNIÓN','BRIDA','JUNTA','EMPAQUE',
                    'TORNILLO','PERNO','TUERCA','BUJE','RODAMIENTO','MANGUERA',
                    'ABRAZADERA','ORIFICIO','REDUCCION','REDUCCIÓN','TUBO',
                    'TUBERIA','TUBERÍA','SELLO','RESORTE','EJE','ACOPLAMIENTO',
                    'CADENA','SPROCKET','POLEA','CORREA','FILTRO',
                    'ADAPTADOR','TAPÓN','TAPON','PLUG','CAP ','SOCKET NPT',
                    'ESPIROMETALICO','ESPIROMETÁLICO','ESPIROME',
                    'ESPARRAGO','ESPÁRRAGO','FLANCHE','FLANGE',
                    'THREADOLET','SOCKOLET','WELDOLET','OLET',
                    'UNIVERSAL','COPA AC','COPA DE AC']
        if   any(k in desc for k in ELEC_KW):  cat = 'electrico'
        elif any(k in desc for k in INSTR_KW): cat = 'instrumentacion'
        elif any(k in desc for k in MECA_KW):  cat = 'mecanico'
        else:                                   cat = 'otros'

    # ── Familia ───────────────────────────────────────────────────────────────
    if fam_excel:
        familia = fam_excel
    else:
        FAM_RULES = [
            (['CODO'],                              'Codos'),
            (['BUSHING','REDUCCION','REDUCCIÓN'],   'Reducciones y Bushings'),
            (['PLATINA'],                           'Platinas'),
            (['VALVULA','VÁLVULA','VALVE'],         'Válvulas'),
            (['NIPLE','NIPPLE'],                    'Niples'),
            (['TEE'],                               'Tees'),
            (['UNION','UNIÓN'],                     'Uniones'),
            (['UNIVERSAL'],                         'Universales'),
            (['BRIDA','FLANCHE','FLANGE'],          'Bridas y Flanches'),
            (['ESPIROMETALICO','ESPIROMETÁLICO','ESPIROME'], 'Juntas Espirometálicas'),
            (['JUNTA','EMPAQUE','SELLO'],           'Juntas y Empaques'),
            (['TORNILLO','PERNO','TUERCA','ARANDELA','ESPARRAGO','ESPÁRRAGO'], 'Tornillería'),
            (['THREADOLET','SOCKOLET','WELDOLET','OLET'], 'Olets'),
            (['COPA AC','COPA DE AC'],              'Copas AC'),
            (['FILTRO'],                            'Filtros'),
            (['RODAMIENTO','BUJE'],                 'Rodamientos y Bujes'),
            (['MANGUERA','TUBERIA','TUBERÍA','TUBO'], 'Tuberías y Mangueras'),
            (['CABLE','CONDUCTOR'],                 'Cables y Conductores'),
            (['LIQUID TIGHT','LIQUIDTIGHT','CORAZA','CONDUIT','EMT'],
                                                   'Conduit y Accesorios'),
            (['CURVA EMT','CURVA CONDUIT'],         'Curvas Conduit'),
            (['CAJA DE PASO','CAJA FUNDIDA','CAJA ELECTRICA'], 'Cajas Eléctricas'),
            (['AISLADOR'],                          'Aisladores'),
            (['TOMA','TOMACORRIENTE','ENCHUFE'],    'Tomas y Enchufes'),
            (['SENSOR','TRANSMISOR'],               'Sensores y Transmisores'),
            (['MANOMETRO','MANÓMETRO','PRESOSTATO','TERMOSTATO'], 'Instrumentos de Medición'),
            (['MOTOR','BOBINA','CONTACTOR'],        'Motores y Accionamiento'),
            (['BREAKER','FUSIBLE','INTERRUPTOR'],   'Protecciones Eléctricas'),
        ]
        familia = 'Otros'
        for keywords, fam_name in FAM_RULES:
            if any(k in desc for k in keywords):
                familia = fam_name
                break

    return cat, familia


def _familia_id(cat: str, familia: str) -> str:
    """Genera un id HTML seguro para anclas de familia."""
    import re
    safe = re.sub(r'[^a-zA-Z0-9]', '-', (cat + '-' + familia).lower())
    return 'fam-' + safe


def _build_familia_block(familia: str, productos: list, accent: str, cat: str = '') -> str:
    # Ordenar: primero alfabético por descripción (agrupa tees, codos, etc.)
    # Los agotados van al final dentro de cada grupo
    inv_sorted = sorted(productos, key=lambda p: (
        _stock(p) == 0,
        str(p.get('DESCRIPCION', '')).upper()
    ))
    con_stock = sum(1 for p in productos if _stock(p) > 0)
    agotados  = sum(1 for p in productos if _stock(p) == 0)

    rows = ''.join(
        "<tr>"
        "<td style='font-family:monospace;font-size:11px;color:#555;'>" + str(p.get('CODIGO PRODUCTO','')) + "</td>"
        "<td><b style='font-size:12px;'>" + str(p.get('DESCRIPCION','')) + "</b></td>"
        "<td style='color:#666;'>" + str(p.get('MARCA','-')) + "</td>"
        "<td style='color:#666;'>" + str(p.get('UBICACIÓN') or p.get('UBICACION','-')) + "</td>"
        "<td style='text-align:center;'>" + _stock_badge(_stock(p)) + "</td>"
        "<td style='text-align:center;color:#555;font-size:11px;'>+" + str(_num(p,'ENTRADAS')) + " / -" + str(_num(p,'SALIDAS')) + "</td>"
        "</tr>"
        for p in inv_sorted
    )

    badge_parts = str(len(productos)) + " items"
    if agotados:
        badge_parts += " &nbsp;&#9888; " + str(agotados) + " agotados"

    fam_id = _familia_id(cat, familia)

    return (
        "<tr id='" + fam_id + "' style='background:" + accent + "18;'>"
        "<td colspan='6' style='padding:6px 12px;font-size:11px;font-weight:700;"
        "color:" + accent + ";border-left:3px solid " + accent + ";letter-spacing:.5px;'>"
        "&#128281; " + familia +
        " <span style='font-weight:400;color:#888;font-size:10px;'>— " + badge_parts + "</span>"
        " <a href='#idx-top' style='float:right;font-size:9px;color:#aaa;text-decoration:none;'>&#8593; inicio</a>"
        "</td></tr>"
        + rows
    )

def _build_category_section(cat: str, productos: list) -> str:
    label = _cat_label(cat)
    hdr_color, accent = _cat_color(cat)
    con_stock = sum(1 for p in productos if _stock(p) > 0)
    agotados  = sum(1 for p in productos if _stock(p) == 0)

    # Agrupar por FAMILIA dentro de la categoría
    familias = {}
    for p in productos:
        fam = str(p.get('FAMILIA','') or 'Sin familia').strip()
        familias.setdefault(fam, []).append(p)

    # Bloques por familia ordenados por nombre
    familia_blocks = ''
    for fam in sorted(familias.keys()):
        familia_blocks += _build_familia_block(fam, familias[fam], accent, cat)

    cat_id = 'cat-' + cat.replace('ó','o').replace('ú','u').replace('é','e').replace('á','a').replace('í','i')

    return (
        "<div id='" + cat_id + "' style='margin:18px 0 8px;'>"
        "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'><tr>"
        "<td bgcolor='" + hdr_color + "' style='background:" + hdr_color + ";padding:10px 14px;'>"
        "<span style='font-size:13px;font-weight:700;letter-spacing:1px;color:#ffffff;'>" + label + "</span>"
        "</td>"
        "<td bgcolor='" + hdr_color + "' style='background:" + hdr_color + ";padding:10px 14px;text-align:right;'>"
        "<span style='font-size:11px;color:#ffffff;'>"
        + str(len(productos)) + " productos &nbsp;|&nbsp; "
        + str(con_stock) + " con stock &nbsp;|&nbsp; "
        + str(agotados) + " agotados</span>"
        "</td>"
        "</tr></table>"
        "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;font-size:12px;'>"
        "<thead><tr>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>C&oacute;digo</th>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Descripci&oacute;n</th>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Marca</th>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Ubic.</th>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Stock</th>"
        "<th bgcolor='" + accent + "' style='background:" + accent + ";color:#ffffff;padding:7px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Ent/Sal</th>"
        "</tr></thead>"
        "<tbody>" + familia_blocks + "</tbody>"
        "</table></div>"
    )

def _build_rotation_section(rot: list) -> str:
    if not rot:
        return ''
    from collections import Counter
    cats = Counter(r['CATEGORIA'] for r in rot)

    CAT_CONFIG = {
        'ESTRELLA':      ('#1B5E20', '#2E7D32', '⭐', 'Alta rotación — productos clave'),
        'NORMAL':        ('#0F2B5B', '#1A3A8F', '✅', 'Rotación normal'),
        'LENTO':         ('#7B3F00', '#C0641E', '🐢', 'Rotación lenta — monitorear'),
        'INACTIVO':      ('#7B0000', '#c0392b', '⚠️', 'Sin movimiento >60 días — evaluar descuento'),
        'NUNCA VENDIDO': ('#37474F', '#546E7A', '📦', 'Sin historial de ventas — revisar relevancia'),
        'SIN STOCK':     ('#4A148C', '#7B1FA2', '🔴', 'Sin stock disponible'),
    }

    # Resumen por categoría
    summary_cells = ''
    for cat, (hdr, acc, icon, desc) in CAT_CONFIG.items():
        count = cats.get(cat, 0)
        if count == 0:
            continue
        summary_cells += (
            "<td style='padding:8px;text-align:center;background:" + hdr + "18;"
            "border-left:3px solid " + acc + ";border-radius:6px;'>"
            "<div style='font-size:22px;font-weight:900;color:" + acc + ";'>" + str(count) + "</div>"
            "<div style='font-size:9px;color:#555;text-transform:uppercase;letter-spacing:.5px;margin-top:2px;'>" + cat + "</div>"
            "</td>"
        )

    # Detalle por categoría (solo ESTRELLA, INACTIVO, LENTO)
    detail_html = ''
    for cat in ['ESTRELLA', 'INACTIVO', 'LENTO']:
        items = [r for r in rot if r['CATEGORIA'] == cat]
        if not items:
            continue
        hdr, acc, icon, desc = CAT_CONFIG[cat]
        rows = ''.join(
            "<tr>"
            "<td style='font-family:monospace;font-size:11px;color:#555;padding:6px 10px;border-bottom:1px solid #eef1f8;'>" + str(r.get('CODIGO','')) + "</td>"
            "<td style='padding:6px 10px;border-bottom:1px solid #eef1f8;font-size:12px;'>" + str(r.get('DESCRIPCION','')) + "</td>"
            "<td style='padding:6px 10px;border-bottom:1px solid #eef1f8;text-align:center;font-size:12px;'>" + str(r.get('STOCK',0)) + "</td>"
            "<td style='padding:6px 10px;border-bottom:1px solid #eef1f8;text-align:center;font-size:12px;'>" + str(r.get('VENDIDO',0)) + "</td>"
            "<td style='padding:6px 10px;border-bottom:1px solid #eef1f8;text-align:center;font-size:11px;color:#888;'>" + str(r.get('ULT_VENTA','—')) + "</td>"
            "<td style='padding:6px 10px;border-bottom:1px solid #eef1f8;font-size:10px;color:#666;'>" + str(r.get('DIAGNOSTICO','')) + "</td>"
            "</tr>"
            for r in items
        )
        detail_html += (
            "<div style='margin:10px 0 4px;padding:7px 12px;background:" + hdr + ";color:#fff;"
            "border-radius:6px;font-size:11px;font-weight:700;letter-spacing:.5px;'>"
            + icon + " " + cat + " — " + desc + " (" + str(len(items)) + " productos)</div>"
            "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;font-size:12px;margin-bottom:8px;'>"
            "<thead><tr>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Código</th>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Descripción</th>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Stock</th>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Vendido</th>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:center;font-size:10px;text-transform:uppercase;'>Últ. Venta</th>"
            "<th style='background:" + acc + ";color:#fff;padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;'>Diagnóstico</th>"
            "</tr></thead>"
            "<tbody>" + rows + "</tbody></table>"
        )

    return (
        "<div style='margin-top:24px;'>"
        "<div style='font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;"
        "color:#0F2B5B;border-bottom:2px solid #E8A020;padding-bottom:6px;margin-bottom:14px;'>"
        "&#128200; An&aacute;lisis de Rotaci&oacute;n</div>"
        "<table width='100%' cellpadding='6' cellspacing='6' style='margin-bottom:16px;'>"
        "<tr>" + summary_cells + "</tr></table>"
        + detail_html +
        "</div>"
    )

def _build_index(grupos: dict) -> str:
    """Genera un índice de navegación rápida por categoría y familia."""
    CAT_ORDER   = ['mecanico', 'electrico', 'instrumentacion', 'otros']
    CAT_ICONS   = {'mecanico': '🔧', 'electrico': '⚡', 'instrumentacion': '🎛️', 'otros': '📦'}
    CAT_COLORS  = {'mecanico': '#0F2B5B', 'electrico': '#7B3F00',
                   'instrumentacion': '#1B5E20', 'otros': '#37474F'}

    cat_blocks = ''
    for cat in CAT_ORDER:
        if cat not in grupos or not grupos[cat]:
            continue
        color = CAT_COLORS.get(cat, '#333')
        icon  = CAT_ICONS.get(cat, '📦')
        label = _cat_label(cat)
        cat_id = 'cat-' + cat

        # familias dentro de esta categoría
        familias = {}
        for p in grupos[cat]:
            fam = str(p.get('FAMILIA','') or 'Sin familia').strip()
            familias.setdefault(fam, []).append(p)

        fam_links = ''
        for fam in sorted(familias.keys()):
            fam_id  = _familia_id(cat, fam)
            count   = len(familias[fam])
            agot    = sum(1 for p in familias[fam] if _stock(p) == 0)
            agot_lbl = " <span style='color:#c0392b;font-size:9px;'>&#9888;" + str(agot) + "</span>" if agot else ""
            fam_links += (
                "<a href='#" + fam_id + "' style='display:inline-block;margin:2px 4px 2px 0;"
                "padding:3px 9px;background:#f0f4f8;border:1px solid #d1d9e6;"
                "border-radius:12px;font-size:11px;color:#334155;text-decoration:none;'>"
                + fam + " <span style='color:#888;font-size:10px;'>(" + str(count) + ")</span>" + agot_lbl +
                "</a>"
            )

        con_stock = sum(1 for p in grupos[cat] if _stock(p) > 0)
        agotados  = sum(1 for p in grupos[cat] if _stock(p) == 0)

        cat_blocks += (
            "<tr><td style='padding:10px 14px 8px;border-bottom:1px solid #e5eaf2;vertical-align:top;'>"
            "<a href='#" + cat_id + "' style='text-decoration:none;'>"
            "<span style='display:inline-block;background:" + color + ";color:#fff;"
            "padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700;margin-bottom:7px;'>"
            + icon + " " + label +
            " <span style='font-weight:400;font-size:10px;opacity:.85;'>— "
            + str(len(grupos[cat])) + " prod / "
            + str(con_stock) + " en stock"
            + (" / <span style='color:#fca5a5;'>" + str(agotados) + " agotados</span>" if agotados else "")
            + "</span></span></a><br>"
            + fam_links +
            "</td></tr>"
        )

    return (
        "<div id='idx-top' style='margin:0 0 18px;background:#f8faff;"
        "border:1.5px solid #d1d9e6;border-radius:10px;overflow:hidden;'>"
        "<div style='background:#0F2B5B;padding:9px 14px;'>"
        "<span style='font-size:11px;font-weight:700;color:#fff;letter-spacing:.5px;'>"
        "&#128269; &nbsp;ÍNDICE RÁPIDO — haz clic para ir a la sección</span>"
        "</div>"
        "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>"
        + cat_blocks +
        "</table></div>"
    )


def generate_html(inv: list, stats: dict, date_str: str, rot: list = None) -> str:
    # ── Auto-clasificar productos que no tengan CATEGORIA/FAMILIA en el Excel ──
    for p in inv:
        cat_excel = str(p.get('CATEGORIA', '') or '').lower().strip()
        fam_excel = str(p.get('FAMILIA',   '') or '').strip()
        if not cat_excel or not fam_excel:
            auto_cat, auto_fam = _auto_classify(p)
            if not cat_excel:
                p['CATEGORIA'] = auto_cat
            if not fam_excel:
                p['FAMILIA'] = auto_fam

    # Agrupar por categoría
    CAT_ORDER = ['mecanico', 'electrico', 'instrumentacion', 'otros']
    grupos = {}
    for p in inv:
        cat = str(p.get('CATEGORIA', '') or '').lower().strip()
        if cat not in ('mecanico','electrico','instrumentacion','instrumentación'):
            cat = 'otros'
        grupos.setdefault(cat, []).append(p)

    # Índice de navegación
    index_block = _build_index(grupos)

    # Secciones por categoría
    cat_sections = ''
    for cat in CAT_ORDER:
        if cat in grupos and grupos[cat]:
            cat_sections += _build_category_section(cat, grupos[cat])

    # Filas de productos agotados (todas las categorías)
    agotados_rows = ''.join(
        "<tr>"
        "<td style='font-family:monospace;font-size:11px;'>" + str(p.get('CODIGO PRODUCTO','')) + "</td>"
        "<td>" + str(p.get('DESCRIPCION','')) + "</td>"
        "<td>" + str(p.get('MARCA','-')) + "</td>"
        "<td>" + str(p.get('UBICACIÓN') or p.get('UBICACION','-')) + "</td>"
        "<td style='color:#888;font-size:11px;'>" + _cat_label(p.get('CATEGORIA','')) + "</td>"
        "</tr>"
        for p in inv if _stock(p) == 0
    )
    agotados_count = stats['agotados']

    # Pre-calcular condicionales (no pueden ir dentro de {} en f-strings — Python < 3.12)
    pct_agot   = stats['agotados'] / max(stats['total'], 1)
    health_msg = '&#9989; Inventario saludable' if pct_agot < 0.1 else '&#9888; +10% agotado &mdash; revisar reabastecimiento'

    alert_box = ''
    if agotados_count > 0:
        alert_box = ('<div class="alert-box">&#9888;&#65039; <strong>' + str(agotados_count) +
                     ' productos agotados</strong> &mdash; ver secci&oacute;n al final del reporte.</div>')

    if agotados_count > 0:
        agotados_section = (
            '<div class="sec-title" style="color:#c0392b;">&#128308; Productos Agotados (' + str(agotados_count) + ')</div>'
            '<table><thead><tr><th>C&oacute;digo</th><th>Descripci&oacute;n</th><th>Marca</th><th>Ubicaci&oacute;n</th><th>Categor&iacute;a</th></tr></thead>'
            '<tbody>' + agotados_rows + '</tbody></table>'
        )
    else:
        agotados_section = '<div class="meta">&#9989; No hay productos agotados en este momento.</div>'

    rotation_section = _build_rotation_section(rot) if rot else ''

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4ff;margin:0;padding:16px;}}
  .wrap{{max-width:900px;margin:0 auto;box-shadow:0 8px 32px rgba(0,0,0,.15);border-radius:14px;overflow:hidden;}}
  .hdr{{background:#0F2B5B;padding:28px;text-align:center;}}
  .hdr h1{{margin:0;font-size:24px;color:#fff;letter-spacing:3px;font-weight:900;}}
  .hdr p{{margin:6px 0 0;color:#8899bb;font-size:13px;}}
  .gold{{height:3px;background:linear-gradient(90deg,#E8A020,transparent);}}
  .body{{background:#fff;padding:24px;}}
  .sec-title{{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;
              color:#0F2B5B;border-bottom:2px solid #E8A020;padding-bottom:6px;margin:22px 0 14px;}}
  .stats{{display:table;width:100%;border-spacing:10px;}}
  .st{{display:table-cell;text-align:center;padding:16px 8px;border-radius:10px;}}
  .st.blue{{background:#e8f0fe;border-left:4px solid #1A3A8F;}}
  .st.green{{background:#e6f4ea;border-left:4px solid #2EAA4A;}}
  .st.red{{background:#fce8e6;border-left:4px solid #e53e3e;}}
  .st .num{{font-size:32px;font-weight:900;color:#0F2B5B;line-height:1;}}
  .st .lbl{{font-size:10px;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;}}
  .kd{{display:table;width:100%;border-spacing:10px;margin:10px 0;}}
  .kd-box{{display:table-cell;text-align:center;padding:14px;border-radius:10px;}}
  .kd-box.ent{{background:#e6f4ea;border:1px solid #2EAA4A;}}
  .kd-box.sal{{background:#fff8e1;border:1px solid #E8A020;}}
  .kd-box .knum{{font-size:26px;font-weight:900;}}
  .kd-box.ent .knum{{color:#2EAA4A;}}
  .kd-box.sal .knum{{color:#E8A020;}}
  .kd-box .klbl{{font-size:10px;color:#666;margin-top:4px;}}
  .meta{{background:#f8f9ff;border-radius:8px;padding:10px 14px;font-size:12px;color:#555;margin:10px 0;}}
  .meta span{{color:#0F2B5B;font-weight:700;}}
  .alert-box{{background:#fff3cd;border-left:4px solid #e8a020;border-radius:6px;
              padding:10px 14px;font-size:12px;color:#7d5a00;margin:10px 0;}}
  .search-wrap{{margin-bottom:10px;}}
  .search-bar{{width:100%;box-sizing:border-box;padding:9px 14px;font-size:13px;
               border:2px solid #c5d0e8;border-radius:8px;outline:none;}}
  .search-bar:focus{{border-color:#0F2B5B;}}
  .search-tip{{font-size:11px;color:#888;margin:0 0 8px;}}
  table{{width:100%;border-collapse:collapse;font-size:12px;}}
  th{{background:#0F2B5B;color:#fff;padding:8px 10px;text-align:left;font-size:10px;
      text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;}}
  td{{padding:7px 10px;border-bottom:1px solid #eef1f8;vertical-align:middle;}}
  tr:nth-child(even) td{{background:#f8f9ff;}}
  tr.hidden{{display:none;}}
  #sinResultados{{display:none;text-align:center;color:#999;padding:20px;font-size:13px;}}
  .footer{{background:#071525;padding:16px;text-align:center;}}
  .footer p{{margin:3px 0;font-size:11px;color:#8899bb;}}
  .footer strong{{color:#E8A020;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <h1>&#9889; ENERGY</h1>
    <p>Reporte Diario de Inventario &mdash; {date_str}</p>
  </div>
  <div class="gold"></div>
  <div style="background:#0a1a30;padding:12px 20px;text-align:center;border-bottom:1px solid #1e3a6e;">
    <a href="https://cami902026-oss.github.io/plataforma-eyg/Buscador_Inventario_2026.html" target="_blank"
       style="display:inline-block;background:#2EAA4A;color:#fff;padding:10px 28px;border-radius:8px;font-weight:700;font-size:14px;text-decoration:none;letter-spacing:0.5px;">
      &#128269; Ver inventario interactivo con b&uacute;squeda
    </a>
    <div style="font-size:11px;color:#8899bb;margin-top:6px;">Abre el buscador en tu navegador para filtrar y buscar productos</div>
  </div>
  <div class="body">

    <div class="sec-title">&#128230; Resumen General</div>
    <table width="100%" cellpadding="8" cellspacing="6" style="margin:10px 0;">
      <tr>
        <td width="33%" bgcolor="#e8f0fe" style="background:#e8f0fe;border-left:4px solid #1A3A8F;padding:16px 8px;text-align:center;">
          <div style="font-size:32px;font-weight:900;color:#0F2B5B;line-height:1;">{stats['total']}</div>
          <div style="font-size:10px;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">Total Productos</div>
        </td>
        <td width="33%" bgcolor="#e6f4ea" style="background:#e6f4ea;border-left:4px solid #2EAA4A;padding:16px 8px;text-align:center;">
          <div style="font-size:32px;font-weight:900;color:#1B5E20;line-height:1;">{stats['con_stock']}</div>
          <div style="font-size:10px;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">Con Stock</div>
        </td>
        <td width="33%" bgcolor="#fce8e6" style="background:#fce8e6;border-left:4px solid #e53e3e;padding:16px 8px;text-align:center;">
          <div style="font-size:32px;font-weight:900;color:#c0392b;line-height:1;">{stats['agotados']}</div>
          <div style="font-size:10px;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">Agotados</div>
        </td>
      </tr>
    </table>

    <div class="sec-title">&#128202; Movimientos Acumulados</div>
    <table width="100%" cellpadding="8" cellspacing="6" style="margin:10px 0;">
      <tr>
        <td width="50%" bgcolor="#e6f4ea" style="background:#e6f4ea;border-left:4px solid #2EAA4A;padding:14px;text-align:center;">
          <div style="font-size:26px;font-weight:900;color:#2EAA4A;">+{stats['entradas']}</div>
          <div style="font-size:10px;color:#666;margin-top:4px;">Total Entradas</div>
        </td>
        <td width="50%" bgcolor="#fff8e1" style="background:#fff8e1;border-left:4px solid #E8A020;padding:14px;text-align:center;">
          <div style="font-size:26px;font-weight:900;color:#E8A020;">-{stats['salidas']}</div>
          <div style="font-size:10px;color:#666;margin-top:4px;">Total Salidas</div>
        </td>
      </tr>
    </table>
    <div class="meta">
      &#128260; Rotaci&oacute;n: <span>{stats['rotacion']}</span> sal/prod &nbsp;|&nbsp;
      &#127919; Disponibilidad: <span>{stats['disponib']}%</span> &nbsp;|&nbsp;
      {health_msg}
    </div>

    {alert_box}

    <div class="sec-title">&#128230; Inventario por Categor&iacute;a &mdash; {stats['total']} productos</div>
    {index_block}
    <div class="search-wrap">
      <div style="display:flex;align-items:center;background:white;border:2px solid #c5d0e8;border-radius:10px;padding:8px 14px;margin-bottom:6px;" id="inv-sbox">
        <span style="font-size:16px;margin-right:8px;color:#94a3b8;">&#128269;</span>
        <input id="inv-report-search" class="search-bar" type="text" placeholder="Buscar por c&oacute;digo, descripci&oacute;n, marca, ubicaci&oacute;n..." autocomplete="off" style="border:none;outline:none;font-size:14px;width:100%;color:#1e293b;background:transparent;padding:0;" oninput="filtrarInventario(this.value)">
        <button onclick="document.getElementById('inv-report-search').value='';filtrarInventario('');" style="background:none;border:none;cursor:pointer;color:#94a3b8;font-size:18px;line-height:1;" title="Limpiar">&#215;</button>
      </div>
      <p class="search-tip" id="inv-search-tip">Escribe para buscar entre los {stats['total']} productos del reporte.</p>
    </div>
    {cat_sections}

    {agotados_section}

    {rotation_section}

  </div>
  <div class="footer">
    <p>&#9889; Generado por <strong>ENERGY &mdash; Asistente Administrativo</strong></p>
    <p>E&amp;G Energy Group &middot; Reporte autom&aacute;tico L&ndash;V 5:00 PM Colombia</p>
  </div>
</div>
<script>
function filtrarInventario(term) {{
  const q = (term || '').trim().toLowerCase();
  const tip = document.getElementById('inv-search-tip');
  const wrap = document.querySelector('.body');
  const allRows = [];
  wrap.querySelectorAll('table').forEach(function(tbl) {{
    tbl.querySelectorAll('tbody tr').forEach(function(tr) {{ allRows.push(tr); }});
  }});
  if (!q) {{
    allRows.forEach(function(tr) {{ tr.style.display = ''; }});
    tip.textContent = 'Escribe para buscar entre los {stats['total']} productos del reporte.';
    tip.style.color = '#888';
    return;
  }}
  let visibles = 0;
  allRows.forEach(function(tr) {{
    if (tr.textContent.toLowerCase().includes(q)) {{ tr.style.display = ''; visibles++; }}
    else {{ tr.style.display = 'none'; }}
  }});
  tip.textContent = visibles === 0 ? '\u26a0\ufe0f No se encontraron productos para "' + term + '"'
    : visibles + ' producto' + (visibles !== 1 ? 's' : '') + ' encontrado' + (visibles !== 1 ? 's' : '') + ' para "' + term + '"';
  tip.style.color = visibles === 0 ? '#c0392b' : '#15803d';
}}
</script>
</body>
</html>"""


# ─── 4. AUTENTICACIÓN MICROSOFT GRAPH ─────────────────────────────────────────

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    endpoints = [
        'https://login.microsoftonline.com/' + tenant_id + '/oauth2/v2.0/token',
        'https://login.microsoftonline.com/' + tenant_id + '/oauth2/token',
    ]
    scopes = ['https://graph.microsoft.com/.default', 'https://graph.microsoft.com/']
    for i, url in enumerate(endpoints):
        data = urllib.parse.urlencode({
            'grant_type':    'client_credentials',
            'client_id':     client_id,
            'client_secret': client_secret,
            'scope' if i == 0 else 'resource': scopes[i]
        }).encode()
        req = urllib.request.Request(url, data=data, method='POST')
        print("🔑 Intentando endpoint " + str(i+1) + ": " + url)
        try:
            with urllib.request.urlopen(req) as resp:
                print("✅ Token obtenido con endpoint " + str(i+1))
                return json.loads(resp.read())['access_token']
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            print("⚠️  Endpoint " + str(i+1) + " falló: " + str(e.code) + " — " + body)
            if i == len(endpoints) - 1:
                print("ERROR: Todos los endpoints fallaron.")
                sys.exit(1)
            print("↪️  Intentando siguiente...")


# ─── 5. ENVÍA EL CORREO VIA GRAPH ─────────────────────────────────────────────

def send_email(token: str, sender: str, recipients: list, subject: str, html_body: str):
    payload = json.dumps({
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': r.strip()}} for r in recipients]
        },
        'saveToSentItems': True
    }).encode('utf-8')
    url = 'https://graph.microsoft.com/v1.0/users/' + sender + '/sendMail'
    req = urllib.request.Request(url, data=payload, method='POST',
        headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req) as resp:
            print("✅ Correo enviado (HTTP " + str(resp.status) + ")")
    except urllib.error.HTTPError as e:
        print("ERROR enviando correo: " + str(e.code) + " — " + e.read().decode())
        sys.exit(1)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = os.path.join(os.path.dirname(__file__), 'cowork_config.json')
    if not os.path.exists(config_path):
        return {}
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

if __name__ == '__main__':
    cfg = load_config()
    tenant_id     = os.environ.get('MS_TENANT_ID',     cfg.get('ms_tenant_id',     '9876dbde-5a7f-4139-8c2d-60a4395fd7d6')).strip()
    client_id     = os.environ.get('MS_CLIENT_ID',     cfg.get('ms_client_id',     '0c8bd7f5-a027-4da9-9d11-ccc27682b0ec')).strip()
    client_secret = os.environ.get('MS_CLIENT_SECRET', cfg.get('ms_client_secret', '')).strip()
    sender_email  = os.environ.get('SENDER_EMAIL',     cfg.get('sender_email',     '')).strip()
    recipients    = [r.strip() for r in os.environ.get(
                        'RECIPIENT_EMAILS', cfg.get('recipient_emails', '')).split(',')]

    print("🔍 Tenant ID  : '" + tenant_id + "' (len=" + str(len(tenant_id)) + ")")
    print("🔍 Client ID  : '" + client_id + "' (len=" + str(len(client_id)) + ")")
    print("🔍 Secret len : " + str(len(client_secret)) + " chars")
    print("🔍 Sender     : '" + sender_email + "'")

    html_path = os.path.join(os.path.dirname(__file__), '..', 'Buscador_Inventario_2026.html')
    print("📂 Leyendo inventario desde: " + html_path)
    inv = extract_inv_from_html(html_path)
    if not inv:
        print("ERROR: Inventario vacío.")
        sys.exit(1)
    print("✅ " + str(len(inv)) + " productos cargados")

    rot = extract_rot_from_html(html_path)
    print("📈 " + str(len(rot)) + " registros de rotación cargados")

    stats = calculate_stats(inv)
    print("📊 Stats: Total=" + str(stats['total']) + " | Con stock=" + str(stats['con_stock']) + " | Agotados=" + str(stats['agotados']))

    now = datetime.now()
    day      = DAYS_ES.get(now.strftime('%A'), now.strftime('%A'))
    month    = MONTHS_ES.get(now.strftime('%B'), now.strftime('%B'))
    date_str = day + " " + str(now.day) + " de " + month + " de " + str(now.year)

    html_body = generate_html(inv, stats, date_str, rot)
    subject   = "📦 Reporte Inventario E&G — " + date_str

    print("🔑 Obteniendo token Microsoft...")
    token = get_access_token(tenant_id, client_id, client_secret)

    print("📧 Enviando correo a: " + ', '.join(recipients))
    send_email(token, sender_email, recipients, subject, html_body)
    print("🎉 Reporte enviado exitosamente.")
