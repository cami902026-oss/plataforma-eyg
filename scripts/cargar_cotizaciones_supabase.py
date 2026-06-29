# -*- coding: utf-8 -*-
"""
Carga las cotizaciones a Supabase (tablas: cotizaciones + cotizacion_items).

Fuentes y prioridad de merge (gana el de mayor prioridad por id):
   1) Plataforma web  (data/cotizaciones.json)            <- mayor prioridad
   2) LIBRO 3 meses   (preview historicas_preview.json, ya incluye historico viejo)
   3) Historico viejo (incluido dentro del preview)

Requisitos previos:
   - Haber corrido extraer_libro_3meses.py (genera historicas_preview.json)
   - Haber creado las tablas con cotizaciones_supabase_schema.sql en Supabase

Modo:
   python cargar_cotizaciones_supabase.py            -> DRY RUN (no inserta; prueba conexion)
   python cargar_cotizaciones_supabase.py --write    -> inserta/actualiza en Supabase
"""
import json, os, sys, urllib.request, urllib.error

BASE = "https://juprjevxkcitqpsnemto.supabase.co/rest/v1"
KEY  = "sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM"

DATA    = r"C:\Users\Lenovo\OneDrive\Escritorio\Energy_bot\plataforma-eyg\data"
WEB     = os.path.join(DATA, "cotizaciones.json")
TMPDIR  = r"C:\Users\Lenovo\AppData\Local\Temp\eyg_cotiz"
os.makedirs(TMPDIR, exist_ok=True)
PREVIEW = os.path.join(TMPDIR, "historicas_preview.json")
OUT_COT = os.path.join(TMPDIR, "sb_cotizaciones.json")
OUT_IT  = os.path.join(TMPDIR, "sb_items.json")


# Cotizaciones de prueba que NO deben cargarse (id normalizado a mayusculas sin espacios/guiones)
EXCLUIR_IDS = {"LM1001", "LM1002"}

def _norm_id(v):
    import re
    return re.sub(r"[\s\-_]", "", str(v or "")).upper()

def numornull(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except Exception:
        return None

def dateornull(v):
    s = str(v or "")[:10]
    return s if len(s) == 10 and s[4] == "-" else None

def txt(v):
    return str(v).strip() if v not in (None, "") else None

def intornull(v):
    try:
        return int(float(v))
    except Exception:
        return None


def build():
    flat = json.load(open(PREVIEW, encoding="utf-8"))
    web  = json.load(open(WEB, encoding="utf-8"))

    headers = {}   # id -> dict cabecera
    items   = {}   # id -> [items]

    # ---- Fuente 2/3: preview (historico viejo + LIBRO 3 meses, ya consolidado) ----
    for r in flat:
        cid = str(r.get("id") or "").strip()
        if not cid:
            continue
        h = headers.setdefault(cid, {
            "id": cid, "fecha": None, "fecha_venc": None, "cliente": None,
            "contacto": None, "correo": None, "estado": None,
            "subtotal": None, "iva": None, "total": None,
            "forma_pago": None, "sitio_entrega": None, "validez": None,
            "observaciones": None, "vendedor": None, "realizada_por": None,
            "aprobada_por": None, "fuente": "Historico/LIBRO",
        })
        h["fecha"]         = dateornull(r.get("fecha")) or h["fecha"]
        h["fecha_venc"]    = dateornull(r.get("fecha_venc")) or h["fecha_venc"]
        h["cliente"]       = txt(r.get("cliente")) or h["cliente"]
        h["contacto"]      = txt(r.get("contacto")) or h["contacto"]
        h["correo"]        = txt(r.get("correo")) or h["correo"]
        h["estado"]        = txt(r.get("estado")) or h["estado"]
        h["subtotal"]      = numornull(r.get("subtotal")) if r.get("subtotal") else h["subtotal"]
        h["total"]         = numornull(r.get("total")) if r.get("total") else h["total"]
        h["forma_pago"]    = txt(r.get("forma_pago")) or h["forma_pago"]
        h["sitio_entrega"] = txt(r.get("sitio_entrega")) or h["sitio_entrega"]
        h["validez"]       = txt(r.get("validez")) or h["validez"]
        h["observaciones"] = txt(r.get("observaciones")) or h["observaciones"]
        # item real (item con descripcion)
        if txt(r.get("desc")):
            items.setdefault(cid, []).append({
                "cotizacion_id": cid, "item": intornull(r.get("item")),
                "descripcion": txt(r.get("desc")), "udm": txt(r.get("udm")),
                "qty": numornull(r.get("qty")), "v_unit": numornull(r.get("v_unit")),
                "v_total": numornull(r.get("v_total")), "marca": txt(r.get("marca")),
                "proveedor": txt(r.get("proveedor")),
                "precio_proveedor": numornull(r.get("costo")),
                "tiempo_entrega": txt(r.get("tiempo")),
                "desviacion_tecnica": txt(r.get("desv")),
            })

    # ---- Fuente 1: plataforma web (PISA por id) ----
    for c in web:
        cid = str(c.get("id") or "").strip()
        if not cid:
            continue
        headers[cid] = {
            "id": cid, "fecha": dateornull(c.get("fecha")),
            "fecha_venc": dateornull(c.get("fechaVenc")), "cliente": txt(c.get("cliente")),
            "contacto": txt(c.get("contacto")), "correo": txt(c.get("correo")),
            "estado": txt(c.get("estado")), "subtotal": numornull(c.get("subtotal")),
            "iva": numornull(c.get("ivaVal")), "total": numornull(c.get("total")),
            "forma_pago": txt(c.get("formaPago")), "sitio_entrega": txt(c.get("sitioEntrega")),
            "validez": txt(c.get("validez")), "observaciones": txt(c.get("observaciones")),
            "vendedor": txt(c.get("vendedor")), "realizada_por": txt(c.get("realizadaPor")),
            "aprobada_por": txt(c.get("aprobadaPor")), "fuente": "Plataforma",
        }
        its = []
        for n, it in enumerate(c.get("items") or [], start=1):
            its.append({
                "cotizacion_id": cid, "item": intornull(it.get("id")) or n,
                "descripcion": txt(it.get("desc")), "udm": txt(it.get("udm")),
                "qty": numornull(it.get("qty")), "v_unit": numornull(it.get("precio")),
                "v_total": (numornull(it.get("qty")) or 0) * (numornull(it.get("precio")) or 0),
                "marca": txt(it.get("marca")), "proveedor": txt(it.get("proveedor")),
                "precio_proveedor": numornull(it.get("precioProveedor")),
                "tiempo_entrega": txt(it.get("tiempoEntrega")),
                "desviacion_tecnica": txt(it.get("desvTec")),
            })
        items[cid] = its

    # excluir cotizaciones de prueba (LM1001/LM1002 'Yo'/'Yotas')
    cot_rows  = [h for h in headers.values() if _norm_id(h["id"]) not in EXCLUIR_IDS]
    item_rows = [it for cid, lst in items.items() if _norm_id(cid) not in EXCLUIR_IDS for it in lst]
    return cot_rows, item_rows


def req(method, endpoint, body=None, prefer=None):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    r = urllib.request.Request(f"{BASE}/{endpoint}", data=data, method=method)
    r.add_header("apikey", KEY); r.add_header("Authorization", "Bearer " + KEY)
    r.add_header("Content-Type", "application/json")
    if prefer:
        r.add_header("Prefer", prefer)
    with urllib.request.urlopen(r, timeout=120) as resp:
        t = resp.read().decode("utf-8")
        return resp.status, t


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def main():
    write = "--write" in sys.argv
    cot_rows, item_rows = build()

    from collections import Counter
    fuentes = Counter(c["fuente"] for c in cot_rows)
    con_obs = sum(1 for c in cot_rows if c["observaciones"])
    print("=" * 60)
    print(f"COTIZACIONES (cabeceras) a cargar : {len(cot_rows)}")
    for f, n in fuentes.items():
        print(f"     - fuente {f:18}: {n}")
    print(f"     - con observaciones          : {con_obs}")
    print(f"ITEMS (lineas) a cargar           : {len(item_rows)}")
    print("=" * 60)

    json.dump(cot_rows, open(OUT_COT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(item_rows, open(OUT_IT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"Preview tablas: \n  {OUT_COT}\n  {OUT_IT}")

    # prueba de conexion (GET a tabla existente)
    try:
        st, _ = req("GET", "productos?select=id&limit=1")
        print(f"\nConexion Supabase OK (GET productos -> {st})")
    except urllib.error.HTTPError as e:
        print(f"\nConexion Supabase ERROR {e.code}: {e.read().decode()[:200]}")
    except Exception as e:
        print(f"\nConexion Supabase ERROR: {e}")

    if not write:
        print("\n(DRY RUN - no se inserto nada. Crea las tablas con el .sql y luego usa --write.)")
        return

    # ---- ESCRITURA ----
    print("\nLimpiando items previos...")
    try:
        req("DELETE", "cotizacion_items?id=gte.0", prefer="return=minimal")
    except urllib.error.HTTPError as e:
        print("  (sin items previos o tabla vacia)", e.code)

    print("Upsert de cabeceras...")
    tot = 0
    for ch in chunks(cot_rows, 500):
        st, _ = req("POST", "cotizaciones", ch, prefer="resolution=merge-duplicates,return=minimal")
        tot += len(ch)
        print(f"  cabeceras {tot}/{len(cot_rows)}  (HTTP {st})")

    print("Insertando items...")
    tot = 0
    for ch in chunks(item_rows, 500):
        st, _ = req("POST", "cotizacion_items", ch, prefer="return=minimal")
        tot += len(ch)
        print(f"  items {tot}/{len(item_rows)}  (HTTP {st})")

    print("\n*** CARGA COMPLETA en Supabase ***")


if __name__ == "__main__":
    main()
