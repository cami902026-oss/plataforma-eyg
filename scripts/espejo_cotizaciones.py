"""
Espejo diario de cotizaciones: Supabase -> data/cotizaciones.json

30-sep-2026. La plataforma dejó de subir data/cotizaciones.json (4,9 MB) en cada
guardado: la lista de cotizaciones se sincroniza desde Supabase, que es la fuente
real desde el 13-jul. Este archivo queda como RESPALDO y como punto de arranque
para un equipo nuevo (la plataforma lo carga una vez y luego pide a Supabase solo
lo que cambió). Lo corre el respaldo diario (backup-supabase.yml).

Mismas reglas de unión que la plataforma (_cotizAplicarFilas en Index.html):
  · el registro completo vive en `extra`; estado y adjudicación mandan las columnas;
  · gana la versión más reciente (updatedAt); un borrado del JSON es pegajoso;
  · lo que solo está en el JSON se conserva (nunca se pierde un registro).
Escribe el archivo solo si algo cambió.

Prueba sin escribir:  python scripts/espejo_cotizaciones.py --prueba
"""
import json, os, sys, urllib.request

BASE = (os.environ.get("SUPABASE_URL") or "https://juprjevxkcitqpsnemto.supabase.co").rstrip("/") + "/rest/v1"
KEY = "sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM"
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = os.path.join(RAIZ, "data", "cotizaciones.json")
COLS = "id,estado,deleted,updated_at,items_adjudicados,valor_adjudicado,adjudicacion_por,adjudicacion_at,factura,extra"


def traer():
    filas, off = [], 0
    while True:
        req = urllib.request.Request(f"{BASE}/cotizaciones?extra=not.is.null&select={COLS}&order=id&limit=500&offset={off}")
        req.add_header("apikey", KEY)
        req.add_header("Authorization", "Bearer " + KEY)
        with urllib.request.urlopen(req, timeout=60) as r:
            lote = json.loads(r.read().decode("utf-8"))
        filas += lote
        if len(lote) < 500:
            return filas
        off += 500


def registro(f):
    if not f.get("id") or not isinstance(f.get("extra"), dict):
        return None
    c = dict(f["extra"])
    c["id"] = f["id"]
    if f.get("estado"):
        c["estado"] = f["estado"]
    if f.get("deleted"):
        c["deleted"] = True
        c.setdefault("deletedAt", f.get("updated_at") or c.get("updatedAt") or "")
    if f.get("items_adjudicados"):
        c["itemsAdjudicados"] = f["items_adjudicados"]
    if f.get("valor_adjudicado") is not None:
        c["valorAdjudicado"] = f["valor_adjudicado"]
    if f.get("adjudicacion_por") and not c.get("adjudicacionPor"):
        c["adjudicacionPor"] = f["adjudicacion_por"]
    if f.get("adjudicacion_at") and not c.get("adjudicacionAt"):
        c["adjudicacionAt"] = f["adjudicacion_at"]
    if f.get("factura") and not c.get("factura"):
        c["factura"] = f["factura"]
    return c


def main():
    prueba = "--prueba" in sys.argv
    with open(ARCHIVO, encoding="utf-8") as fh:
        texto_antes = fh.read()
    actual = json.loads(texto_antes)
    pos = {c.get("id"): i for i, c in enumerate(actual) if isinstance(c, dict)}
    nuevas = cambiadas = 0
    for f in traer():
        rec = registro(f)
        if not rec:
            continue
        i = pos.get(rec["id"])
        if i is None:
            if rec.get("deleted"):
                continue
            actual.append(rec)
            pos[rec["id"]] = len(actual) - 1
            nuevas += 1
            continue
        loc = actual[i]
        if loc.get("deleted") and not rec.get("deleted"):
            continue
        if not rec.get("deleted") and str(loc.get("updatedAt") or "") > str(rec.get("updatedAt") or ""):
            continue
        if json.dumps(loc, sort_keys=True) == json.dumps(rec, sort_keys=True):
            continue
        actual[i] = rec
        cambiadas += 1
    texto = json.dumps(actual, ensure_ascii=False, indent=2)
    print(f"Espejo de cotizaciones: {len(actual)} registros · {nuevas} nuevas · {cambiadas} actualizadas desde Supabase")
    if texto == texto_antes or (not nuevas and not cambiadas):
        print("Sin cambios: no se escribe.")
        return
    if prueba:
        print("--prueba: no se escribe.")
        return
    with open(ARCHIVO, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(texto)
    print("Escrito", ARCHIVO)


if __name__ == "__main__":
    main()
