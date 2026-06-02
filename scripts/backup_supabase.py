"""
ENERGY — Respaldo diario de Supabase
====================================
Descarga las tablas de Supabase (inventario, kardex, remisiones, conteos)
y las guarda como JSON en la carpeta backups/. Como se sobrescriben los
mismos archivos cada día, el historial de Git guarda automáticamente una
copia de cada día (se puede recuperar cualquier fecha anterior).

Lo ejecuta GitHub Actions (.github/workflows/backup-supabase.yml) una vez al día.
"""
import json
import os
import datetime
import urllib.request

BASE = "https://juprjevxkcitqpsnemto.supabase.co/rest/v1"
# Clave pública (anon) — la misma que ya usa la plataforma en el navegador (solo lectura aquí)
KEY = "sb_publishable_zZrmpmvqbz4AJCGHRHQ8Xw_8tnf5ObM"
TABLES = ["productos", "kardex", "familias", "conteos", "conteo_items", "remisiones"]

os.makedirs("backups", exist_ok=True)


def fetch(table):
    req = urllib.request.Request(f"{BASE}/{table}?select=*&limit=100000")
    req.add_header("apikey", KEY)
    req.add_header("Authorization", "Bearer " + KEY)
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.loads(r.read().decode("utf-8"))


summary = {}
for t in TABLES:
    try:
        data = fetch(t)
        with open(f"backups/{t}.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        summary[t] = len(data)
        print(f"  {t}: {len(data)} registros")
    except Exception as e:
        summary[t] = f"ERROR: {e}"
        print(f"  {t}: ERROR {e}")

with open("backups/_resumen.json", "w", encoding="utf-8") as f:
    json.dump(
        {"actualizado": datetime.datetime.utcnow().isoformat() + "Z", "registros": summary},
        f, ensure_ascii=False, indent=1,
    )
print("Respaldo completado:", summary)
