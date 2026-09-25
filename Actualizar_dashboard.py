# -*- coding: utf-8 -*-
"""
Actualiza index.html desde DATA GRUAS.xlsx sin eliminar la persistencia de filtros.
Requiere openpyxl.
"""
from pathlib import Path
from datetime import datetime, date
import json
import math
import re
import openpyxl

BASE = Path(__file__).resolve().parent
XLSX = BASE / "DATA GRUAS.xlsx"
HTML = BASE / "index.html"

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return str(v).strip()

def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(v)
        return x if math.isfinite(x) else None
    except Exception:
        return None

def date_key(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except Exception:
                pass
    return None

def json_dump(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

def replace_const(text, name, value):
    pattern = rf"const {re.escape(name)} = .*?;\s*\n"
    replacement = f"const {name} = {json_dump(value)};\n"
    new, n = re.subn(pattern, replacement, text, count=1, flags=re.S)
    if n != 1:
        raise RuntimeError(f"No se encontró la constante {name} en index.html")
    return new

# ---------- DATA ----------
wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
ws = wb["DATA"]
rows = ws.iter_rows(values_only=True)
headers = list(next(rows))
idx = {clean(h): i for i, h in enumerate(headers) if clean(h)}

required = ["Fecha", "ESTADO", "Detalle de Actividad", "Equipo",
            "Horas Efectivas o metrados", "HDNT", "HND", "HNP", "Observaciones"]
missing = [h for h in required if h not in idx]
if missing:
    raise RuntimeError("Faltan columnas en DATA: " + ", ".join(missing))

dates_set = set()
detail = {}
production = {}

status_map = {
    "PRODUCCIÓN": "PRODUCCION",
    "PRODUCCION": "PRODUCCION",
    "STAND_BY": "STAND_BY",
    "NO_PROGRAMADO": "NO_PROGRAMADO",
    "NO DISPONIBLE": "NO_DISPONIBLE",
    "NO_DISPONIBLE": "NO_DISPONIBLE",
}

for row in rows:
    d = date_key(row[idx["Fecha"]])
    if not d:
        continue
    dates_set.add(d)

    status_raw = clean(row[idx["ESTADO"]]).upper().replace(" ", "_")
    status = status_map.get(status_raw, status_raw)
    equipo = clean(row[idx["Equipo"]])
    actividad = clean(row[idx["Detalle de Actividad"]])
    observ = clean(row[idx["Observaciones"]])

    if status == "PRODUCCION":
        h = num(row[idx["Horas Efectivas o metrados"]])
        if h is not None:
            production[d] = production.get(d, 0.0) + h
        continue

    if status not in ("STAND_BY", "NO_PROGRAMADO", "NO_DISPONIBLE"):
        continue

    # Las horas de cada estado se toman de su columna específica.
    col = {
        "STAND_BY": "HDNT",
        "NO_PROGRAMADO": "HNP",
        "NO_DISPONIBLE": "HND",
    }[status]
    h = num(row[idx[col]])
    if h is None:
        h = num(row[idx["Horas Efectivas o metrados"]])
    if h is None:
        continue

    detail.setdefault(d, {}).setdefault(status, []).append({
        "equipo": equipo,
        "actividad": actividad,
        "observaciones": observ,
        "horas": h
    })

dates = sorted(dates_set)
last_data = dates[-1] if dates else None

# ---------- DM / USAGE ----------
wsdm = wb["DM_USAGE"]
dm_rows = list(wsdm.iter_rows(values_only=True))
dm_headers = [clean(x) for x in dm_rows[0]]
di = {h:i for i,h in enumerate(dm_headers) if h}

dm_usage = {}
usage = {}

def normalize_equipment(name):
    s = clean(name)
    # El dashboard usa 23 TN como denominación ejecutiva, aunque la hoja DATA
    # pueda contener la denominación 36 TN.
    if s.upper() == "CAMIÓN GRÚA 23 TN":
        return "CAMIÓN GRÚA 23 TN"
    return s

for row in dm_rows[1:]:
    d = date_key(row[di.get("FECHA", -1)]) if "FECHA" in di else None
    if not d:
        continue
    eq = normalize_equipment(row[di.get("GRUA", -1)]) if "GRUA" in di else ""
    if eq not in ("Camión Grúa Semitrailer 20 TN", "CAMIÓN GRÚA 23 TN", "GRÚA TADANO 60 TN"):
        continue

    dm = num(row[di.get("DM", -1)]) if "DM" in di else None
    us = num(row[di.get("USAGE", -1)]) if "USAGE" in di else None
    dm_usage.setdefault(d, {})[eq] = 0.0 if dm is None else dm
    usage.setdefault(d, {})[eq] = 0.0 if us is None else us

# ---------- CURVA S ----------
wsc = wb['Curva "S"']
curve_dates = []
target_row = {}
s35, s36, s37 = {}, {}, {}

for c in range(1, wsc.max_column + 1):
    dk = date_key(wsc.cell(29, c).value)
    if not dk:
        continue
    curve_dates.append(dk)
    target_row[dk] = num(wsc.cell(31, c).value) or 0.0
    s35[dk] = num(wsc.cell(35, c).value)
    s36[dk] = num(wsc.cell(36, c).value)
    s37[dk] = num(wsc.cell(37, c).value)

curve_dates = sorted(set(curve_dates))
if last_data and curve_dates:
    # Mostrar hasta el primer día 12 posterior/al cierre de la curva.
    endpoint = None
    for d in curve_dates:
        if d >= last_data and d[:4] + d[5:7] and int(d[8:10]) == 12:
            endpoint = d
            break
    if endpoint is None:
        endpoint = curve_dates[-1]
    curve_dates = [d for d in curve_dates if d <= endpoint]

series = []
for d in curve_dates:
    series.append({
        "date": d,
        "target": target_row.get(d, 0.0),
        "projected": bool(last_data and d > last_data),
        "Camión Grúa 23 TN": s35.get(d),
        "Grua Semitrailer 20 TN": s36.get(d),
        "Tadano 60 TN": s37.get(d),
    })

curva = {
    "dates": curve_dates,
    "series": series,
    "equipment": ["Camión Grúa 23 TN", "Grua Semitrailer 20 TN", "Tadano 60 TN"],
    "defaultStart": next((d for d in curve_dates if d >= last_data), curve_dates[0] if curve_dates else ""),
    "defaultEnd": next((d for d in curve_dates if d.endswith("-12")), curve_dates[-1] if curve_dates else ""),
    "lastActual": last_data,
    "source": 'Curva "S"',
    "projectionEndpointDay": 12
}

payload = {
    "dates": dates,
    "data": detail,
    "production": production,
    "source": "DATA",
    "range": [dates[0], dates[-1]] if dates else ["", ""],
    "dmUsage": dm_usage,
    "usage": usage,
    "dmTarget": 0.92
}

text = HTML.read_text(encoding="utf-8")
text = replace_const(text, "PAYLOAD", payload)
text = replace_const(text, "CURVA_DATA", curva)

# Build marker guarantees a visible content change for Git when data changes.
marker = f"<!-- DASHBOARD_BUILD: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -->"
text = re.sub(r"<!-- DASHBOARD_BUILD: .*? -->\s*", "", text, count=1)
text = marker + "\n" + text
HTML.write_text(text, encoding="utf-8")

print("OK - index.html actualizado")
print("DATA:", payload["range"][0], "->", payload["range"][1])
print("Fechas:", len(dates))
print("DM/Usage:", len(dm_usage))
print("Curva S:", len(curve_dates), "puntos")
