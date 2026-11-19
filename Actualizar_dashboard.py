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
candidates = sorted(
    (p for p in BASE.glob("DATA GRUAS*.xlsx") if p.is_file()),
    key=lambda p: p.stat().st_mtime
)
XLSX = candidates[-1] if candidates else BASE / "DATA GRUAS.xlsx"
HTML = BASE / "index.html"
if not HTML.exists():
    candidates = sorted(p for p in BASE.glob("index*.html") if p.is_file())
    if candidates:
        HTML = candidates[-1]

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
report_observations = []
equipment_hours = {}

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

    # Consolidado de horas por equipo y estado para los gráficos del reporte semanal.
    equipment_hours.setdefault(d, {}).setdefault(equipo, {"PRODUCCION": 0.0, "STAND_BY": 0.0,
                                                               "NO_DISPONIBLE": 0.0, "NO_PROGRAMADO": 0.0})
    equipment_col = {
        "PRODUCCION": "Horas Efectivas o metrados",
        "STAND_BY": "HDNT",
        "NO_PROGRAMADO": "HNP",
        "NO_DISPONIBLE": "HND",
    }.get(status)
    if equipment_col:
        equipment_value = num(row[idx[equipment_col]])
        if equipment_value is None:
            equipment_value = num(row[idx["Horas Efectivas o metrados"]])
        if equipment_value is not None:
            equipment_hours[d][equipo][status] += equipment_value

    # Observaciones para el nuevo apartado "Reporte semanal" (columna Y).
    if observ:
        report_col = {
            "PRODUCCION": "Horas Efectivas o metrados",
            "STAND_BY": "HDNT",
            "NO_PROGRAMADO": "HNP",
            "NO_DISPONIBLE": "HND",
        }.get(status, "Horas Efectivas o metrados")
        report_hours = num(row[idx[report_col]]) if report_col in idx else None
        if report_hours is None:
            report_hours = num(row[idx["Horas Efectivas o metrados"]])
        report_observations.append({
            "date": d,
            "status": status,
            "equipo": equipo,
            "actividad": actividad,
            "horas": report_hours,
            "observaciones": observ,
        })

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

# ---------- CURVA S HÍBRIDA: HISTÓRICO + MES ACTUAL ----------
# Histórico hasta el 12/11/2026: hoja "Curva S Acum".
# Mes actual 13/11/2026 al 12/12/2026: hoja "Curva S".
# Desde el 13/11 la curva se reinicia: los acumulados de la hoja Curva "S" se usan directamente, sin sumar el cierre anterior.
if 'Curva "S" Acum' not in wb.sheetnames:
    raise RuntimeError('No se encontró la hoja "Curva \\"S\\" Acum" en el Excel.')
if 'Curva "S"' not in wb.sheetnames:
    raise RuntimeError('No se encontró la hoja "Curva \\"S\\"" en el Excel.')

ws_hist = wb['Curva "S" Acum']
ws_current = wb['Curva "S"']
historical_cutoff = "2026-11-12"
current_start = "2026-11-13"
current_end = "2026-12-12"

# Leer histórico desde "Curva S Acum", únicamente hasta el 12/11 inclusive.
historical = {}
for c in range(3, ws_hist.max_column + 1):
    dk = date_key(ws_hist.cell(6, c).value) or date_key(ws_hist.cell(7, c).value)
    if not dk or dk > historical_cutoff:
        continue
    historical[dk] = {
        "target": num(ws_hist.cell(9, c).value) or 0.0,
        "Camión Grúa 23 TN": num(ws_hist.cell(13, c).value),
        "Grua Semitrailer 20 TN": num(ws_hist.cell(14, c).value),
        "Tadano 60 TN": num(ws_hist.cell(15, c).value),
        "daily_target": num(ws_hist.cell(8, c).value) or 0.0,
        "daily_camion": num(ws_hist.cell(10, c).value),
        "daily_semi": num(ws_hist.cell(11, c).value),
        "daily_tadano": num(ws_hist.cell(12, c).value),
    }

if historical_cutoff not in historical:
    raise RuntimeError('La hoja "Curva S Acum" no contiene el cierre histórico del 12/11/2026.')

# Leer el mes actual desde "Curva S". A partir del 13/11 la curva se reinicia;
# se usan directamente sus acumulados, sin arrastrar los valores hasta el 12/11.
current = {}
for c in range(3, ws_current.max_column + 1):
    dk = date_key(ws_current.cell(29, c).value)
    if not dk or dk < current_start or dk > current_end:
        continue
    current[dk] = {
        "target": num(ws_current.cell(31, c).value) or 0.0,
        "Camión Grúa 23 TN": num(ws_current.cell(35, c).value),
        "Grua Semitrailer 20 TN": num(ws_current.cell(36, c).value),
        "Tadano 60 TN": num(ws_current.cell(37, c).value),
        "daily_target": num(ws_current.cell(30, c).value) or 0.0,
        "daily_camion": num(ws_current.cell(32, c).value),
        "daily_semi": num(ws_current.cell(33, c).value),
        "daily_tadano": num(ws_current.cell(34, c).value),
    }

if not current:
    raise RuntimeError('La hoja "Curva S" no contiene fechas entre el 13/11/2026 y el 12/12/2026.')

# Unir sin duplicar: histórico hasta el 12/11 + mes actual del 13/11 al 12/12.
merged = dict(historical)
merged.update(current)
curve_dates = sorted(merged)

series = []
curve_daily = []
for d in curve_dates:
    item = merged[d]
    is_current_period = d >= current_start
    is_projected = bool(last_data and d > last_data)
    series.append({
        "date": d,
        "target": item["target"],
        "projected": is_projected,
        "Camión Grúa 23 TN": item["Camión Grúa 23 TN"],
        "Grua Semitrailer 20 TN": item["Grua Semitrailer 20 TN"],
        "Tadano 60 TN": item["Tadano 60 TN"],
    })
    curve_daily.append({
        "date": d,
        "target": item["daily_target"],
        "Camión Grúa 23 TN": item["daily_camion"],
        "Grua Semitrailer 20 TN": item["daily_semi"],
        "Tadano 60 TN": item["daily_tadano"],
        "projected": is_projected,
    })

curva = {
    "dates": curve_dates,
    "series": series,
    "daily": curve_daily,
    "equipment": ["Camión Grúa 23 TN", "Grua Semitrailer 20 TN", "Tadano 60 TN"],
    "defaultStart": curve_dates[0] if curve_dates else "",
    "defaultEnd": current_end if current_end in curve_dates else (curve_dates[-1] if curve_dates else ""),
    "lastActual": last_data,
    "source": 'Curva reiniciada: Curva "S" Acum hasta 12/11 + Curva "S" desde 13/11 (acumulados propios)',
    "historicalCutoff": historical_cutoff,
    "currentPeriod": [current_start, current_end],
    "projectionEndpointDay": 12
}

payload = {
    "dates": dates,
    "data": detail,
    "production": production,
    "equipmentHours": equipment_hours,
    "source": "DATA",
    "range": [dates[0], dates[-1]] if dates else ["", ""],
    "dmUsage": dm_usage,
    "usage": usage,
    "dmTarget": 0.92
}

text = HTML.read_text(encoding="utf-8")
text = replace_const(text, "PAYLOAD", payload)
text = replace_const(text, "REPORT_OBSERVATIONS", report_observations)
text = replace_const(text, "CURVA_DATA", curva)

# Compatibilidad con el JavaScript existente:
# el gráfico diario usa CURVA_EXCEL_DAILY.
text = replace_const(text, "CURVA_EXCEL_DAILY", curva["daily"])

# Build marker guarantees a visible content change for Git when data changes.
marker = f"<!-- DASHBOARD_BUILD: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -->"
text = re.sub(r"<!-- DASHBOARD_BUILD: .*? -->\s*", "", text, count=1)
text = marker + "\n" + text
HTML.write_text(text, encoding="utf-8")

print("OK - index.html actualizado")
print("DATA:", payload["range"][0], "->", payload["range"][1])
print("Fechas:", len(dates))
print("Observaciones para Reporte semanal:", len(report_observations))
print("DM/Usage:", len(dm_usage))
print("Curva S (reinicio 13/11):", len(curve_dates), "puntos")
if curve_dates:
    end = curve_dates[-1]
    end_values = next((p for p in curve_daily if p["date"] == end), {})
    print("Cierre Curva S:", end)
    print("Target diario al cierre:", end_values.get("target"))
    print("Reales al cierre:", {
        "Camión Grúa 23 TN": end_values.get("Camión Grúa 23 TN"),
        "Grua Semitrailer 20 TN": end_values.get("Grua Semitrailer 20 TN"),
        "Tadano 60 TN": end_values.get("Tadano 60 TN"),
    })
