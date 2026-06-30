#!/usr/bin/env python3
"""
update_hospedaje.py
AAFY 2026 – Impuesto sobre Hospedaje
Descarga archivos XLS de Google Drive, calcula proyecciones y actualiza hospedaje.html
"""

import os
import re
import json
import requests
import xlrd
import unicodedata

API_KEY   = os.environ["DRIVE_API_KEY"]
FOLDER_ID = "1MjWHAaGzvzPnG9A_s3GPE2DWpNjWleFd"
HTML_FILE = "hospedaje.html"

METAS = {
    1: 16765704, 2: 15528008, 3: 15696442,  4: 15728438,
    5: 14668401, 6: 13065948, 7: 12698332,  8: 14813279,
    9: 15294695,10: 12925171,11: 14923585, 12: 16659886,
}

MONTH_NAMES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
}

PREV_YEAR_TOTALS = [
    17181838,17239964,16991252,17746000,15075887,13542804,
    14086552,16086465,18251188,15865876,15799923,19033141,
]

# Columnas (base 0 = columna A)
HSP_RFC     = 0   # Columna A
HSP_CONTRIB = 1   # Columna B
HSP_PERIODO = 5   # Columna F
HSP_R       = 17  # Columna R  (suma)
HSP_O       = 14  # Columna O  (resta)
# Recaudación = R - O


def normalize(s):
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode().lower()


def drive_list_files():
    url = (
        f"https://www.googleapis.com/drive/v3/files"
        f"?q=%27{FOLDER_ID}%27+in+parents+and+trashed%3Dfalse"
        f"&fields=files(id,name)&pageSize=100&key={API_KEY}"
    )
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.json().get("files", [])


def drive_download(file_id):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&key={API_KEY}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


def _find_data_start(sheet, col):
    """Busca la primera fila con un RFC válido (≥12 chars, no solo letras)."""
    for i in range(min(20, sheet.nrows)):
        v = str(sheet.cell_value(i, col) if col < sheet.ncols else "").strip()
        if len(v) >= 12 and not re.match(r'^[a-zA-Z\s]+$', v):
            return i
    return -1


def parse_hospedaje(content):
    try:
        wb = xlrd.open_workbook(file_contents=content)
    except Exception as e:
        print(f"  [WARN] No se pudo abrir el workbook: {e}")
        return []

    sheet = wb.sheets()[0]
    start = _find_data_start(sheet, HSP_RFC)
    if start < 0:
        print("  [WARN] No se encontró fila de datos válida")
        return []

    records = []
    for i in range(start, sheet.nrows):
        rfc = str(sheet.cell_value(i, HSP_RFC) if HSP_RFC < sheet.ncols else "").strip().upper()
        if not rfc or len(rfc) < 12:
            continue

        # Período YYYYMM: float → int → str, validado a 6 dígitos
        raw_p = sheet.cell_value(i, HSP_PERIODO) if HSP_PERIODO < sheet.ncols else ""
        p = str(int(float(raw_p))) if str(raw_p).replace(".", "", 1).isdigit() else str(raw_p).strip()
        if len(p) != 6:
            continue

        R = float(sheet.cell_value(i, HSP_R)) if HSP_R < sheet.ncols else 0.0
        O = float(sheet.cell_value(i, HSP_O)) if HSP_O < sheet.ncols else 0.0
        contrib = str(sheet.cell_value(i, HSP_CONTRIB) if HSP_CONTRIB < sheet.ncols else "").strip()

        records.append({
            "rfc":         rfc,
            "periodo":     p,
            "recaudacion": R - O,
            "contrib":     contrib,
        })

    return records


# ── Utilidades de período ──────────────────────────────────────────────────

def prev_period(p):
    p = str(p)
    y, m = int(p[:4]), int(p[4:])
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f"{y}{str(m).zfill(2)}"


def format_period(p):
    s = str(p)
    labels = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    try:
        return labels[int(s[4:6])] + "-" + s[2:4]
    except Exception:
        return s


def get_dominant(records):
    t = {}
    for r in records:
        if r["periodo"]:
            t[r["periodo"]] = t.get(r["periodo"], 0) + r["recaudacion"]
    return max(t, key=lambda k: t[k]) if t else None


def get_missing_periods(paid_set, dominant, max_back, stop_before=None):
    out = []
    p = str(dominant)
    while p not in paid_set:
        if stop_before and int(p) < int(stop_before):
            break
        out.append(p)
        p = prev_period(p)
        if len(out) >= max_back:
            break
    return out


# ── Cálculo de proyección ──────────────────────────────────────────────────

def compute_month(month_num, all_month_data):
    cur      = all_month_data.get(month_num, [])
    dominant = get_dominant(cur)
    acumulado = sum(r["recaudacion"] for r in cur)

    ref_months = [m for m in range(max(1, month_num - 4), month_num)
                  if all_month_data.get(m)]
    n_ref = len(ref_months)

    global_periods = {}
    global_contrib = {}
    for m in range(1, month_num + 1):
        for r in all_month_data.get(m, []):
            gp = global_periods.setdefault(r["rfc"], {})
            if r["periodo"] not in gp or r["recaudacion"] > gp[r["periodo"]]:
                gp[r["periodo"]] = r["recaudacion"]
            if r["rfc"] not in global_contrib and r["contrib"]:
                global_contrib[r["rfc"]] = r["contrib"]

    rfc_ref_count    = {}
    rfc_ref_periods  = {}
    rfc_contrib      = {}
    paid_2026_in_ref = {}

    for rm in ref_months:
        seen = set()
        for r in all_month_data.get(rm, []):
            rp = rfc_ref_periods.setdefault(r["rfc"], {})
            if r["periodo"] not in rp or r["recaudacion"] > rp[r["periodo"]]:
                rp[r["periodo"]] = r["recaudacion"]
            if str(r["periodo"]).startswith("2026"):
                paid_2026_in_ref.setdefault(r["rfc"], set()).add(str(r["periodo"]))
            if r["rfc"] not in seen:
                seen.add(r["rfc"])
                rfc_ref_count[r["rfc"]] = rfc_ref_count.get(r["rfc"], 0) + 1
            if r["rfc"] not in rfc_contrib and r["contrib"]:
                rfc_contrib[r["rfc"]] = r["contrib"]

    candidates = set()
    for rfc, cnt in rfc_ref_count.items():
        if cnt >= 2:
            candidates.add(rfc)
    for rfc, ps in paid_2026_in_ref.items():
        if len(ps) >= 2:
            candidates.add(rfc)

    omisos = []
    for rfc in candidates:
        cnt = rfc_ref_count.get(rfc, 0)
        p26 = paid_2026_in_ref.get(rfc, set())
        if cnt < 2 and len(p26) < 2:
            continue
        paid_set = set(global_periods.get(rfc, {}).keys())
        if not dominant or dominant in paid_set:
            continue
        has_2026 = any(p.startswith("2026") for p in paid_set)
        missing = (
            get_missing_periods(paid_set, dominant, 12)
            if has_2026
            else get_missing_periods(paid_set, dominant, 12, "202601")
        )
        if not missing:
            continue
        ref_amounts = list(rfc_ref_periods.get(rfc, {}).values())
        if not ref_amounts:
            continue
        avg = sum(ref_amounts) / len(ref_amounts)

        if not has_2026:
            seg = "omisos_totales"
        elif cnt >= n_ref:
            seg = "alta"
        elif cnt >= 3:
            seg = "media"
        else:
            seg = "seguimiento"

        omisos.append({
            "rfc":      rfc,
            "contrib":  rfc_contrib.get(rfc) or global_contrib.get(rfc, ""),
            "count":    cnt,
            "avg":      round(avg * len(missing)),
            "nMissing": len(missing),
            "pending":  [format_period(p) for p in missing],
            "seg":      seg,
        })

    omisos.sort(key=lambda o: -o["avg"])
    esperado   = sum(o["avg"] for o in omisos if o["seg"] in ("alta", "media"))
    proyeccion = acumulado + esperado
    meta       = METAS.get(month_num, 0)

    segments = {}
    for o in omisos:
        s = segments.setdefault(o["seg"], {"count": 0, "monto": 0, "omisos": []})
        s["count"] += 1
        s["monto"]  += o["avg"]
        s["omisos"].append({k: o[k] for k in ("rfc","contrib","avg","count","nMissing","pending")})
    for s in segments.values():
        s["monto"] = round(s["monto"])
        s["omisos"].sort(key=lambda x: -x["avg"])

    month_labels = {
        1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
        7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre",
    }

    return {
        "mes_label":        month_labels.get(month_num, str(month_num)),
        "mes_num":          month_num,
        "meta":             meta,
        "dominant_period":  int(dominant) if dominant else 0,
        "ref_months":       ref_months,
        "acumulado_real":   round(acumulado),
        "total_omisos":     len(omisos),
        "total_esperado":   round(esperado),
        "proyeccion_cierre": round(proyeccion),
        "meta_cruzada":     proyeccion >= meta,
        "pct_acumulado":    acumulado / meta * 100 if meta else 0,
        "pct_proyeccion":   proyeccion / meta * 100 if meta else 0,
        "segmentos":        segments,
        "omisos":           omisos[:5000],
    }


# ── HTML update ────────────────────────────────────────────────────────────

def load_existing_data(html_path):
    """Lee allData del HTML existente y filtra claves inválidas (YYYYMM)."""
    try:
        with open(html_path, encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'let allData\s*=\s*(\{.*?\});', content, re.DOTALL)
        if not m:
            return {}
        data = json.loads(m.group(1))
        # Solo conservar claves "1"–"12" (descartar claves YYYYMM u otras)
        return {k: v for k, v in data.items() if k.isdigit() and 1 <= int(k) <= 12}
    except Exception as e:
        print(f"[WARN] No se pudo leer allData existente: {e}")
        return {}


def update_html(html_path, new_data, last_updated):
    """Fusiona new_data con el HTML; solo actualiza si acumulado_real es mayor."""
    existing = load_existing_data(html_path)

    merged = dict(existing)
    for key, val in new_data.items():
        if key not in merged:
            merged[key] = val
        elif val.get("acumulado_real", 0) > merged[key].get("acumulado_real", 0):
            merged[key] = val

    with open(html_path, encoding="utf-8") as f:
        html = f.read()

    data_json = json.dumps(merged, ensure_ascii=False, separators=(",", ":"))
    html = re.sub(r'let allData\s*=\s*\{.*?\};', f'let allData = {data_json};', html, flags=re.DOTALL)

    updated_str = f"var lastUpdated = '{last_updated}';"
    html = re.sub(r"var lastUpdated\s*=\s*'.*?';", updated_str, html)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[OK] {html_path} actualizado — meses: {sorted(merged.keys(), key=int)}")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    from datetime import datetime, timezone, timedelta

    tz_mx = timezone(timedelta(hours=-6))
    last_updated = datetime.now(tz_mx).strftime("%d/%m/%Y %H:%M")

    print("=" * 40)
    print("update_hospedaje.py  |  AAFY 2026")
    print(f"Fecha/hora (MX): {last_updated}")
    print("=" * 40)

    files = drive_list_files()
    print(f"\nArchivos en Drive: {len(files)}")

    month_files = []
    for f in files:
        name_n = normalize(f["name"])
        mn = next((MONTH_NAMES[m] for m in MONTH_NAMES if m in name_n), None)
        if mn:
            month_files.append({**f, "num": mn})

    if not month_files:
        print("[ERROR] No se encontraron archivos con nombre de mes. Abortando.")
        return

    month_files.sort(key=lambda x: x["num"])
    print(f"Archivos detectados ({len(month_files)}): {[f['name'] for f in month_files]}\n")

    all_month_data = {}
    for f in month_files:
        print(f"Descargando: {f['name']}  (mes {f['num']})")
        content = drive_download(f["id"])
        records = parse_hospedaje(content)
        total_rec = sum(r["recaudacion"] for r in records)
        print(f"  → Registros: {len(records)} | Recaudación: ${total_rec:,.0f}")
        all_month_data.setdefault(f["num"], []).extend(records)

    print("\nCalculando proyecciones...")
    new_data = {}
    for num in sorted(all_month_data.keys()):
        result = compute_month(num, all_month_data)
        new_data[str(num)] = result
        print(f"  Mes {num:>2} ({result['mes_label']:<12}): "
              f"acumulado=${result['acumulado_real']:>15,.0f} | "
              f"omisos={result['total_omisos']}")

    update_html(HTML_FILE, new_data, last_updated)
    print("\n=== Proceso completado ===")


if __name__ == "__main__":
    main()
