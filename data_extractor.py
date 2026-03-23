"""
data_extractor.py
Centro Floricultor de Baja California

- Hojas WK####  → Excel en SharePoint (link público, sin autenticación)
- Hojas PR####  → Google Sheets (service account, sin cambios)
"""

import re
import io
import logging
import requests
import openpyxl
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
SHAREPOINT_URL = (
    "https://pacificafarms-my.sharepoint.com/:x:/g/personal/anahi_mora_cfbc_co"
    "/IQAQCb79SzHtRrTQR71pSNQcASOWqFXyeGGzEhUcT9FRRJ4?e=ClxLCN&download=1"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

RANCH_KEYS = ["PROP", "POSCO", "CAMPO", "ISABEL", "HOOPS", "CECILIA", "CHRISTINA", "ALBAHACA"]

CATEGORIAS_ORDEN = [
    "DESINFECCION Y FERTILIZACION",
    "AMPLIACION",
    "CULTIVO TIERRA, CHAROLAS",
    "MATERIAL VEGETAL",
    "PREPARACION DE SUELO",
    "FERTILIZANTES",
    "DESINFECCION / PLAGUICIDAS",
    "MANTENIMIENTO",
    "EXPANSION CECILIA 25",
    "RENOVACION DE SIEMBRA",
    "MATERIAL DE EMPAQUE",
]

SKIP = {"ACUMULADO", "GRAFICOS I-IV", "COMPARATIVO", "DATOS", "HOJA1", "SHEET1"}

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# HELPERS DE NORMALIZACIÓN
# ─────────────────────────────────────────────
def norm_ranch(s: str):
    s = str(s).upper().strip()
    if "PROP" in s:                                      return "Prop-RM"
    if "POSCO" in s:                                     return "PosCo-RM"
    if "CAMPO-VI" in s or "CAMPO-IV" in s:               return "Campo-VI"
    if "ALBAHACA" in s:                                  return "Albahaca-RM"
    if "HOOPS" in s:                                     return "HOOPS"
    if "CHRISTINA" in s:                                 return "Christina"
    if "CECILIA 25" in s:                                return "Cecilia 25"
    if "CECILIA" in s:                                   return "Cecilia"
    if "ISABEL" in s:                                    return "Isabela"
    if "CAMPO" in s and "VI" not in s and "IV" not in s: return "Campo-RM"
    return None


def norm_cat(s: str):
    s = str(s).upper().strip()
    if "DESINFECCION" in s and "FERTILIZ" in s:  return "DESINFECCION Y FERTILIZACION"
    if s.startswith("AMPLIACION"):                return "AMPLIACION"
    if "CULTIVO" in s:                            return "CULTIVO TIERRA, CHAROLAS"
    if "MATERIAL VEG" in s:                       return "MATERIAL VEGETAL"
    if "PREPARACION" in s:                        return "PREPARACION DE SUELO"
    if "FERTILIZANTE" in s:                       return "FERTILIZANTES"
    if "SANIDAD" in s or "PLAGUICIDA" in s:       return "DESINFECCION / PLAGUICIDAS"
    if "MANTENIMIENTO" in s:                      return "MANTENIMIENTO"
    if "EXPANSION" in s:                          return "EXPANSION CECILIA 25"
    if "RENOVACION" in s:                         return "RENOVACION DE SIEMBRA"
    if "MATERIAL DE EMP" in s:                    return "MATERIAL DE EMPAQUE"
    if "COSTO DE MAT" in s:                       return "COSTO_STOP"
    return None


def sv(v) -> float:
    """Convierte un valor a float de forma segura, devuelve 0.0 si falla."""
    try:
        f = float(v)
        return f if f == f else 0.0  # NaN check
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────────────────────────
# SHAREPOINT — LEER HOJAS WK
# ─────────────────────────────────────────────
def get_workbook_from_sharepoint(url: str = SHAREPOINT_URL) -> openpyxl.Workbook:
    """Descarga el Excel de SharePoint y retorna un Workbook de openpyxl."""
    log.info("Descargando Excel desde SharePoint...")
    try:
        resp = requests.get(url, timeout=60)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise RuntimeError(f"No se pudo descargar el archivo de SharePoint: {e}")
    return openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)


def sheet_to_rows(ws) -> list:
    """Convierte una hoja de openpyxl a lista de listas (igual que gspread)."""
    rows = []
    for row in ws.iter_rows(values_only=True):
        # Convertir None → "" para consistencia con gspread
        rows.append(["" if v is None else v for v in row])
    # Quitar filas completamente vacías al final
    while rows and all(v == "" for v in rows[-1]):
        rows.pop()
    return rows


def extraer_wk_desde_excel(wb: openpyxl.Workbook) -> tuple:
    """
    Lee todas las hojas WK#### del workbook de SharePoint.
    Retorna (all_data, hojas_encontradas) con el mismo formato
    que el procesador original de Google Sheets.
    """
    all_data = []
    hojas_encontradas = []

    log.info("Hojas disponibles en el Excel: %s", wb.sheetnames)

    for sname in wb.sheetnames:
        sname_clean = sname.strip()

        if sname_clean.upper() in SKIP:
            log.info("SKIP hoja: %s", sname_clean)
            continue

        wk_match = re.match(r'^WK\s*\d{4}$', sname_clean, re.IGNORECASE)
        if not wk_match:
            log.info("No es WK: %s", sname_clean)
            continue

        code_raw = re.sub(r"WK\s*", "", sname_clean, flags=re.IGNORECASE).strip()
        try:
            code = int(code_raw)
        except ValueError:
            log.warning("No se pudo convertir código WK: %s", code_raw)
            continue

        year = 2000 + (code // 100)
        if not (2018 <= year <= 2030):
            log.warning("Año fuera de rango: %d (hoja %s)", year, sname_clean)
            continue

        ww = code % 100
        log.info("Procesando hoja WK%d → Año %d Semana %d", code, year, ww)
        hojas_encontradas.append(sname_clean)

        ws   = wb[sname]
        data = sheet_to_rows(ws)
        if not data:
            continue

        # Limitar a primeras 60 filas y 35 columnas (igual que el rango A1:AI60)
        data = [row[:35] for row in data[:60]]

        # Pad filas cortas
        max_cols = max((len(r) for r in data), default=0)
        data = [r + [""] * (max_cols - len(r)) for r in data]

        date_range = ""
        if len(data) > 3 and len(data[3]) > 1:
            date_range = str(data[3][1]).strip()

        # Buscar fila "EJECUCION SEMANAL"
        exec_idx = -1
        for i, row in enumerate(data):
            if any(isinstance(c, str) and "EJECUCION SEMANAL" in c.upper() for c in row):
                exec_idx = i
                break
        if exec_idx < 0:
            log.warning("No se encontró 'EJECUCION SEMANAL' en hoja %s", sname_clean)
            continue

        # Buscar fila de encabezados (ranchos)
        header_idx = -1
        for i in range(exec_idx - 1, max(0, exec_idx - 6) - 1, -1):
            if any(isinstance(v, str) and any(k in v.upper() for k in RANCH_KEYS) for v in data[i]):
                header_idx = i
                break
        if header_idx < 0:
            log.warning("No se encontró header de ranchos en hoja %s", sname_clean)
            continue

        header = data[header_idx]

        total_cols = [j for j, v in enumerate(header)
                      if isinstance(v, str) and v.strip().upper() == "TOTAL"]
        if not total_cols:
            log.warning("No se encontró columna TOTAL en hoja %s", sname_clean)
            continue

        mxn_total_col = total_cols[0]
        usd_total_col = total_cols[1] if len(total_cols) >= 2 else None

        mxn_ranch_cols, usd_ranch_cols = {}, {}
        for j, v in enumerate(header):
            rn = norm_ranch(str(v)) if v else None
            if not rn:
                continue
            if j < mxn_total_col:
                mxn_ranch_cols[j] = rn
            elif usd_total_col and mxn_total_col < j < usd_total_col:
                mxn_ranch_cols[j] = rn
            elif usd_total_col and j > usd_total_col:
                usd_ranch_cols[j] = rn

        for i in range(exec_idx + 1, min(exec_idx + 18, len(data))):
            row   = data[i]
            label = next((str(row[c]).strip() for c in range(5)
                          if c < len(row) and row[c] and len(str(row[c]).strip()) > 3), None)
            if not label:
                continue
            cat = norm_cat(label)
            if not cat:
                continue
            if cat == "COSTO_STOP":
                break

            mxn_ranches = {rn: sv(row[j]) for j, rn in mxn_ranch_cols.items() if j < len(row)}
            usd_ranches = {rn: sv(row[j]) for j, rn in usd_ranch_cols.items() if j < len(row)}

            all_data.append({
                "semana":      code,
                "year":        year,
                "week":        ww,
                "date_range":  date_range,
                "categoria":   cat,
                "mxn_total":   round(sv(row[mxn_total_col]) if mxn_total_col < len(row) else 0, 2),
                "usd_total":   round(sv(row[usd_total_col]) if usd_total_col and usd_total_col < len(row) else 0, 2),
                "mxn_ranches": mxn_ranches,
                "usd_ranches": usd_ranches,
            })

    log.info("Hojas WK procesadas: %d | Registros extraídos: %d", len(hojas_encontradas), len(all_data))
    return all_data, hojas_encontradas


# ─────────────────────────────────────────────
# GOOGLE SHEETS — LEER HOJAS PR (sin cambios)
# ─────────────────────────────────────────────
def get_client(credentials_path: str = "credentials.json") -> gspread.Client:
    import streamlit as st
    if "gcp_service_account" in st.secrets:
        info = {
            "type":                        st.secrets["gcp_service_account"]["type"],
            "project_id":                  st.secrets["gcp_service_account"]["project_id"],
            "private_key_id":              st.secrets["gcp_service_account"]["private_key_id"],
            "private_key":                 st.secrets["gcp_service_account"]["private_key"],
            "client_email":                st.secrets["gcp_service_account"]["client_email"],
            "client_id":                   st.secrets["gcp_service_account"]["client_id"],
            "auth_uri":                    st.secrets["gcp_service_account"]["auth_uri"],
            "token_uri":                   st.secrets["gcp_service_account"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url":        st.secrets["gcp_service_account"]["client_x509_cert_url"],
        }
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    return gspread.authorize(creds)


def _parse_pr(rows: list) -> dict:
    """
    Lee filas del reporte PR#### del Google Sheet.
    Estructura:
      Col 2: UBICACION  Col 5: PRODUCTO  Col 7: UNIDADES  Col 9: GASTO
    Retorna: { rancho: { tipo: [[producto, unidades, gasto], ...] } }
    """
    RANCH_MAP = {
        'VIV': 'Prop-RM',
        'RAM': 'Campo-RM',   # ← corregido (sin espacio)
        'ISA': 'Isabela',
        'CHR': 'Christina',
        'CEC': 'Cecilia',
        'C25': 'Cecilia 25',
        'POS': 'PosCo-RM',
        'CAM': 'Campo-RM',
        'ALB': 'Albahaca-RM',
        'HOO': 'HOOPS',
    }

    UBICACION_COL = 2
    PRODUCTO_COL  = 5
    UNIDADES_COL  = 7
    GASTO_COL     = 9

    result = {}
    seen   = set()

    for row in rows:
        if not row or len(row) < 10:
            continue

        ubicacion = str(row[UBICACION_COL]).strip().upper() if len(row) > UBICACION_COL else ''
        if not ubicacion or len(ubicacion) < 6:
            continue
        if not re.match(r'^[A-Z0-9]+$', ubicacion):
            continue
        if 'MIR' not in ubicacion and 'MIP' not in ubicacion:
            continue

        ranch_code = ubicacion[:3]
        rancho = RANCH_MAP.get(ranch_code)
        if not rancho and ubicacion.startswith('VIV'):
            rancho = 'Prop-RM'
        if not rancho:
            continue

        tipo     = 'MIRFE' if 'MIR' in ubicacion else 'MIPE'
        producto = str(row[PRODUCTO_COL]).strip() if len(row) > PRODUCTO_COL else ''
        if not producto or producto.upper() in ('PRODUCTO', 'NOMBRE', ''):
            continue

        unidades = str(row[UNIDADES_COL]).strip() if len(row) > UNIDADES_COL else ''
        try:
            u = float(str(unidades).replace(',', ''))
            unidades = str(int(u)) if u == int(u) else str(round(u, 2))
        except Exception:
            unidades = '0'

        gasto = str(row[GASTO_COL]).strip() if len(row) > GASTO_COL else ''
        try:
            g = float(str(gasto).replace(',', ''))
            gasto = str(round(g, 2))
        except Exception:
            gasto = '0'

        if rancho not in result:
            result[rancho] = {}
        if tipo not in result[rancho]:
            result[rancho][tipo] = []

        if (rancho, tipo, producto) not in seen:
            seen.add((rancho, tipo, producto))
            result[rancho][tipo].append([producto, unidades, gasto])

    return result


def extraer_pr_desde_gsheets(spreadsheet: gspread.Spreadsheet) -> tuple:
    """
    Lee las hojas PR#### del Google Sheet.
    Retorna (productos, productos_debug).
    """
    pr_hojas = []

    for ws in spreadsheet.worksheets():
        sname = ws.title.strip()
        if sname.upper() in SKIP:
            continue
        pr_match = re.match(r'^PR\s*\d{4}$', sname, re.IGNORECASE)
        if not pr_match:
            continue
        pr_raw = re.sub(r'PR\s*', '', sname, flags=re.IGNORECASE).strip()
        try:
            pr_code = int(pr_raw)
            pr_year = 2000 + (pr_code // 100)
            if 2018 <= pr_year <= 2030:
                pr_hojas.append((ws.title, pr_code))
        except ValueError:
            pass

    log.info("Hojas PR encontradas en Google Sheets: %d", len(pr_hojas))

    productos       = {}
    productos_debug = {"hojas_pr_encontradas": [t for t, _ in pr_hojas]}

    if not pr_hojas:
        return productos, productos_debug

    BATCH    = 100
    pr_rangos = [f"'{t}'!A1:K500" for t, _ in pr_hojas]

    for i in range(0, len(pr_rangos), BATCH):
        grupo = pr_rangos[i:i + BATCH]
        res   = spreadsheet.values_batch_get(
            grupo, params={"valueRenderOption": "UNFORMATTED_VALUE"}
        )
        for item in res.get("valueRanges", []):
            rng  = item.get("range", "")
            vals = item.get("values", [])
            tit  = rng.split("!")[0].strip("'")
            for pt, pc in pr_hojas:
                if pt == tit:
                    parsed = _parse_pr(vals)
                    productos[pc] = parsed
                    productos_debug[f"PR{pc}_ranchos"] = list(parsed.keys()) if parsed else []
                    break

    return productos, productos_debug


# ─────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────
def get_datos(
    credentials_path:  str = "credentials.json",
    spreadsheet_name:  str = "WK 2026-08",
    sharepoint_url:    str = SHAREPOINT_URL,
) -> dict:
    """
    Orquesta la extracción desde ambas fuentes:
      • WK  → Excel de SharePoint (link público)
      • PR  → Google Sheets (service account)
    """

    # ── 1. Leer WK desde SharePoint ──────────────────
    try:
        wb = get_workbook_from_sharepoint(sharepoint_url)
        all_data, wk_hojas = extraer_wk_desde_excel(wb)
    except Exception as e:
        log.error("Error leyendo SharePoint: %s", e)
        return {"error": f"No se pudo leer el Excel de SharePoint: {e}"}

    if not all_data:
        return {"error": "No se encontraron hojas WK válidas en el Excel de SharePoint."}

    # ── 2. Leer PR desde Google Sheets ───────────────
    try:
        client = get_client(credentials_path)
        spreadsheet = None
        for name in [spreadsheet_name, spreadsheet_name.replace(" ", "_")]:
            try:
                spreadsheet = client.open(name)
                break
            except gspread.SpreadsheetNotFound:
                pass
        if not spreadsheet:
            for ss in client.openall():
                if "WK" in ss.title.upper() and "2026" in ss.title:
                    spreadsheet = ss
                    break

        if spreadsheet:
            productos, productos_debug = extraer_pr_desde_gsheets(spreadsheet)
        else:
            log.warning("Google Sheet no encontrado — datos PR no disponibles.")
            productos       = {}
            productos_debug = {"hojas_pr_encontradas": [], "error": "Sheet no encontrado"}

    except Exception as e:
        log.warning("Error leyendo Google Sheets para PR: %s", e)
        productos       = {}
        productos_debug = {"hojas_pr_encontradas": [], "error": str(e)}

    # ── 3. Construir resultado ────────────────────────
    cats_found = {r["categoria"] for r in all_data}
    cats       = [c for c in CATEGORIAS_ORDEN if c in cats_found]
    years      = sorted({r["year"] for r in all_data})

    ranches_seen: set = set()
    for r in all_data:
        ranches_seen.update(r["mxn_ranches"])
        ranches_seen.update(r["usd_ranches"])
    ranches = sorted(ranches_seen)

    summary: dict = {cat: {yr: {"usd": 0.0, "mxn": 0.0, "ranches": {}, "ranches_mxn": {}}
                            for yr in years} for cat in cats}
    for r in all_data:
        s = summary.get(r["categoria"], {}).get(r["year"])
        if not s:
            continue
        s["usd"] += r["usd_total"]
        s["mxn"] += r["mxn_total"]
        for rn, v in r["usd_ranches"].items():
            s["ranches"][rn]     = round(s["ranches"].get(rn, 0) + v, 2)
        for rn, v in r["mxn_ranches"].items():
            s["ranches_mxn"][rn] = round(s["ranches_mxn"].get(rn, 0) + v, 2)
    for cat in cats:
        for yr in years:
            d = summary[cat][yr]
            d["usd"] = round(d["usd"], 2)
            d["mxn"] = round(d["mxn"], 2)

    weeks_per_year: dict = {}
    for r in all_data:
        weeks_per_year.setdefault(r["year"], set()).add(r["week"])
    weeks_per_year = {yr: sorted(wks) for yr, wks in weeks_per_year.items()}

    return {
        "years":           years,
        "categories":      cats,
        "ranches":         ranches,
        "summary":         summary,
        "weeks_per_year":  weeks_per_year,
        "weekly_detail":   all_data,
        "productos":       productos,
        "productos_debug": productos_debug,
    }
