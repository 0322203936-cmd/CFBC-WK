"""
data_extractor.py
Centro Floricultor de Baja California
Extrae datos desde Google Sheets — lectura en batch (rapido)
"""

import re
import gspread
from google.oauth2.service_account import Credentials

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
    try:
        f = float(v)
        return f if f == f else 0.0
    except (TypeError, ValueError):
        return 0.0


def _parse_pr(rows: list) -> dict:
    """
    Lee filas del reporte CONTPAQi PR####.
    Columnas FIJAS en el reporte CONTPAQi:
      Col 0 : Codigo concepto (ej: ISAMIRPTR, CECMIPSNF)
      Col 8 : Nombre del producto
      Col 9 : Unidades
    El codigo indica rancho (prefijo) y tipo (MIR=MIRFE, MIP=MIPE).
    Retorna: { rancho: { tipo: [[nombre, unidades], ...] } }
    """
    RANCH_MAP = {
        'C25': 'Cecilia 25',
        'RAM': 'Prop-RM',
        'ISA': 'Isabela',
        'CHR': 'Christina',
        'CEC': 'Cecilia',
        'POS': 'PosCo-RM',
        'CAM': 'Campo-RM',
        'VIV': 'Vivero',
    }
    NOMBRE_COL = 8
    UNIDS_COL  = 9

    result = {}
    seen   = set()

    for row in rows:
        if not row:
            continue
        codigo = str(row[0]).strip().upper()
        # Codigo valido: 5-15 chars, solo letras y numeros, contiene MIR o MIP
        if not (5 <= len(codigo) <= 15):
            continue
        if not re.match(r'^[A-Z0-9]+$', codigo):
            continue
        if 'MIR' not in codigo and 'MIP' not in codigo:
            continue

        nombre = str(row[NOMBRE_COL]).strip() if len(row) > NOMBRE_COL else ''
        if not nombre or nombre.upper() in ('NOMBRE', 'PRODUCTO', ''):
            continue

        unids = str(row[UNIDS_COL]).strip() if len(row) > UNIDS_COL else ''
        try:
            u = float(str(unids).replace(',', ''))
            unids = str(int(u)) if u == int(u) else str(round(u, 2))
        except Exception:
            pass

        tipo = 'MIRFE' if 'MIR' in codigo else 'MIPE'

        rancho = None
        for pfx, rn in RANCH_MAP.items():
            if codigo.startswith(pfx):
                rancho = rn
                break
        if not rancho:
            continue

        if rancho not in result:
            result[rancho] = {}
        if tipo not in result[rancho]:
            result[rancho][tipo] = []
        if (rancho, tipo, nombre) not in seen:
            seen.add((rancho, tipo, nombre))
            result[rancho][tipo].append([nombre, unids])

    return result


def extraer_datos(spreadsheet: gspread.Spreadsheet) -> dict:
    all_data = []
    SKIP = {"ACUMULADO", "GRAFICOS I-IV", "COMPARATIVO", "DATOS", "HOJA1", "SHEET1"}

    # 1. Filtrar hojas validas
    hojas_validas = []
    pr_hojas = []
    for ws in spreadsheet.worksheets():
        sname = ws.title.strip()
        if sname.upper() in SKIP:
            continue
        # Detectar hojas PR#### (productos CONTPAQi)
        if re.match(r'^PR\s*\d{4}$', sname, re.IGNORECASE):
            pr_raw = re.sub(r'PR\s*', '', sname, flags=re.IGNORECASE).strip()
            try:
                pr_code = int(pr_raw)
                pr_year = 2000 + (pr_code // 100)
                if 2018 <= pr_year <= 2030:
                    pr_hojas.append((ws.title, pr_code))
                    continue
            except ValueError:
                pass
        # Hojas WK####
        code_raw = re.sub(r"WK\s*", "", sname, flags=re.IGNORECASE).strip()
        try:
            code = int(code_raw)
        except ValueError:
            continue
        year = 2000 + (code // 100)
        if not (2018 <= year <= 2030):
            continue
        hojas_validas.append((ws.title, code))

    if not hojas_validas:
        return {"error": "No se encontraron hojas WK validas."}

    # 2. Leer hojas WK en batch
    batch_data = {}
    rangos = [f"'{titulo}'!A1:AI60" for titulo, _ in hojas_validas]
    BATCH = 100
    for i in range(0, len(rangos), BATCH):
        grupo = rangos[i:i + BATCH]
        resultado = spreadsheet.values_batch_get(
            grupo,
            params={"valueRenderOption": "UNFORMATTED_VALUE"}
        )
        for item in resultado.get("valueRanges", []):
            rng    = item.get("range", "")
            vals   = item.get("values", [])
            titulo = rng.split("!")[0].strip("'")
            batch_data[titulo] = vals

    # 2b. Leer hojas PR en batch
    productos = {}
    if pr_hojas:
        pr_rangos = [f"'{t}'!A1:P300" for t, _ in pr_hojas]
        for i in range(0, len(pr_rangos), BATCH):
            grupo = pr_rangos[i:i + BATCH]
            res = spreadsheet.values_batch_get(
                grupo, params={"valueRenderOption": "UNFORMATTED_VALUE"}
            )
            for item in res.get("valueRanges", []):
                rng  = item.get("range", "")
                vals = item.get("values", [])
                tit  = rng.split("!")[0].strip("'")
                for pt, pc in pr_hojas:
                    if pt == tit:
                        productos[pc] = _parse_pr(vals)
                        break

    # 3. Procesar cada hoja WK
    for titulo, code in hojas_validas:
        raw = batch_data.get(titulo, [])
        if not raw:
            continue

        yy   = code // 100
        ww   = code % 100
        year = 2000 + yy

        max_cols = max((len(r) for r in raw), default=0)
        data = [r + [""] * (max_cols - len(r)) for r in raw]

        date_range = ""
        if len(data) > 3 and len(data[3]) > 1:
            date_range = str(data[3][1]).strip()

        exec_idx = -1
        for i, row in enumerate(data):
            if any(isinstance(c, str) and "EJECUCION SEMANAL" in c.upper() for c in row):
                exec_idx = i
                break
        if exec_idx < 0:
            continue

        header_idx = -1
        for i in range(exec_idx - 1, max(0, exec_idx - 6) - 1, -1):
            if any(isinstance(v, str) and any(k in v.upper() for k in RANCH_KEYS) for v in data[i]):
                header_idx = i
                break
        if header_idx < 0:
            continue

        header = data[header_idx]

        total_cols = [j for j, v in enumerate(header)
                      if isinstance(v, str) and v.strip().upper() == "TOTAL"]
        if not total_cols:
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

    cats_found = {r["categoria"] for r in all_data}
    cats  = [c for c in CATEGORIAS_ORDEN if c in cats_found]
    years = sorted({r["year"] for r in all_data})

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
            s["ranches"][rn] = round(s["ranches"].get(rn, 0) + v, 2)
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
        "years":          years,
        "categories":     cats,
        "ranches":        ranches,
        "summary":        summary,
        "weeks_per_year": weeks_per_year,
        "weekly_detail":  all_data,
        "productos":      productos,
    }


def get_datos(credentials_path: str = "credentials.json",
              spreadsheet_name: str = "WK 2026-08") -> dict:
    client = get_client(credentials_path)

    for name in [spreadsheet_name, spreadsheet_name.replace(" ", "_")]:
        try:
            return extraer_datos(client.open(name))
        except gspread.SpreadsheetNotFound:
            pass

    for ss in client.openall():
        if "WK" in ss.title.upper() and "2026" in ss.title:
            return extraer_datos(ss)

    return {"error": f"No se encontro '{spreadsheet_name}' en Drive. Verifica que lo compartiste con la service account."}
