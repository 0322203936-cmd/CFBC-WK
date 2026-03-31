"""
data_extractor.py
Centro Floricultor de Baja California
- Hojas WK  → Excel en SharePoint/OneDrive (pandas + requests)
- Hojas PR  → Excel separado en SharePoint  (pandas + requests)
- Hojas MP  → Excel separado en SharePoint  (MANTENIMIENTO)
- Hojas ME  → Excel separado en SharePoint  (MATERIAL DE EMPAQUE)
"""

import re
import requests
import pandas as pd
import openpyxl
from copy import copy
from io import BytesIO

# ─── URLs de SharePoint ───────────────────────────────────────────────────────
# Archivo principal: hojas WK####
SHAREPOINT_URL_WK = (
    "https://pacificafarms-my.sharepoint.com/:x:/g/personal/"
    "anahi_mora_cfbc_co/IQAQCb79SzHtRrTQR71pSNQcASOWqFXyeGGzEhUcT9FRRJ4?e=ClxLCN"
)

# Archivo secundario: hojas PR####, MP####, ME####
SHAREPOINT_URL_PR = (
    "https://pacificafarms-my.sharepoint.com/:x:/g/personal/"
    "jesus_sandoval_cfbc_co/IQCecMwUnigFQa1m-0AYEw-rAenSSKPasiHLi1p2cqtPHkc?e=wpBfv7"
)

# ─── Constantes ───────────────────────────────────────────────────────────────
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


# ─── Descarga de Excel desde SharePoint ──────────────────────────────────────
def _descargar_excel(url: str, label: str = "archivo") -> BytesIO | None:
    """
    Descarga un archivo .xlsx desde un link público de SharePoint/OneDrive.
    Agrega el parámetro download=1 necesario para la descarga directa.
    """
    download_url = url.replace("?e=", "?download=1&e=")
    try:
        response = requests.get(download_url, timeout=30)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        print(f"❌ Error descargando {label}: {e}")
        return None


# Alias para compatibilidad con get_sheet_xlsx
def descargar_excel() -> BytesIO | None:
    return _descargar_excel(SHAREPOINT_URL_WK, "Excel WK")


def _leer_hoja(xls: pd.ExcelFile, titulo: str, rango_filas: int = 60,
               rango_cols: int = 35) -> list[list]:
    """
    Lee una hoja del ExcelFile y la retorna como lista de listas.
    Las celdas vacías / NaN se convierten a "".
    """
    try:
        df = pd.read_excel(
            xls,
            sheet_name=titulo,
            header=None,
            nrows=rango_filas,
        ).fillna("")
        if df.shape[1] > rango_cols:
            df = df.iloc[:, :rango_cols]
        return df.values.tolist()
    except Exception as e:
        print(f"   ⚠️  Error leyendo hoja '{titulo}': {e}")
        return []


# ─── Helpers de normalización ─────────────────────────────────────────────────
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
    if "COSTO DE SERV" in s:                      return "COSTO_STOP"
    if s.startswith("ELECTRICIDAD"):                        return "SV:Electricidad"
    if s.startswith("FLETES Y ACARREOS"):                   return "SV:Fletes y Acarreos"
    if s.startswith("GASTOS DE EXPORTACION"):               return "SV:Gastos de Exportación"
    if s.startswith("CERTIFICADO DE FITOSANITARIO"):        return "SV:Certificado Fitosanitario"
    if s.startswith("TRANSPORTE DE PERSONAL"):              return "SV:Transporte de Personal"
    if s.startswith("COMPRA DE FLOR"):                      return "SV:Compra de Flor a Terceros"
    if s.startswith("COMIDA PARA EL PERSONAL"):             return "SV:Comida para el Personal"
    if s.startswith("RO, TEL") or s.startswith("RO , TEL"): return "SV:RO, TEL, RTA.Alim"
    return None


def sv(v) -> float:
    try:
        if isinstance(v, str):
            v = v.replace("$", "").replace(",", "").strip()
        f = float(v)
        return f if f == f else 0.0
    except (TypeError, ValueError):
        return 0.0


# ─── Parser de hojas PR#### ───────────────────────────────────────────────────
def _parse_pr(rows: list) -> dict:
    RANCH_MAP = {
        'VIV': 'Prop-RM',
        'RAM': 'Campo-RM',
        'ISA': 'Isabela',
        'CHR': 'Christina',
        'CEC': 'Cecilia',
        'C25': 'Cecilia 25',
        'POS': 'PosCo-RM',
        'CAM': 'Campo-RM',
        'ALB': 'Albahaca-RM',
        'HOO': 'HOOPS',
    }
    return _parse_generic(rows, RANCH_MAP)


# ─── Parser de hojas MP#### (MANTENIMIENTO) ───────────────────────────────────
def _parse_mp(rows: list) -> dict:
    RANCH_MAP = {
        'VIV': 'Prop-RM',
        'POS': 'PosCo-RM',
        'RAM': 'Campo-RM',
        'ISA': 'Isabela',
        'CEC': 'Cecilia',
        'C25': 'Cecilia 25',
        'CHR': 'Christina',
    }
    return _parse_generic(rows, RANCH_MAP)


# ─── Parser de hojas ME#### (MATERIAL DE EMPAQUE) ────────────────────────────
def _parse_me(rows: list) -> dict:
    RANCH_MAP = {
        'VIV': 'Prop-RM',
        'POS': 'PosCo-RM',
        'LIM': 'PosCo-RM',
        'RAM': 'Campo-RM',
        'ISA': 'Isabela',
        'CEC': 'Cecilia',
        'C25': 'Cecilia 25',
        'CHR': 'Christina',
        'ALB': 'Albahaca-RM',
        'HOO': 'HOOPS',
    }
    return _parse_generic(rows, RANCH_MAP)


# ─── Parser genérico compartido (PR / MP / ME tienen el mismo formato) ────────
def _parse_generic(rows: list, ranch_map: dict) -> dict:
    """
    Formato común a PR####, MP####, ME####:
      Col 2: UBICACION  (ej: RAMMIPRNN, CECMIPSNF)
      Col 5: PRODUCTO
      Col 7: UNIDADES
      Col 9: GASTO
    Retorna: { rancho: { tipo: [[producto, unidades, gasto, ubicacion], ...] } }
    """
    UBICACION_COL = 2
    PRODUCTO_COL  = 5
    UNIDADES_COL  = 7
    GASTO_COL     = 9

    result = {}
    accum  = {}   # (rancho, tipo, producto, ubicacion) → [u_total, g_total]

    for row in rows:
        if not row or len(row) < 10:
            continue

        ubicacion = str(row[UBICACION_COL]).strip().upper() if len(row) > UBICACION_COL else ''
        ubicacion = re.sub(r'\s+', '', ubicacion)

        if not ubicacion or len(ubicacion) < 6:
            continue
        if not re.match(r'^[A-Z0-9]+$', ubicacion):
            continue

        ranch_code = ubicacion[:3]
        rancho = ranch_map.get(ranch_code)

        if not rancho and ubicacion.startswith('VIV'):
            rancho = 'Prop-RM'

        if not rancho:
            continue

        tipo = 'MIPE' if 'MIP' in ubicacion else 'MIRFE'

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

        u_f = float(unidades) if unidades else 0.0
        g_f = float(gasto)    if gasto    else 0.0

        key = (rancho, tipo, producto, ubicacion)
        if key in accum:
            accum[key][0] += u_f
            accum[key][1] += g_f
        else:
            accum[key] = [u_f, g_f]

    for (rancho, tipo, producto, ubicacion), (u_tot, g_tot) in accum.items():
        u_str = str(int(u_tot)) if u_tot == int(u_tot) else str(round(u_tot, 2))
        g_str = str(round(g_tot, 2))
        result.setdefault(rancho, {}).setdefault(tipo, []).append([producto, u_str, g_str, ubicacion])

    return result


# ─── Fetch hojas PR / MP / ME desde el segundo Excel de SharePoint ────────────
def _fetch_desde_sharepoint(prefix: str, parser_fn, label: str) -> tuple[dict, dict]:
    """
    Descarga el Excel secundario de SharePoint y extrae todas las hojas
    que coincidan con el patrón  {PREFIX}####  (ej: PR2611, MP2608, ME2610).

    Args:
        prefix:    "PR", "MP" o "ME"
        parser_fn: función que convierte list[list] → dict de ranchos
        label:     nombre legible para logs

    Returns:
        (datos, debug)  con el mismo formato que antes usaban las funciones gspread
    """
    datos = {}
    debug = {f"hojas_{prefix.lower()}_encontradas": []}

    archivo = _descargar_excel(SHAREPOINT_URL_PR, f"Excel {label}")
    if archivo is None:
        print(f"⚠️  No se pudo descargar el archivo para hojas {prefix}")
        return datos, debug

    try:
        xls = pd.ExcelFile(archivo)
    except Exception as e:
        print(f"⚠️  No se pudo abrir el Excel de {label}: {e}")
        return datos, debug

    hojas_encontradas = []
    pat = re.compile(rf'^{prefix}\s*\d{{4}}$', re.IGNORECASE)

    for sname in xls.sheet_names:
        sname = sname.strip()
        if pat.match(sname):
            raw_code = re.sub(rf'{prefix}\s*', '', sname, flags=re.IGNORECASE).strip()
            try:
                code = int(raw_code)
                year = 2000 + (code // 100)
                if 2018 <= year <= 2030:
                    print(f"   ✅ {prefix}{code} encontrada en SharePoint: {sname}")
                    hojas_encontradas.append((sname, code))
                else:
                    print(f"   ❌ {prefix}{code} año {year} fuera de rango")
            except ValueError as e:
                print(f"   ❌ Error código '{raw_code}': {e}")

    debug[f"hojas_{prefix.lower()}_encontradas"] = [t for t, _ in hojas_encontradas]

    if not hojas_encontradas:
        print(f"   ℹ️  No hay hojas {prefix} en el Excel de SharePoint")
        return datos, debug

    for titulo, code in hojas_encontradas:
        vals   = _leer_hoja(xls, titulo, rango_filas=500, rango_cols=11)
        parsed = parser_fn(vals)
        datos[code] = parsed
        debug[f"{prefix}{code}_ranchos"] = list(parsed.keys()) if parsed else []
        print(f"   📦 {prefix}{code} ranchos detectados: {list(parsed.keys())}")

    return datos, debug


# ─── Extractor principal ──────────────────────────────────────────────────────
def extraer_datos(xls: pd.ExcelFile) -> dict:
    all_data       = []
    servicios_data = []

    hojas_validas = []
    pr_hojas      = []

    print("\n" + "=" * 60)
    print("🔍 DETECTANDO HOJAS EN EL EXCEL WK")
    print("=" * 60)

    for sname in xls.sheet_names:
        sname = sname.strip()
        print(f"\n📄 Hoja: '{sname}'")

        if sname.upper() in SKIP:
            print("   ⏭️  SKIP (en lista de exclusión)")
            continue

        pr_match = re.match(r'^PR\s*\d{4}$', sname, re.IGNORECASE)
        if pr_match:
            pr_raw = re.sub(r'PR\s*', '', sname, flags=re.IGNORECASE).strip()
            try:
                pr_code = int(pr_raw)
                pr_year = 2000 + (pr_code // 100)
                if 2018 <= pr_year <= 2030:
                    print("   ✅ PR DETECTADA Y VÁLIDA (en WK Excel)")
                    pr_hojas.append((sname, pr_code))
                    continue
            except ValueError:
                pass

        wk_match = re.match(r'^WK\s*\d{4}$', sname, re.IGNORECASE)
        if wk_match:
            code_raw = re.sub(r"WK\s*", "", sname, flags=re.IGNORECASE).strip()
            try:
                code = int(code_raw)
                year = 2000 + (code // 100)
                if 2018 <= year <= 2030:
                    print("   ✅ WK DETECTADA Y VÁLIDA")
                    hojas_validas.append((sname, code))
                else:
                    print(f"   ❌ Año {year} fuera de rango")
            except ValueError:
                print("   ❌ Error convirtiendo código")
        else:
            if not pr_match:
                print("   ℹ️  No es WK ni PR")

    print("\n" + "=" * 60)
    print("📊 RESUMEN:")
    print(f"   • Hojas WK encontradas: {len(hojas_validas)}")
    print(f"   • Hojas PR en WK Excel: {len(pr_hojas)}")
    print("=" * 60 + "\n")

    if not hojas_validas:
        return {"error": "No se encontraron hojas WK validas."}

    # 2. Leer hojas WK
    batch_data = {}
    for titulo, _ in hojas_validas:
        batch_data[titulo] = _leer_hoja(xls, titulo, rango_filas=120, rango_cols=35)

    # 2b. Leer hojas PR que estén en el Excel WK (fallback)
    productos       = {}
    productos_debug = {"hojas_pr_encontradas": [t for t, _ in pr_hojas]}
    for titulo, pr_code in pr_hojas:
        vals   = _leer_hoja(xls, titulo, rango_filas=500, rango_cols=11)
        parsed = _parse_pr(vals)
        productos[pr_code] = parsed
        productos_debug[f"PR{pr_code}_ranchos"] = list(parsed.keys()) if parsed else []

    # 3. Procesar cada hoja WK
    for titulo, code in hojas_validas:
        raw = batch_data.get(titulo, [])
        if not raw:
            continue

        yy   = code // 100
        ww   = code % 100
        year = 2000 + yy

        max_cols = max((len(r) for r in raw), default=0)
        data     = [r + [""] * (max_cols - len(r)) for r in raw]

        date_range = ""
        for _dr in range(min(8, len(data))):
            for _dc in range(min(5, len(data[_dr]))):
                _v = str(data[_dr][_dc]).strip()
                if _v and " al " in _v.lower() and len(_v) > 8:
                    date_range = _v
                    break
            if date_range:
                break

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

        print(f"\n[DEBUG {titulo}]")
        print(f"   exec_idx={exec_idx}, header_idx={header_idx}")
        print(f"   mxn_total_col={mxn_total_col}, usd_total_col={usd_total_col}")
        print(f"   mxn_ranch_cols={mxn_ranch_cols}")
        print(f"   usd_ranch_cols={usd_ranch_cols}")
        hdr_vals = [(j, str(header[j])[:15]) for j in range(len(header)) if str(header[j]).strip()]
        print(f"   header non-empty: {hdr_vals}")

        for i in range(exec_idx + 1, min(exec_idx + 120, len(data))):
            row   = data[i]
            label = next((str(row[c]).strip() for c in range(5)
                          if c < len(row) and row[c] and len(str(row[c]).strip()) > 3), None)
            if not label:
                continue

            cat = norm_cat(label)
            if not cat:
                continue
            if cat == "COSTO_STOP":
                continue

            mxn_ranches = {rn: sv(row[j]) for j, rn in mxn_ranch_cols.items() if j < len(row)}
            usd_ranches = {rn: sv(row[j]) for j, rn in usd_ranch_cols.items() if j < len(row)}

            if cat.startswith("SV:"):
                print(f"   [SV] fila={i} label='{label[:30]}' cat='{cat}' mxn_ranches={mxn_ranches}")
                servicios_data.append({
                    "semana":      code,
                    "year":        year,
                    "week":        ww,
                    "date_range":  date_range,
                    "subcat":      cat[3:],
                    "mxn_total":   round(sv(row[mxn_total_col]) if mxn_total_col < len(row) else 0, 2),
                    "usd_total":   round(sv(row[usd_total_col]) if usd_total_col and usd_total_col < len(row) else 0, 2),
                    "mxn_ranches": mxn_ranches,
                    "usd_ranches": usd_ranches,
                })
            else:
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

    print(f"\n✅ servicios_data: {len(servicios_data)} registros encontrados")
    if servicios_data:
        print(f"   subcats: {list({r['subcat'] for r in servicios_data})}")

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
    week_date_ranges: dict = {}
    for r in all_data:
        weeks_per_year.setdefault(r["year"], set()).add(r["week"])
        key = f"{r['year']}-{r['week']}"
        if key not in week_date_ranges and r.get("date_range"):
            week_date_ranges[key] = r["date_range"]
    weeks_per_year = {yr: sorted(wks) for yr, wks in weeks_per_year.items()}

    return {
        "years":            years,
        "categories":       cats,
        "ranches":          ranches,
        "summary":          summary,
        "weeks_per_year":   weeks_per_year,
        "week_date_ranges": week_date_ranges,
        "weekly_detail":    all_data,
        "productos":        productos,
        "productos_debug":  productos_debug,
        "servicios_data":   servicios_data,
    }


# ─── Punto de entrada público ─────────────────────────────────────────────────
def get_datos() -> dict:
    """
    - Hojas WK  → Excel principal en SharePoint
    - Hojas PR  → Excel secundario en SharePoint
    - Hojas MP  → Excel secundario en SharePoint (MANTENIMIENTO)
    - Hojas ME  → Excel secundario en SharePoint (MATERIAL DE EMPAQUE)
    """
    # 1. Descargar y leer Excel WK
    archivo = _descargar_excel(SHAREPOINT_URL_WK, "Excel WK")
    if archivo is None:
        return {"error": "No se pudo descargar el archivo WK de SharePoint."}

    try:
        xls = pd.ExcelFile(archivo)
    except Exception as e:
        return {"error": f"No se pudo abrir el Excel WK: {e}"}

    resultado = extraer_datos(xls)

    if "error" not in resultado:
        # 2. Leer PR desde Excel secundario de SharePoint
        print("\n" + "=" * 60)
        print("🔍 LEYENDO HOJAS PR DESDE SHAREPOINT")
        print("=" * 60)
        productos, productos_debug = _fetch_desde_sharepoint("PR", _parse_pr, "PR")
        # Merge con cualquier PR que ya estuviera en el Excel WK
        resultado["productos"].update(productos)
        resultado["productos_debug"].update(productos_debug)

        # 3. Leer MP desde Excel secundario de SharePoint (MANTENIMIENTO)
        print("\n" + "=" * 60)
        print("🔍 LEYENDO HOJAS MP DESDE SHAREPOINT (MANTENIMIENTO)")
        print("=" * 60)
        productos_mp, productos_mp_debug = _fetch_desde_sharepoint("MP", _parse_mp, "MP")
        resultado["productos_mp"]       = productos_mp
        resultado["productos_mp_debug"] = productos_mp_debug

        # 4. Leer ME desde Excel secundario de SharePoint (MATERIAL DE EMPAQUE)
        print("\n" + "=" * 60)
        print("🔍 LEYENDO HOJAS ME DESDE SHAREPOINT (MATERIAL DE EMPAQUE)")
        print("=" * 60)
        productos_me, productos_me_debug = _fetch_desde_sharepoint("ME", _parse_me, "ME")
        resultado["productos_me"]       = productos_me
        resultado["productos_me_debug"] = productos_me_debug

    return resultado


# ─── Crear nueva hoja WK en SharePoint via Microsoft Graph API ───────────────
def crear_hoja_wk(nombre_hoja: str, tenant_id: str, client_id: str, client_secret: str) -> dict:
    """
    Crea una nueva hoja en el Excel WK de SharePoint copiando el formato
    de la hoja WK más reciente como plantilla.

    Requiere una App Registration en Azure AD con permisos:
        Files.ReadWrite  (o Sites.ReadWrite.All)

    Args:
        nombre_hoja:   Nombre de la nueva hoja, ej: "WK2518"
        tenant_id:     Directory (tenant) ID de Azure AD
        client_id:     Application (client) ID
        client_secret: Client secret value

    Returns:
        {"ok": True, "mensaje": "..."} o {"ok": False, "error": "..."}
    """
    import json as _json

    # 1. Obtener token OAuth2 (client credentials flow)
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_resp = requests.post(token_url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=20)

    if token_resp.status_code != 200:
        return {"ok": False, "error": f"Error obteniendo token: {token_resp.text}"}

    token = token_resp.json().get("access_token")
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }

    # 2. Resolver el drive item ID desde el link de SharePoint
    #    Usamos el endpoint /shares para resolver un link compartido
    import base64 as _b64
    encoded = _b64.b64encode(SHAREPOINT_URL_WK.encode()).decode().rstrip("=")
    encoded = "u!" + encoded.replace("/", "_").replace("+", "-")

    resolve_url = f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem"
    res = requests.get(resolve_url, headers=headers, timeout=20)
    if res.status_code != 200:
        return {"ok": False, "error": f"No se pudo resolver el archivo SharePoint: {res.text}"}

    item = res.json()
    drive_id = item.get("parentReference", {}).get("driveId")
    item_id  = item.get("id")

    if not drive_id or not item_id:
        return {"ok": False, "error": "No se pudo obtener driveId o itemId del archivo."}

    # 3. Listar hojas existentes para encontrar la plantilla (WK más reciente)
    sheets_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets"
    )
    sheets_resp = requests.get(sheets_url, headers=headers, timeout=20)
    if sheets_resp.status_code != 200:
        return {"ok": False, "error": f"Error listando hojas: {sheets_resp.text}"}

    hojas = sheets_resp.json().get("value", [])
    nombres = [h["name"].strip() for h in hojas]

    # Verificar que la hoja no exista ya
    if nombre_hoja.upper() in [n.upper() for n in nombres]:
        return {"ok": False, "error": f"La hoja '{nombre_hoja}' ya existe en el archivo."}

    # Encontrar la hoja WK más reciente para usar como plantilla
    pat_wk = re.compile(r'^WK\s*(\d{4})$', re.IGNORECASE)
    wk_hojas = []
    for h in hojas:
        m = pat_wk.match(h["name"].strip())
        if m:
            wk_hojas.append((int(m.group(1)), h["name"].strip(), h["id"]))
    wk_hojas.sort(reverse=True)

    if not wk_hojas:
        return {"ok": False, "error": "No se encontró ninguna hoja WK para usar como plantilla."}

    plantilla_nombre = wk_hojas[0][1]

    # Base URL del workbook via Graph API
    wb_base = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook"

    # ── 4. Crear sesión persistente (necesaria para operaciones de escritura) ──
    session_resp = requests.post(
        f"{wb_base}/createSession",
        headers=headers,
        json={"persistChanges": True},
        timeout=30,
    )
    if session_resp.status_code not in (200, 201):
        return {"ok": False, "error": f"No se pudo abrir sesión en el workbook: {session_resp.text}"}

    session_id = session_resp.json().get("id", "")
    hdrs = {**headers, "workbook-session-id": session_id}

    def _close_session():
        try:
            requests.post(f"{wb_base}/closeSession", headers=hdrs, json={}, timeout=10)
        except Exception:
            pass

    import urllib.parse as _up

    # ID real de la hoja plantilla (viene del paso 3, más confiable que el nombre)
    plantilla_id = wk_hojas[0][2]

    # ── 5. Leer valores de la hoja plantilla por ID ───────────────────────────
    # El ID viene con caracteres como {, } que hay que URL-encodear
    enc_id = _up.quote(plantilla_id, safe="")
    used_range_url = f"{wb_base}/worksheets/{enc_id}/usedRange"
    range_resp = requests.get(used_range_url, headers=hdrs, timeout=30)

    # Fallback: intentar por nombre si el ID falla
    if range_resp.status_code != 200:
        enc_nombre = _up.quote(plantilla_nombre, safe="")
        used_range_url = f"{wb_base}/worksheets/{enc_nombre}/usedRange"
        range_resp = requests.get(used_range_url, headers=hdrs, timeout=30)

    if range_resp.status_code != 200:
        _close_session()
        return {
            "ok": False,
            "error": f"Error leyendo rango de plantilla '{plantilla_nombre}' (id={plantilla_id}): {range_resp.text}"
        }

    range_data   = range_resp.json()
    valores       = range_data.get("values", [])
    address       = range_data.get("address", "")  # ej: "WK2513!A1:AM55"
    rango_coords  = address.split("!")[-1] if "!" in address else address

    # ── 6. Añadir la nueva hoja vacía ────────────────────────────────────────
    add_resp = requests.post(
        f"{wb_base}/worksheets/add",
        headers=hdrs,
        json={"name": nombre_hoja},
        timeout=20,
    )
    if add_resp.status_code not in (200, 201):
        _close_session()
        return {"ok": False, "error": f"Error creando hoja nueva: {add_resp.text}"}

    nueva_id  = add_resp.json().get("id", "")
    enc_nueva = _up.quote(nueva_id, safe="")

    # ── 7. Escribir los valores en la nueva hoja por ID ──────────────────────
    patch_url = f"{wb_base}/worksheets/{enc_nueva}/range(address='{rango_coords}')"
    patch_resp = requests.patch(
        patch_url,
        headers=hdrs,
        json={"values": valores},
        timeout=30,
    )
    if patch_resp.status_code not in (200, 201):
        _close_session()
        return {"ok": False, "error": f"Error escribiendo valores en hoja nueva: {patch_resp.text}"}

    # ── 8. Cerrar sesión (confirma los cambios) ───────────────────────────────
    _close_session()

    return {
        "ok": True,
        "mensaje": (
            f"✅ Hoja '{nombre_hoja}' creada como copia de '{plantilla_nombre}' en SharePoint."
        ),
    }


# ─── Descarga de una hoja WK#### como xlsx con formato completo ───────────────
def get_sheet_xlsx(week_code: str) -> bytes | None:
    """
    Descarga el Excel de SharePoint y extrae la hoja WK{week_code}
    como un archivo .xlsx independiente con formato completo.
    """
    archivo = _descargar_excel(SHAREPOINT_URL_WK, "Excel WK")
    if archivo is None:
        return None

    archivo_bytes = archivo.getvalue()
    sheet_name = f"WK{week_code}"

    try:
        wb = openpyxl.load_workbook(BytesIO(archivo_bytes))

        target = None
        for sname in wb.sheetnames:
            normalized = re.sub(r'\s+', '', sname.strip()).upper()
            if normalized == sheet_name.upper():
                target = sname
                break

        if target is None:
            return None

        src_ws = wb[target]
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = target

        for row in src_ws.iter_rows():
            for cell in row:
                new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font          = copy(cell.font)
                    new_cell.border        = copy(cell.border)
                    new_cell.fill          = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.protection    = copy(cell.protection)
                    new_cell.alignment     = copy(cell.alignment)

        for merge in src_ws.merged_cells.ranges:
            new_ws.merge_cells(str(merge))

        for col_letter, col_dim in src_ws.column_dimensions.items():
            new_ws.column_dimensions[col_letter].width  = col_dim.width
            new_ws.column_dimensions[col_letter].hidden = col_dim.hidden
        for row_num, row_dim in src_ws.row_dimensions.items():
            new_ws.row_dimensions[row_num].height = row_dim.height
            new_ws.row_dimensions[row_num].hidden = row_dim.hidden

        buf = BytesIO()
        new_wb.save(buf)
        buf.seek(0)
        return buf.read()

    except Exception as e:
        print(f"⚠️  Error extrayendo hoja {sheet_name}: {e}")
        return None
