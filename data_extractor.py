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
    Lee filas del reporte PR#### del Excel.
    Estructura del archivo PR:
      Col 2: UBICACION (código del rancho, ej: RAMMIPRNN, CECMIPSNF)
      Col 5: PRODUCTO (nombre del químico/producto)
      Col 7: UNIDADES
      Col 9: GASTO
    
    El código UBICACION indica:
      - Primeros 3 chars: rancho (RAM, CEC, ISA, CHR, POS, CAM, etc)
      - Chars siguientes: tipo MIR=MIRFE o MIP=MIPE
    
    Retorna: { rancho: { tipo: [[producto, unidades, gasto], ...] } }
    """
    RANCH_MAP = {
        'RAM': 'Prop-RM',
        'C25': 'Cecilia 25',
        'ISA': 'Isabela',
        'CHR': 'Christina',
        'CEC': 'Cecilia',
        'POS': 'PosCo-RM',
        'CAM': 'Campo-RM',
        'VIV': 'Vivero',
        'ALB': 'Albahaca-RM',
        'HOO': 'HOOPS',
    }
    
    UBICACION_COL = 2
    PRODUCTO_COL = 5
    UNIDADES_COL = 7
    GASTO_COL = 9

    result = {}
    seen = set()

    for row in rows:
        if not row or len(row) < 10:
            continue
        
        # Leer UBICACION (código del rancho)
        ubicacion = str(row[UBICACION_COL]).strip().upper() if len(row) > UBICACION_COL else ''
        
        # Validar que sea un código válido
        if not ubicacion or len(ubicacion) < 6:
            continue
        if not re.match(r'^[A-Z0-9]+$', ubicacion):
            continue
        if 'MIR' not in ubicacion and 'MIP' not in ubicacion:
            continue
        
        # Extraer rancho (primeros 3 caracteres)
        ranch_code = ubicacion[:3]
        rancho = RANCH_MAP.get(ranch_code)
        if not rancho:
            continue
        
        # Determinar tipo (MIRFE o MIPE)
        tipo = 'MIRFE' if 'MIR' in ubicacion else 'MIPE'
        
        # Leer producto
        producto = str(row[PRODUCTO_COL]).strip() if len(row) > PRODUCTO_COL else ''
        if not producto or producto.upper() in ('PRODUCTO', 'NOMBRE', ''):
            continue
        
        # Leer unidades
        unidades = str(row[UNIDADES_COL]).strip() if len(row) > UNIDADES_COL else ''
        try:
            u = float(str(unidades).replace(',', ''))
            unidades = str(int(u)) if u == int(u) else str(round(u, 2))
        except Exception:
            unidades = '0'
        
        # Leer gasto
        gasto = str(row[GASTO_COL]).strip() if len(row) > GASTO_COL else ''
        try:
            g = float(str(gasto).replace(',', ''))
            gasto = str(round(g, 2))
        except Exception:
            gasto = '0'
        
        # Agregar al resultado
        if rancho not in result:
            result[rancho] = {}
        if tipo not in result[rancho]:
            result[rancho][tipo] = []
        
        # Evitar duplicados
        if (rancho, tipo, producto) not in seen:
            seen.add((rancho, tipo, producto))
            result[rancho][tipo].append([producto, unidades, gasto])

    return result


def extraer_datos(spreadsheet: gspread.Spreadsheet) -> dict:
    all_data = []
    SKIP = {"ACUMULADO", "GRAFICOS I-IV", "COMPARATIVO", "DATOS", "HOJA1", "SHEET1"}

    # 1. Filtrar hojas validas
    hojas_validas = []
    pr_hojas = []
    print("\n" + "=" * 60)
    print("🔍 DETECTANDO HOJAS EN GOOGLE SHEET")
    print("=" * 60)
    
    for ws in spreadsheet.worksheets():
        sname = ws.title.strip()
        print(f"\n📄 Hoja: '{sname}'")
        
        if sname.upper() in SKIP:
            print(f"   ⏭️  SKIP (en lista de exclusión)")
            continue
            
        # Detectar hojas PR#### (productos CONTPAQi)
        pr_match = re.match(r'^PR\s*\d{4}$', sname, re.IGNORECASE)
        print(f"   🔍 Regex PR: {bool(pr_match)}")
        
        if pr_match:
            pr_raw = re.sub(r'PR\s*', '', sname, flags=re.IGNORECASE).strip()
            print(f"   📊 Código extraído: '{pr_raw}'")
            try:
                pr_code = int(pr_raw)
                pr_year = 2000 + (pr_code // 100)
                print(f"   📅 PR{pr_code} → Año {pr_year}, Semana {pr_code % 100}")
                if 2018 <= pr_year <= 2030:
                    print(f"   ✅ PR DETECTADA Y VÁLIDA")
                    pr_hojas.append((ws.title, pr_code))
                    continue
                else:
                    print(f"   ❌ Año {pr_year} fuera de rango (2018-2030)")
            except ValueError as e:
                print(f"   ❌ Error: {e}")
                pass
                
        # Hojas WK####
        wk_match = re.match(r'^WK\s*\d{4}$', sname, re.IGNORECASE)
        print(f"   🔍 Regex WK: {bool(wk_match)}")
        
        if wk_match:
            code_raw = re.sub(r"WK\s*", "", sname, flags=re.IGNORECASE).strip()
            try:
                code = int(code_raw)
                year = 2000 + (code // 100)
                print(f"   📅 WK{code} → Año {year}, Semana {code % 100}")
                if 2018 <= year <= 2030:
                    print(f"   ✅ WK DETECTADA Y VÁLIDA")
                    hojas_validas.append((ws.title, code))
                else:
                    print(f"   ❌ Año {year} fuera de rango")
            except ValueError:
                print(f"   ❌ Error convirtiendo código")
                continue
        else:
            print(f"   ℹ️  No es WK ni PR")
    
    print("\n" + "=" * 60)
    print(f"📊 RESUMEN:")
    print(f"   • Hojas WK encontradas: {len(hojas_validas)}")
    print(f"   • Hojas PR encontradas: {len(pr_hojas)}")
    print("=" * 60 + "\n")

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
    productos_debug = {"hojas_pr_encontradas": [t for t, _ in pr_hojas]}
    
    if pr_hojas:
        pr_rangos = [f"'{t}'!A1:K500" for t, _ in pr_hojas]
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
                        parsed = _parse_pr(vals)
                        productos[pc] = parsed
                        # Debug: guardar info sobre qué se encontró
                        ranchos_en_parsed = list(parsed.keys()) if parsed else []
                        productos_debug[f"PR{pc}_ranchos"] = ranchos_en_parsed
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
        "productos_debug": productos_debug,
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
