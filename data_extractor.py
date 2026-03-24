# ═══════════════════════════════════════════════════════════════════════════════
# PATCH — Agregar al final de data_extractor.py (antes de get_datos)
# Maneja hojas MP#### para la categoría MANTENIMIENTO
# ═══════════════════════════════════════════════════════════════════════════════


# ─── Parser de hojas MP#### ───────────────────────────────────────────────────
def _parse_mp(rows: list) -> dict:
    """
    Lee filas del reporte MP#### del Google Sheet.
    Mismo formato que PR####:
      Col 2: UBICACION  (ej: VIVEVIV, POSCOMIP, RAMMIR...)
      Col 5: PRODUCTO
      Col 7: UNIDADES
      Col 9: GASTO
    Retorna: { rancho: { tipo: [[producto, unidades, gasto], ...] } }

    Ranchos detectados para MANTENIMIENTO:
      VIV → Prop-RM
      POS → PosCo-RM
      RAM → Campo-RM
      ISA → Isabela
      CEC → Cecilia
      C25 → Cecilia 25
      CHR → Christina
    """
    RANCH_MAP = {
        'VIV': 'Prop-RM',
        'POS': 'PosCo-RM',
        'RAM': 'Campo-RM',
        'ISA': 'Isabela',
        'CEC': 'Cecilia',
        'C25': 'Cecilia 25',
        'CHR': 'Christina',
    }

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
        rancho = RANCH_MAP.get(ranch_code)

        # VIVEVIV y variantes → Prop-RM
        if not rancho and ubicacion.startswith('VIV'):
            rancho = 'Prop-RM'

        if not rancho:
            continue

        # Determinar tipo por código de ubicación
        if 'MIP' in ubicacion:
            tipo = 'MIPE'
        else:
            tipo = 'MIRFE'

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

        u = float(unidades) if unidades else 0.0
        g = float(gasto)    if gasto    else 0.0

        key = (rancho, tipo, producto, ubicacion)
        if key in accum:
            accum[key][0] += u
            accum[key][1] += g
        else:
            accum[key] = [u, g]

    # Construir resultado agrupado igual que _parse_pr
    for (rancho, tipo, producto, ubicacion), (u_total, g_total) in accum.items():
        result.setdefault(rancho, {}).setdefault(tipo, []).append([
            producto,
            str(int(u_total)) if u_total == int(u_total) else str(round(u_total, 2)),
            str(round(g_total, 2)),
            ubicacion,
        ])

    return result


# ─── Fetch hojas MP desde Google Sheets ──────────────────────────────────────
def _fetch_mp_desde_sheets(spreadsheet_name: str = "WK 2026-08") -> tuple[dict, dict]:
    """
    Se conecta a Google Sheets y lee SOLO las hojas MP####.
    Retorna (productos_mp, productos_mp_debug) con el mismo formato que PR.
    """
    productos_mp       = {}
    productos_mp_debug = {"hojas_mp_encontradas": []}

    try:
        client = get_gsheets_client()

        ss = None
        for name in [spreadsheet_name, spreadsheet_name.replace(" ", "_")]:
            try:
                ss = client.open(name)
                break
            except gspread.SpreadsheetNotFound:
                pass

        if ss is None:
            for s in client.openall():
                if "WK" in s.title.upper() and "2026" in s.title:
                    ss = s
                    break

        if ss is None:
            print("⚠️  No se encontró el Google Sheet para hojas MP")
            return productos_mp, productos_mp_debug

        # Filtrar solo hojas MP####
        mp_hojas = []
        for ws in ss.worksheets():
            sname = ws.title.strip()
            mp_match = re.match(r'^MP\s*\d{4}$', sname, re.IGNORECASE)
            if mp_match:
                mp_raw = re.sub(r'MP\s*', '', sname, flags=re.IGNORECASE).strip()
                try:
                    mp_code = int(mp_raw)
                    mp_year = 2000 + (mp_code // 100)
                    if 2018 <= mp_year <= 2030:
                        print(f"   ✅ MP encontrada en Sheets: {sname}")
                        mp_hojas.append((ws.title, mp_code))
                except ValueError:
                    pass

        productos_mp_debug["hojas_mp_encontradas"] = [t for t, _ in mp_hojas]

        if not mp_hojas:
            print("   ℹ️  No hay hojas MP en el Google Sheet")
            return productos_mp, productos_mp_debug

        # Leer hojas MP en batch
        BATCH = 100
        mp_rangos = [f"'{t}'!A1:K500" for t, _ in mp_hojas]
        for i in range(0, len(mp_rangos), BATCH):
            grupo = mp_rangos[i:i + BATCH]
            res   = ss.values_batch_get(grupo, params={"valueRenderOption": "UNFORMATTED_VALUE"})
            for item in res.get("valueRanges", []):
                rng  = item.get("range", "")
                vals = item.get("values", [])
                tit  = rng.split("!")[0].strip("'")
                for pt, pc in mp_hojas:
                    if pt == tit:
                        parsed = _parse_mp(vals)
                        productos_mp[pc] = parsed
                        productos_mp_debug[f"MP{pc}_ranchos"] = list(parsed.keys()) if parsed else []
                        break

    except Exception as e:
        print(f"⚠️  Error leyendo MP desde Google Sheets: {e}")

    return productos_mp, productos_mp_debug


# ═══════════════════════════════════════════════════════════════════════════════
# MODIFICACIÓN a get_datos — reemplazar el bloque final con este:
# ═══════════════════════════════════════════════════════════════════════════════

def get_datos(spreadsheet_name: str = "WK 2026-08") -> dict:
    """
    - Hojas WK  → descargadas desde el Excel de OneDrive
    - Hojas PR  → leídas desde Google Sheets (productos generales)
    - Hojas MP  → leídas desde Google Sheets (MANTENIMIENTO)
    """
    # 1. Leer WK desde OneDrive
    archivo = descargar_excel()
    if archivo is None:
        return {"error": "No se pudo descargar el archivo de OneDrive."}

    try:
        xls = pd.ExcelFile(archivo)
    except Exception as e:
        return {"error": f"No se pudo abrir el Excel: {e}"}

    resultado = extraer_datos(xls)

    if "error" not in resultado:
        # 2. Leer PR desde Google Sheets
        print("\n" + "=" * 60)
        print("🔍 LEYENDO HOJAS PR DESDE GOOGLE SHEETS")
        print("=" * 60)
        productos, productos_debug = _fetch_pr_desde_sheets(spreadsheet_name)
        resultado["productos"]       = productos
        resultado["productos_debug"] = productos_debug

        # 3. Leer MP desde Google Sheets (MANTENIMIENTO)
        print("\n" + "=" * 60)
        print("🔍 LEYENDO HOJAS MP DESDE GOOGLE SHEETS (MANTENIMIENTO)")
        print("=" * 60)
        productos_mp, productos_mp_debug = _fetch_mp_desde_sheets(spreadsheet_name)
        resultado["productos_mp"]       = productos_mp
        resultado["productos_mp_debug"] = productos_mp_debug

    return resultado
