"""
app.py  v2 — Centro Floricultor de Baja California
AG-Grid · orientado a datos · sin gráficas · exporta Excel completo
"""

import io
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, ColumnsAutoSizeMode
from st_aggrid.shared import JsCode, AgGridTheme

from data_extractor import get_datos, CATEGORIAS_ORDEN

# ── Página ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CFBC — Ejecución Semanal",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  #MainMenu, header, footer { display:none !important; }
  .block-container { padding:0.6rem 1rem 1rem !important; max-width:100% !important; }
  section[data-testid="stSidebar"] { display:none !important; }
  div[data-testid="stHorizontalBlock"] { gap:0.5rem; }
  .stTabs [data-baseweb="tab"] { font-size:0.78rem; font-weight:700; padding:6px 18px; }
  h3 { margin:0 !important; padding:0 !important; font-size:1rem !important; }
  .stDownloadButton button { height:38px; font-size:0.75rem; font-weight:700; }
  .stSelectbox label { font-size:0.7rem; }
  div[data-testid="stMetric"] { background:#f8fafc; border:1px solid #e2e8f0;
       border-radius:6px; padding:6px 14px; }
  div[data-testid="stMetricLabel"] p { font-size:0.65rem !important; }
  div[data-testid="stMetricValue"] { font-size:1rem !important; font-family:monospace; }
  div[data-testid="stMetricDelta"] { font-size:0.7rem !important; }
</style>
""", unsafe_allow_html=True)


# ── Constantes ────────────────────────────────────────────────────────────────
RANCH_ORDER = [
    "Prop-RM", "PosCo-RM", "Campo-RM", "Isabela",
    "HOOPS", "Cecilia", "Cecilia 25", "Christina",
    "Albahaca-RM", "Campo-VI",
]


# ── Carga de datos ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load():
    return get_datos()


with st.spinner("Cargando datos desde OneDrive y Google Sheets…"):
    DATA = load()

if "error" in DATA:
    st.error(DATA["error"])
    if st.button("↺ Reintentar"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# Extraer campos del DATA
years_list     = DATA.get("years", [])
categories     = DATA.get("categories", [])
weekly_detail  = DATA.get("weekly_detail", [])   # lista de dicts con usd_ranches
servicios_data = DATA.get("servicios_data", [])   # lista de dicts con usd_ranches
weeks_per_year = DATA.get("weeks_per_year", {})
week_date_ranges = DATA.get("week_date_ranges", {})
productos      = DATA.get("productos", {})        # keyed by 4-digit code int
productos_mp   = DATA.get("productos_mp", {})
productos_me   = DATA.get("productos_me", {})

# Construir lista plana de semanas disponibles
all_weeks = []
for yr in sorted(weeks_per_year.keys()):
    for wk in weeks_per_year[yr]:
        code = (yr - 2000) * 100 + wk
        dr   = week_date_ranges.get(f"{yr}-{wk}", "")
        all_weeks.append({
            "year": yr, "week": wk, "code": code,
            "date_range": dr,
            "label": f"WK{str(wk).zfill(2)} · {yr}  {('  [' + dr + ']') if dr else ''}",
        })

if not all_weeks:
    st.error("No se encontraron semanas con datos.")
    st.stop()


# ── Helpers ───────────────────────────────────────────────────────────────────
def sv(v) -> float:
    try:
        return float(v) if v else 0.0
    except (TypeError, ValueError):
        return 0.0


def fmt_pct(a, b):
    """Retorna string de variación porcentual."""
    if b == 0:
        return ""
    p = (a - b) / b * 100
    sign = "▲" if p > 0 else "▼"
    col  = "red" if p > 0 else "green"
    return f":{col}[{sign} {abs(p):.1f}%]"


# ── Funciones de construcción de DataFrames ───────────────────────────────────

def build_semana_df(year: int, week: int, currency: str = "usd") -> pd.DataFrame:
    """Categoría × Rancho para la semana dada. Última fila = TOTAL SEMANA."""
    key = "usd_ranches" if currency == "usd" else "mxn_ranches"
    tot = "usd_total"   if currency == "usd" else "mxn_total"

    rows_raw = [r for r in weekly_detail
                if r.get("year") == year and r.get("week") == week]

    # Acumular por categoría
    accum: dict = {}
    for r in rows_raw:
        cat = r.get("categoria", "")
        if not cat:
            continue
        if cat not in accum:
            accum[cat] = {rn: 0.0 for rn in RANCH_ORDER}
            accum[cat]["TOTAL"] = 0.0
        for rn in RANCH_ORDER:
            accum[cat][rn] += sv(r.get(key, {}).get(rn, 0))
        accum[cat]["TOTAL"] += sv(r.get(tot, 0))

    # Ordenar según CATEGORIAS_ORDEN
    ordered = []
    seen = set()
    for cat in CATEGORIAS_ORDEN:
        if cat in accum:
            row = {"Categoría": cat, **accum[cat]}
            ordered.append(row)
            seen.add(cat)
    for cat in accum:
        if cat not in seen:
            row = {"Categoría": cat, **accum[cat]}
            ordered.append(row)

    if not ordered:
        return pd.DataFrame()

    df = pd.DataFrame(ordered)
    # Quitar ranchos completamente vacíos
    ranch_cols = [c for c in RANCH_ORDER if c in df.columns and df[c].sum() > 0]
    df = df[["Categoría"] + ranch_cols + ["TOTAL"]]

    # Fila total
    total_row = {"Categoría": "▶  TOTAL SEMANA"}
    for col in df.columns[1:]:
        total_row[col] = df[col].sum()
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df


def build_comparativo_df(weeks_list: list, currency: str = "usd") -> pd.DataFrame:
    """Categoría × [Semana…] comparativo."""
    key = "usd_ranches" if currency == "usd" else "mxn_ranches"
    tot = "usd_total"   if currency == "usd" else "mxn_total"

    col_labels = [f"WK{str(w['week']).zfill(2)}-{str(w['year'])[2:]}" for w in weeks_list]
    accum: dict = {}

    for i, w in enumerate(weeks_list):
        lbl = col_labels[i]
        rows_raw = [r for r in weekly_detail
                    if r.get("year") == w["year"] and r.get("week") == w["week"]]
        for r in rows_raw:
            cat = r.get("categoria", "")
            if not cat:
                continue
            accum.setdefault(cat, {})
            accum[cat][lbl] = accum[cat].get(lbl, 0.0) + sv(r.get(tot, 0))

    ordered = []
    seen = set()
    for cat in CATEGORIAS_ORDEN:
        if cat in accum:
            row = {"Categoría": cat}
            for lbl in col_labels:
                row[lbl] = accum[cat].get(lbl, 0.0)
            ordered.append(row)
            seen.add(cat)

    if not ordered:
        return pd.DataFrame()

    df = pd.DataFrame(ordered)
    total_row = {"Categoría": "▶  TOTAL"}
    for lbl in col_labels:
        total_row[lbl] = df[lbl].sum() if lbl in df.columns else 0.0
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df


def build_anual_df(year: int, currency: str = "usd") -> pd.DataFrame:
    """Categoría × Semana para todo el año."""
    tot = "usd_total" if currency == "usd" else "mxn_total"
    rows_raw = [r for r in weekly_detail if r.get("year") == year]

    wk_set = sorted({r.get("week") for r in rows_raw if r.get("week")})
    col_labels = [f"W{str(wk).zfill(2)}" for wk in wk_set]

    accum: dict = {}
    for r in rows_raw:
        cat = r.get("categoria", "")
        if not cat:
            continue
        wk = r.get("week")
        lbl = f"W{str(wk).zfill(2)}"
        accum.setdefault(cat, {})
        accum[cat][lbl] = accum[cat].get(lbl, 0.0) + sv(r.get(tot, 0))

    ordered = []
    seen = set()
    for cat in CATEGORIAS_ORDEN:
        if cat in accum:
            row = {"Categoría": cat}
            total = 0.0
            for lbl in col_labels:
                v = accum[cat].get(lbl, 0.0)
                row[lbl] = v
                total += v
            row["TOTAL"] = total
            ordered.append(row)
            seen.add(cat)

    if not ordered:
        return pd.DataFrame()

    df = pd.DataFrame(ordered)
    total_row = {"Categoría": "▶  TOTAL"}
    for col in df.columns[1:]:
        total_row[col] = df[col].sum()
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df


def build_servicios_df(semana_code: int, currency: str = "usd") -> pd.DataFrame:
    """Subcategoría × Rancho para costo de servicios."""
    key = "usd_ranches" if currency == "usd" else "mxn_ranches"
    tot = "usd_total"   if currency == "usd" else "mxn_total"

    rows_raw = [r for r in servicios_data if r.get("semana") == semana_code]
    if not rows_raw:
        return pd.DataFrame()

    accum: dict = {}
    for r in rows_raw:
        sc = r.get("subcat", "")
        if not sc:
            continue
        accum.setdefault(sc, {rn: 0.0 for rn in RANCH_ORDER})
        accum[sc]["TOTAL"] = accum[sc].get("TOTAL", 0.0)
        for rn in RANCH_ORDER:
            accum[sc][rn] += sv(r.get(key, {}).get(rn, 0))
        accum[sc]["TOTAL"] += sv(r.get(tot, 0))

    rows = []
    for sc, vals in accum.items():
        row = {"Subcategoría": sc, **vals}
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    ranch_cols = [c for c in RANCH_ORDER if c in df.columns and df[c].sum() > 0]
    df = df[["Subcategoría"] + ranch_cols + ["TOTAL"]]

    total_row = {"Subcategoría": "▶  TOTAL"}
    for col in df.columns[1:]:
        total_row[col] = df[col].sum()
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df


def build_productos_df(semana_code: int, src: dict) -> pd.DataFrame:
    """Tabla de productos: Rancho, Tipo, Producto, Unidades, Gasto."""
    prod_data = src.get(semana_code, {})
    if not prod_data:
        return pd.DataFrame()
    rows = []
    for ranch, tipos in prod_data.items():
        for tipo, items in tipos.items():
            for item in items:
                producto  = item[0] if len(item) > 0 else ""
                unidades  = item[1] if len(item) > 1 else ""
                gasto     = sv(item[2]) if len(item) > 2 else 0.0
                ubicacion = item[3] if len(item) > 3 else ""
                rows.append({
                    "Rancho": ranch, "Tipo": tipo,
                    "Producto": producto, "Unidades": unidades,
                    "Gasto USD": gasto, "Ubicación": ubicacion,
                })
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df.sort_values(["Rancho", "Tipo", "Gasto USD"], ascending=[True, True, False])
    return df.reset_index(drop=True)


# ── AG-Grid helper ────────────────────────────────────────────────────────────
_money_fmt = JsCode("""
function(p) {
  var v = p.value;
  if (v === null || v === undefined || v === '' || v === 0) return '—';
  return '$' + parseFloat(v).toLocaleString('en-US', {minimumFractionDigits:0, maximumFractionDigits:0});
}
""")

_row_style = JsCode("""
function(p) {
  if (!p.api) return {};
  var last = p.api.getDisplayedRowCount() - 1;
  if (p.node.rowIndex === last) {
    return {background:'#ecfdf5', fontWeight:'700',
            borderTop:'2px solid #059669', color:'#065f46'};
  }
  return {};
}
""")

_cell_style_money = JsCode("""
function(p) {
  if (!p.api) return {textAlign:'right', fontFamily:'monospace', fontSize:'12px'};
  var last = p.api.getDisplayedRowCount() - 1;
  if (p.node.rowIndex === last)
    return {textAlign:'right', fontFamily:'monospace', fontSize:'12px',
            fontWeight:'700', color:'#065f46'};
  if (p.value > 0)
    return {textAlign:'right', fontFamily:'monospace', fontSize:'12px', color:'#1e293b'};
  return {textAlign:'right', fontFamily:'monospace', fontSize:'12px', color:'#94a3b8'};
}
""")


def show_grid(df: pd.DataFrame, key: str, height: int = 460,
              money_cols: list | None = None, first_col_width: int = 220):
    """Renderiza AG-Grid Balham con formato corporativo."""
    if df.empty:
        st.info("Sin datos para el período seleccionado.")
        return

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True, sortable=True, filter=True, minWidth=80,
        cellStyle={"fontFamily": "monospace", "fontSize": "12px"},
    )

    first_col = df.columns[0]
    gb.configure_column(
        first_col, pinned="left", width=first_col_width, minWidth=first_col_width,
        cellStyle={"fontFamily": "monospace", "fontSize": "12px",
                   "fontWeight": "600", "color": "#0f2044"},
    )

    # Auto-detectar columnas monetarias si no se especifican
    if money_cols is None:
        money_cols = [c for c in df.columns
                      if c != first_col and pd.api.types.is_numeric_dtype(df[c])]

    for col in money_cols:
        gb.configure_column(
            col, width=105, minWidth=90,
            type=["numericColumn"],
            valueFormatter=_money_fmt,
            cellStyle=_cell_style_money,
        )

    gb.configure_grid_options(
        getRowStyle=_row_style,
        suppressMovableColumns=False,
        enableBrowserTooltips=True,
        rowHeight=26,
        headerHeight=30,
    )

    AgGrid(
        df,
        gridOptions=gb.build(),
        height=height,
        theme=AgGridTheme.BALHAM,
        fit_columns_on_grid_load=False,
        columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
        update_mode=GridUpdateMode.NO_UPDATE,
        key=key,
        allow_unsafe_jscode=True,
    )


# ── Export a Excel ────────────────────────────────────────────────────────────
def export_excel(sel: dict, currency: str = "usd") -> bytes:
    """Genera workbook Excel multi-hoja."""
    yr, wk, code = sel["year"], sel["week"], sel["code"]

    dfs = {
        f"WK{str(wk).zfill(2)}-{yr}":     build_semana_df(yr, wk, currency),
        f"Anual-{yr}":                     build_anual_df(yr, currency),
        "Servicios":                       build_servicios_df(code, currency),
        "Productos-PR":                    build_productos_df(code, productos),
        "Mantenimiento-MP":               build_productos_df(code, productos_mp),
        "Mat-Empaque-ME":                  build_productos_df(code, productos_me),
    }

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#0F2044",
                                   "font_color": "#FFFFFF", "border": 1,
                                   "align": "center", "font_name": "Calibri",
                                   "font_size": 10})
        money_f  = wb.add_format({"num_format": '$#,##0', "font_name": "Calibri",
                                   "font_size": 10, "align": "right"})
        total_f  = wb.add_format({"bold": True, "bg_color": "#ECFDF5",
                                   "font_color": "#065F46", "num_format": '$#,##0',
                                   "border": 1, "top": 2, "font_name": "Calibri",
                                   "font_size": 10})
        text_f   = wb.add_format({"font_name": "Calibri", "font_size": 10})

        for sheet_name, df in dfs.items():
            if df.empty:
                continue
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            ws = writer.sheets[sheet_name]

            # Encabezados
            for ci, col_name in enumerate(df.columns):
                ws.write(0, ci, col_name, hdr_fmt)

            # Ancho de columnas y formato monetario
            num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
            last_row = len(df)
            for ci, col_name in enumerate(df.columns):
                col_w = max(len(str(col_name)) + 3, 10)
                ws.set_column(ci, ci, col_w)
                if col_name in num_cols:
                    # Fila total (última) en verde
                    ws.write(last_row, ci,
                              df[col_name].iloc[-1] if last_row > 0 else 0, total_f)
                    for ri in range(last_row - 1):
                        v = df[col_name].iloc[ri]
                        ws.write(ri + 1, ci, v, money_f)

            ws.freeze_panes(1, 1)
            ws.autofilter(0, 0, last_row, len(df.columns) - 1)

    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# UI — Header
# ═══════════════════════════════════════════════════════════════
c1, c2, c3, c4 = st.columns([3, 2, 1, 1])

with c1:
    st.markdown("### 🌸 Centro Floricultor de Baja California")

with c2:
    week_labels = [w["label"] for w in all_weeks]
    sel_idx = st.selectbox("Semana", range(len(week_labels)),
                            format_func=lambda i: week_labels[i],
                            index=len(all_weeks) - 1,
                            label_visibility="collapsed")
    sel_week = all_weeks[sel_idx]

with c3:
    currency = st.selectbox("Moneda", ["usd", "mxn"],
                              format_func=lambda x: "USD $" if x == "usd" else "MXN $",
                              label_visibility="collapsed")

with c4:
    xl_data = export_excel(sel_week, currency)
    st.download_button(
        "⬇ Exportar Excel",
        data=xl_data,
        file_name=f"CFBC_WK{str(sel_week['week']).zfill(2)}_{sel_week['year']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── KPIs rápidos ──────────────────────────────────────────────────────────────
df_cur  = build_semana_df(sel_week["year"], sel_week["week"], currency)
tot_cur = df_cur.iloc[-1]["TOTAL"] if not df_cur.empty else 0.0

tot_prev = 0.0
if sel_idx > 0:
    pw = all_weeks[sel_idx - 1]
    df_prev = build_semana_df(pw["year"], pw["week"], currency)
    tot_prev = df_prev.iloc[-1]["TOTAL"] if not df_prev.empty else 0.0

delta_pct = ((tot_cur - tot_prev) / tot_prev * 100) if tot_prev else None
delta_str = (f"{'+' if delta_pct > 0 else ''}{delta_pct:.1f}% vs semana ant."
             if delta_pct is not None else None)

n_cats = len(df_cur) - 1 if not df_cur.empty else 0  # sin fila total

k1, k2, k3, k4 = st.columns(4)
k1.metric("Semana",        f"WK{str(sel_week['week']).zfill(2)} · {sel_week['year']}")
k2.metric("Total semana",  f"${tot_cur:,.0f}", delta=delta_str,
           delta_color="inverse")
k3.metric("Categorías activas", str(n_cats))
k4.metric("Semanas disponibles", str(len(all_weeks)))

st.divider()

# ═══════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋  Semana Actual",
    "📊  Comparativo Semanas",
    "📅  Resumen Anual",
    "🔧  Costo de Servicios",
    "📦  Productos",
])

# ─── TAB 1: Semana actual ─────────────────────────────────────
with tab1:
    yr, wk = sel_week["year"], sel_week["week"]
    st.caption(
        f"Gasto por categoría y rancho — WK{str(wk).zfill(2)}-{yr}"
        + (f"  |  {sel_week['date_range']}" if sel_week.get("date_range") else "")
    )
    show_grid(df_cur, key="t1", height=520)

# ─── TAB 2: Comparativo ──────────────────────────────────────
with tab2:
    n_sel = st.slider("Semanas a mostrar (desde la semana seleccionada hacia atrás)",
                       2, min(26, len(all_weeks)), 8, key="comp_n")
    start = max(0, sel_idx - n_sel + 1)
    comp_weeks = all_weeks[start:sel_idx + 1]

    df_comp = build_comparativo_df(comp_weeks, currency)
    st.caption(
        f"Comparativo {len(comp_weeks)} semanas hasta WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}"
    )
    show_grid(df_comp, key="t2", height=520)

# ─── TAB 3: Anual ────────────────────────────────────────────
with tab3:
    yr_opts = sorted(weeks_per_year.keys(), reverse=True)
    sel_yr  = st.selectbox("Año", yr_opts, key="annual_yr",
                            format_func=str)
    df_ann  = build_anual_df(sel_yr, currency)
    n_weeks_yr = len(weeks_per_year.get(sel_yr, []))
    st.caption(f"Gasto semanal por categoría — {sel_yr} — {n_weeks_yr} semanas")
    show_grid(df_ann, key="t3", height=520)

# ─── TAB 4: Costo de Servicios ────────────────────────────────
with tab4:
    code = sel_week["code"]
    df_sv = build_servicios_df(code, currency)
    st.caption(
        f"Costo de servicios — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}"
    )
    show_grid(df_sv, key="t4", height=400, first_col_width=240)

    # Mini-comparativo de servicios
    if st.checkbox("Ver comparativo de servicios (últimas semanas)", key="sv_comp"):
        n_sv = st.slider("Semanas", 2, min(12, len(all_weeks)), 6, key="sv_n")
        sv_start = max(0, sel_idx - n_sv + 1)
        sv_weeks = all_weeks[sv_start:sel_idx + 1]

        sv_col_labels = [f"WK{str(w['week']).zfill(2)}-{str(w['year'])[2:]}" for w in sv_weeks]
        sv_accum: dict = {}
        tot_key = "usd_total" if currency == "usd" else "mxn_total"

        for i, w in enumerate(sv_weeks):
            lbl = sv_col_labels[i]
            for r in servicios_data:
                if r.get("semana") == w["code"]:
                    sc = r.get("subcat", "")
                    if sc:
                        sv_accum.setdefault(sc, {})
                        sv_accum[sc][lbl] = sv_accum[sc].get(lbl, 0.0) + sv(r.get(tot_key, 0))

        sv_rows = []
        for sc, vals in sv_accum.items():
            row = {"Subcategoría": sc}
            for lbl in sv_col_labels:
                row[lbl] = vals.get(lbl, 0.0)
            sv_rows.append(row)

        if sv_rows:
            df_sv_comp = pd.DataFrame(sv_rows)
            total_sv = {"Subcategoría": "▶  TOTAL"}
            for lbl in sv_col_labels:
                total_sv[lbl] = df_sv_comp[lbl].sum() if lbl in df_sv_comp.columns else 0.0
            df_sv_comp = pd.concat([df_sv_comp, pd.DataFrame([total_sv])], ignore_index=True)
            show_grid(df_sv_comp, key="t4c", height=300, first_col_width=240)

# ─── TAB 5: Productos ────────────────────────────────────────
with tab5:
    p1, p2, p3 = st.tabs(["Productos — PR", "Mantenimiento — MP", "Material Empaque — ME"])

    with p1:
        code = sel_week["code"]
        df_pr = build_productos_df(code, productos)
        st.caption(f"Productos PR — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
        show_grid(df_pr, key="t5a", height=500,
                  money_cols=["Gasto USD"], first_col_width=110)

    with p2:
        df_mp = build_productos_df(code, productos_mp)
        st.caption(f"Mantenimiento MP — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
        show_grid(df_mp, key="t5b", height=500,
                  money_cols=["Gasto USD"], first_col_width=110)

    with p3:
        df_me = build_productos_df(code, productos_me)
        st.caption(f"Material de Empaque ME — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
        show_grid(df_me, key="t5c", height=500,
                  money_cols=["Gasto USD"], first_col_width=110)

# ── Botón recargar al final ────────────────────────────────────────────────────
st.divider()
col_r1, col_r2 = st.columns([1, 5])
with col_r1:
    if st.button("↺ Recargar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with col_r2:
    st.caption("Datos cargados desde OneDrive (WK) y Google Sheets (PR / MP / ME). "
               "Se actualizan automáticamente cada 5 minutos.")
