"""
app.py  v3 — Centro Floricultor de Baja California
st.dataframe nativo · sin dependencias extra · exporta Excel
"""

import io
import streamlit as st
import pandas as pd

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
  .block-container { padding:0.5rem 1rem 1rem !important; max-width:100% !important; }
  section[data-testid="stSidebar"] { display:none !important; }
  div[data-testid="stMetric"] {
    background:#f8fafc; border:1px solid #e2e8f0;
    border-radius:6px; padding:6px 14px;
  }
  div[data-testid="stMetricLabel"] p { font-size:0.65rem !important; color:#64748b; }
  div[data-testid="stMetricValue"]   { font-size:1.05rem !important; font-family:monospace; }
  div[data-testid="stMetricDelta"]   { font-size:0.68rem !important; }
  .stTabs [data-baseweb="tab"]       { font-size:0.78rem; font-weight:700; padding:6px 18px; }
  .stDownloadButton button { height:38px; font-size:0.75rem; font-weight:700; }
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

years_list       = DATA.get("years", [])
categories       = DATA.get("categories", [])
weekly_detail    = DATA.get("weekly_detail", [])
servicios_data   = DATA.get("servicios_data", [])
weeks_per_year   = DATA.get("weeks_per_year", {})
week_date_ranges = DATA.get("week_date_ranges", {})
productos        = DATA.get("productos", {})
productos_mp     = DATA.get("productos_mp", {})
productos_me     = DATA.get("productos_me", {})

# Lista plana de semanas
all_weeks = []
for yr in sorted(weeks_per_year.keys()):
    for wk in weeks_per_year[yr]:
        code = (yr - 2000) * 100 + wk
        dr   = week_date_ranges.get(f"{yr}-{wk}", "")
        all_weeks.append({
            "year": yr, "week": wk, "code": code, "date_range": dr,
            "label": f"WK{str(wk).zfill(2)} · {yr}" + (f"  [{dr}]" if dr else ""),
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


# ── Estilos Pandas Styler ─────────────────────────────────────────────────────
def style_table(df: pd.DataFrame, total_label: str = "▶  TOTAL"):
    def highlight_total(row):
        if str(row.iloc[0]).startswith("▶"):
            return ["background-color:#ecfdf5; font-weight:700; color:#065f46; "
                    "border-top:2px solid #059669"] * len(row)
        return [""] * len(row)

    num_cols = df.select_dtypes(include="number").columns.tolist()

    return (
        df.style
        .apply(highlight_total, axis=1)
        .format({col: lambda v: "—" if (v == 0 or pd.isna(v)) else f"${v:,.0f}"
                 for col in num_cols})
        .set_properties(**{"font-family": "monospace", "font-size": "12px"})
        .set_properties(subset=num_cols, **{"text-align": "right"})
        .set_properties(subset=[df.columns[0]], **{
            "font-weight": "600", "color": "#0f2044", "text-align": "left",
        })
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#0f2044"), ("color", "white"),
                ("font-family", "monospace"), ("font-size", "11px"),
                ("font-weight", "700"), ("text-align", "center"),
                ("padding", "6px 10px"), ("border", "none"),
            ]},
            {"selector": "tbody td", "props": [
                ("padding", "4px 10px"), ("border-bottom", "1px solid #f1f5f9"),
            ]},
            {"selector": "tbody tr:hover td", "props": [
                ("background-color", "#f0fdf4"),
            ]},
        ])
        .hide(axis="index")
    )


# ── Constructores de DataFrames ───────────────────────────────────────────────
def build_semana_df(year, week, currency="usd") -> pd.DataFrame:
    key = "usd_ranches" if currency == "usd" else "mxn_ranches"
    tot = "usd_total"   if currency == "usd" else "mxn_total"

    accum: dict = {}
    for r in weekly_detail:
        if r.get("year") != year or r.get("week") != week:
            continue
        cat = r.get("categoria", "")
        if not cat:
            continue
        accum.setdefault(cat, {rn: 0.0 for rn in RANCH_ORDER})
        accum[cat].setdefault("TOTAL", 0.0)
        for rn in RANCH_ORDER:
            accum[cat][rn] += sv(r.get(key, {}).get(rn, 0))
        accum[cat]["TOTAL"] += sv(r.get(tot, 0))

    ordered, seen = [], set()
    for cat in CATEGORIAS_ORDEN:
        if cat in accum:
            ordered.append({"Categoría": cat, **accum[cat]})
            seen.add(cat)
    for cat in accum:
        if cat not in seen:
            ordered.append({"Categoría": cat, **accum[cat]})

    if not ordered:
        return pd.DataFrame()

    df = pd.DataFrame(ordered)
    ranch_cols = [c for c in RANCH_ORDER if c in df.columns and df[c].sum() > 0]
    df = df[["Categoría"] + ranch_cols + ["TOTAL"]]

    total_row = {"Categoría": "▶  TOTAL SEMANA"}
    for col in df.columns[1:]:
        total_row[col] = df[col].sum()
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def build_comparativo_df(weeks_list, currency="usd") -> pd.DataFrame:
    tot = "usd_total" if currency == "usd" else "mxn_total"
    col_labels = [f"WK{str(w['week']).zfill(2)}-{str(w['year'])[2:]}" for w in weeks_list]

    accum: dict = {}
    for i, w in enumerate(weeks_list):
        lbl = col_labels[i]
        for r in weekly_detail:
            if r.get("year") != w["year"] or r.get("week") != w["week"]:
                continue
            cat = r.get("categoria", "")
            if not cat:
                continue
            accum.setdefault(cat, {})
            accum[cat][lbl] = accum[cat].get(lbl, 0.0) + sv(r.get(tot, 0))

    ordered, seen = [], set()
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
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def build_anual_df(year, currency="usd") -> pd.DataFrame:
    tot      = "usd_total" if currency == "usd" else "mxn_total"
    rows_raw = [r for r in weekly_detail if r.get("year") == year]
    wk_set   = sorted({r.get("week") for r in rows_raw if r.get("week")})
    col_labels = [f"W{str(wk).zfill(2)}" for wk in wk_set]

    accum: dict = {}
    for r in rows_raw:
        cat = r.get("categoria", "")
        if not cat:
            continue
        lbl = f"W{str(r.get('week')).zfill(2)}"
        accum.setdefault(cat, {})
        accum[cat][lbl] = accum[cat].get(lbl, 0.0) + sv(r.get(tot, 0))

    ordered, seen = [], set()
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
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def build_servicios_df(semana_code, currency="usd") -> pd.DataFrame:
    key = "usd_ranches" if currency == "usd" else "mxn_ranches"
    tot = "usd_total"   if currency == "usd" else "mxn_total"

    accum: dict = {}
    for r in servicios_data:
        if r.get("semana") != semana_code:
            continue
        sc = r.get("subcat", "")
        if not sc:
            continue
        accum.setdefault(sc, {rn: 0.0 for rn in RANCH_ORDER})
        accum[sc].setdefault("TOTAL", 0.0)
        for rn in RANCH_ORDER:
            accum[sc][rn] += sv(r.get(key, {}).get(rn, 0))
        accum[sc]["TOTAL"] += sv(r.get(tot, 0))

    if not accum:
        return pd.DataFrame()

    rows = [{"Subcategoría": sc, **vals} for sc, vals in accum.items()]
    df   = pd.DataFrame(rows)
    ranch_cols = [c for c in RANCH_ORDER if c in df.columns and df[c].sum() > 0]
    df   = df[["Subcategoría"] + ranch_cols + ["TOTAL"]]

    total_row = {"Subcategoría": "▶  TOTAL"}
    for col in df.columns[1:]:
        total_row[col] = df[col].sum()
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def build_productos_df(semana_code, src: dict) -> pd.DataFrame:
    prod_data = src.get(semana_code, {})
    if not prod_data:
        return pd.DataFrame()
    rows = []
    for ranch, tipos in prod_data.items():
        for tipo, items in tipos.items():
            for item in items:
                rows.append({
                    "Rancho":    ranch,
                    "Tipo":      tipo,
                    "Producto":  item[0] if len(item) > 0 else "",
                    "Unidades":  item[1] if len(item) > 1 else "",
                    "Gasto USD": sv(item[2]) if len(item) > 2 else 0.0,
                    "Ubicación": item[3] if len(item) > 3 else "",
                })
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    return df.sort_values(["Rancho", "Tipo", "Gasto USD"], ascending=[True, True, False]).reset_index(drop=True)


# ── Export Excel ──────────────────────────────────────────────────────────────
def export_excel(sel: dict, currency: str = "usd") -> bytes:
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    yr, wk, code = sel["year"], sel["week"], sel["code"]
    dfs = {
        f"WK{str(wk).zfill(2)}-{yr}": build_semana_df(yr, wk, currency),
        f"Anual-{yr}":                build_anual_df(yr, currency),
        "Servicios":                  build_servicios_df(code, currency),
        "Productos-PR":               build_productos_df(code, productos),
        "Mantenimiento-MP":           build_productos_df(code, productos_mp),
        "Mat-Empaque-ME":             build_productos_df(code, productos_me),
    }

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            if df.empty:
                continue
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]

            hdr_fill  = PatternFill("solid", fgColor="0F2044")
            hdr_font  = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
            tot_fill  = PatternFill("solid", fgColor="ECFDF5")
            tot_font  = Font(color="065F46", bold=True, size=10, name="Calibri")
            body_font = Font(size=10, name="Calibri")
            top_bdr   = Border(top=Side(style="medium", color="059669"))
            money_fmt = '$#,##0'

            num_cols_idx = {
                ci for ci, col in enumerate(df.columns, start=1)
                if pd.api.types.is_numeric_dtype(df[col])
            }
            last_row = ws.max_row

            for cell in ws[1]:
                cell.fill      = hdr_fill
                cell.font      = hdr_font
                cell.alignment = Alignment(horizontal="center")

            for ri in range(2, last_row + 1):
                is_tot = ri == last_row
                for cell in ws[ri]:
                    cell.font = tot_font if is_tot else body_font
                    if is_tot:
                        cell.fill   = tot_fill
                        cell.border = top_bdr
                    if cell.column in num_cols_idx:
                        cell.number_format = money_fmt
                        cell.alignment     = Alignment(horizontal="right")

            for col_cells in ws.columns:
                length = max(len(str(c.value or "")) for c in col_cells) + 3
                ws.column_dimensions[col_cells[0].column_letter].width = min(length, 32)

            ws.freeze_panes = "B2"

    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════
c1, c2, c3, c4 = st.columns([3, 2, 1, 1])

with c1:
    st.markdown("### 🌸 Centro Floricultor de Baja California")

with c2:
    sel_idx = st.selectbox(
        "Semana", range(len(all_weeks)),
        format_func=lambda i: all_weeks[i]["label"],
        index=len(all_weeks) - 1,
        label_visibility="collapsed",
    )
    sel_week = all_weeks[sel_idx]

with c3:
    currency = st.selectbox(
        "Moneda", ["usd", "mxn"],
        format_func=lambda x: "USD $" if x == "usd" else "MXN $",
        label_visibility="collapsed",
    )

with c4:
    xl_data = export_excel(sel_week, currency)
    st.download_button(
        "⬇ Exportar Excel", data=xl_data,
        file_name=f"CFBC_WK{str(sel_week['week']).zfill(2)}_{sel_week['year']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── KPIs ──────────────────────────────────────────────────────────────────────
df_cur   = build_semana_df(sel_week["year"], sel_week["week"], currency)
tot_cur  = float(df_cur.iloc[-1]["TOTAL"]) if not df_cur.empty else 0.0
tot_prev = 0.0

if sel_idx > 0:
    pw       = all_weeks[sel_idx - 1]
    df_prev  = build_semana_df(pw["year"], pw["week"], currency)
    tot_prev = float(df_prev.iloc[-1]["TOTAL"]) if not df_prev.empty else 0.0

delta_str = None
if tot_prev:
    pct = (tot_cur - tot_prev) / tot_prev * 100
    delta_str = f"{'+' if pct > 0 else ''}{pct:.1f}% vs semana ant."

n_cats = len(df_cur) - 1 if not df_cur.empty else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Semana",              f"WK{str(sel_week['week']).zfill(2)} · {sel_week['year']}")
k2.metric("Total semana",        f"${tot_cur:,.0f}", delta=delta_str, delta_color="inverse")
k3.metric("Categorías activas",  str(n_cats))
k4.metric("Semanas disponibles", str(len(all_weeks)))

st.divider()

# ═══════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋  Semana Actual",
    "📊  Comparativo",
    "📅  Resumen Anual",
    "🔧  Costo de Servicios",
    "📦  Productos",
])

# ── Tab 1: Semana actual ──────────────────────────────────────
with tab1:
    yr, wk = sel_week["year"], sel_week["week"]
    dr_txt = f"  |  {sel_week['date_range']}" if sel_week.get("date_range") else ""
    st.caption(f"Gasto por categoría y rancho — WK{str(wk).zfill(2)}-{yr}{dr_txt}")
    if not df_cur.empty:
        st.dataframe(style_table(df_cur), use_container_width=True,
                     height=min(50 + 30 * len(df_cur), 620))
    else:
        st.info("Sin datos para esta semana.")

# ── Tab 2: Comparativo ───────────────────────────────────────
with tab2:
    n_sel    = st.slider("Semanas a mostrar (hacia atrás)",
                          2, min(26, len(all_weeks)), 8, key="comp_n")
    start    = max(0, sel_idx - n_sel + 1)
    comp_wks = all_weeks[start:sel_idx + 1]
    df_comp  = build_comparativo_df(comp_wks, currency)

    st.caption(f"Comparativo {len(comp_wks)} semanas — hasta WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
    if not df_comp.empty:
        st.dataframe(style_table(df_comp), use_container_width=True,
                     height=min(50 + 30 * len(df_comp), 620))
    else:
        st.info("Sin datos en el rango seleccionado.")

# ── Tab 3: Anual ─────────────────────────────────────────────
with tab3:
    yr_opts = sorted(weeks_per_year.keys(), reverse=True)
    sel_yr  = st.selectbox("Año", yr_opts, format_func=str, key="annual_yr")
    df_ann  = build_anual_df(sel_yr, currency)

    n_wks_yr = len(weeks_per_year.get(sel_yr, []))
    st.caption(f"Gasto semanal por categoría — {sel_yr} — {n_wks_yr} semanas")
    if not df_ann.empty:
        st.dataframe(style_table(df_ann), use_container_width=True,
                     height=min(50 + 30 * len(df_ann), 620))
    else:
        st.info("Sin datos para este año.")

# ── Tab 4: Servicios ─────────────────────────────────────────
with tab4:
    code  = sel_week["code"]
    df_sv = build_servicios_df(code, currency)

    st.caption(f"Costo de servicios — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
    if not df_sv.empty:
        st.dataframe(style_table(df_sv), use_container_width=True,
                     height=min(50 + 35 * len(df_sv), 500))
    else:
        st.info("Sin datos de servicios para esta semana.")

    if st.checkbox("Ver comparativo de servicios", key="sv_comp"):
        n_sv   = st.slider("Semanas", 2, min(12, len(all_weeks)), 6, key="sv_n")
        sv_wks = all_weeks[max(0, sel_idx - n_sv + 1):sel_idx + 1]
        sv_lbls = [f"WK{str(w['week']).zfill(2)}-{str(w['year'])[2:]}" for w in sv_wks]
        tot_k  = "usd_total" if currency == "usd" else "mxn_total"

        sv_accum: dict = {}
        for i, w in enumerate(sv_wks):
            lbl = sv_lbls[i]
            for r in servicios_data:
                if r.get("semana") == w["code"]:
                    sc = r.get("subcat", "")
                    if sc:
                        sv_accum.setdefault(sc, {})
                        sv_accum[sc][lbl] = sv_accum[sc].get(lbl, 0.0) + sv(r.get(tot_k, 0))

        if sv_accum:
            sv_rows = [{"Subcategoría": sc, **{lbl: vals.get(lbl, 0.0) for lbl in sv_lbls}}
                       for sc, vals in sv_accum.items()]
            df_svc  = pd.DataFrame(sv_rows)
            tot_sv  = {"Subcategoría": "▶  TOTAL"}
            for lbl in sv_lbls:
                tot_sv[lbl] = df_svc[lbl].sum() if lbl in df_svc.columns else 0.0
            df_svc  = pd.concat([df_svc, pd.DataFrame([tot_sv])], ignore_index=True)
            st.dataframe(style_table(df_svc), use_container_width=True,
                         height=min(50 + 35 * len(df_svc), 400))

# ── Tab 5: Productos ─────────────────────────────────────────
with tab5:
    code = sel_week["code"]
    p1, p2, p3 = st.tabs(["Productos — PR", "Mantenimiento — MP", "Material Empaque — ME"])

    for tab_prod, src, label in [
        (p1, productos,    "Productos PR"),
        (p2, productos_mp, "Mantenimiento MP"),
        (p3, productos_me, "Material de Empaque ME"),
    ]:
        with tab_prod:
            df_p = build_productos_df(code, src)
            st.caption(f"{label} — WK{str(sel_week['week']).zfill(2)}-{sel_week['year']}")
            if not df_p.empty:
                styler = (
                    df_p.style
                    .format({"Gasto USD": lambda v: f"${v:,.0f}" if v > 0 else "—"})
                    .set_properties(**{"font-family": "monospace", "font-size": "12px"})
                    .set_properties(subset=["Gasto USD"], **{"text-align": "right"})
                    .set_table_styles([
                        {"selector": "thead th", "props": [
                            ("background-color", "#0f2044"), ("color", "white"),
                            ("font-size", "11px"), ("font-weight", "700"),
                            ("padding", "6px 10px"),
                        ]},
                        {"selector": "tbody td", "props": [
                            ("padding", "4px 10px"),
                            ("border-bottom", "1px solid #f1f5f9"),
                        ]},
                        {"selector": "tbody tr:hover td", "props": [
                            ("background-color", "#f0fdf4"),
                        ]},
                    ])
                    .hide(axis="index")
                )
                st.dataframe(styler, use_container_width=True, height=520)
            else:
                st.info("Sin datos para esta semana.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
c_r1, c_r2 = st.columns([1, 5])
with c_r1:
    if st.button("↺ Recargar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with c_r2:
    st.caption(
        "Datos desde OneDrive (WK) y Google Sheets (PR / MP / ME). "
        "Actualización automática cada 5 minutos."
    )
