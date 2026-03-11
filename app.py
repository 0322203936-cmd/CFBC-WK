"""
app.py
Centro Floricultor de Baja California
Streamlit — Ejecución Semanal Comparativo
"""

import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from data_extractor import get_datos

# ─────────────────────────────────────────────
# CONFIG DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CFBC — Ejecución Semanal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
  :root { color-scheme: dark; }
  .stApp { background:#0b1117; color:#e8f0ea; }
  section[data-testid="stSidebar"] { background:#131920; border-right:1px solid #1e3040; }
  .metric-card {
    background:#131920; border:1px solid #1e3040; border-radius:12px;
    padding:14px 18px; position:relative; overflow:hidden; margin-bottom:4px;
  }
  .metric-card::before {
    content:''; position:absolute; top:0; left:0; right:0; height:2px;
    background:linear-gradient(90deg, var(--ac,#00c97d), transparent);
  }
  .metric-yr  { font-size:.63rem; color:#5a7a66; font-family:monospace; letter-spacing:1px; }
  .metric-val { font-size:1.3rem; font-weight:800; font-family:monospace; margin:4px 0 2px; }
  .metric-sub { font-size:.68rem; font-family:monospace; }
  .delta-up   { color:#00c97d; }
  .delta-down { color:#f05252; }
  .delta-flat { color:#5a7a66; }
  .sec-title  { font-size:.68rem; text-transform:uppercase; letter-spacing:1.5px;
                color:#5a7a66; font-family:monospace; margin-bottom:6px; }
  .badge {
    display:inline-block; background:rgba(0,201,125,.08);
    border:1px solid rgba(0,201,125,.2); border-radius:20px;
    padding:3px 10px; font-size:.63rem; font-family:monospace;
    font-weight:600; color:#00c97d; margin-right:4px; margin-bottom:4px;
  }
  .badge-muted { background:rgba(58,90,72,.15); border-color:rgba(58,90,72,.3); color:#3a5a48; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
YEAR_COLORS = {2021:"#4ecdc4",2022:"#f7dc6f",2023:"#82e0aa",
               2024:"#f0b429",2025:"#00c97d",2026:"#ff6b6b"}
RANCH_COLORS = {
    "Prop-RM":"#00c97d","PosCo-RM":"#3b9eff","Campo-RM":"#f0b429",
    "Isabela":"#c084fc","HOOPS":"#fb923c","Cecilia":"#f472b6",
    "Cecilia 25":"#34d399","Christina":"#60a5fa",
    "Albahaca-RM":"#a78bfa","Campo-VI":"#94a3b8",
}
RANCH_ORDER = ["Prop-RM","PosCo-RM","Campo-RM","Isabela","HOOPS",
               "Cecilia","Cecilia 25","Christina","Albahaca-RM","Campo-VI"]
KEY_RANCHES = ["Prop-RM","PosCo-RM","Campo-RM","Isabela","Cecilia","Cecilia 25","Christina"]

PLOTLY_BASE = dict(
    paper_bgcolor="#131920", plot_bgcolor="#131920",
    font=dict(color="#5a7a66", family="IBM Plex Mono, monospace", size=11),
    xaxis=dict(gridcolor="#1e3040", linecolor="#1e3040", zerolinecolor="#1e3040"),
    yaxis=dict(gridcolor="#1e3040", linecolor="#1e3040", zerolinecolor="#1e3040"),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="#1e3040"),
    margin=dict(l=10, r=10, t=30, b=10),
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt(n) -> str:
    if not n or (isinstance(n, float) and n != n): return "—"
    return ("-$" if n < 0 else "$") + f"{abs(n):,.2f}"

def pct_delta(a, b):
    return (a - b) / b * 100 if b else None

def w_fmt(n): return f"W{int(n):02d}"

def annual_val(DATA, cat, yr, currency):
    d = DATA["summary"].get(cat, {}).get(yr, {})
    return d.get("usd" if currency == "USD" else "mxn", 0)

def get_detail(DATA, cat, week=None, year=None):
    return [r for r in DATA["weekly_detail"]
            if r["categoria"] == cat
            and (week is None or r["week"] == week)
            and (year is None or r["year"] == year)]

def agg(records):
    out = {"usd":0,"mxn":0,"ranches":{},"ranches_mxn":{},"date_range":""}
    for r in records:
        out["usd"] += r["usd_total"]; out["mxn"] += r["mxn_total"]
        if r["date_range"]: out["date_range"] = r["date_range"]
        for rn,v in r["usd_ranches"].items(): out["ranches"][rn]    = out["ranches"].get(rn,0)+v
        for rn,v in r["mxn_ranches"].items(): out["ranches_mxn"][rn]= out["ranches_mxn"].get(rn,0)+v
    out["usd"] = round(out["usd"],2); out["mxn"] = round(out["mxn"],2)
    return out

def kpi_card(yr, val, delta_pct, color):
    if delta_pct is None:
        delta_html = ""
    elif delta_pct > 0:
        delta_html = f"<div class='metric-sub delta-up'>▲ {delta_pct:.1f}% vs {yr-1}</div>"
    elif delta_pct < 0:
        delta_html = f"<div class='metric-sub delta-down'>▼ {abs(delta_pct):.1f}% vs {yr-1}</div>"
    else:
        delta_html = f"<div class='metric-sub delta-flat'>= sin cambio</div>"
    return (f"<div class='metric-card' style='--ac:{color}'>"
            f"<div class='metric-yr'>{yr}</div>"
            f"<div class='metric-val' style='color:{color}'>{fmt(val)}</div>"
            f"{delta_html}</div>")

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    st.markdown("---")

    # En Streamlit Cloud no se necesita path — las credenciales vienen de secrets
    running_in_cloud = "gcp_service_account" in st.secrets if hasattr(st, "secrets") else False

    if running_in_cloud:
        st.success("✅ Credenciales cargadas desde Secrets")
        creds_path = "credentials.json"   # no se usa en cloud, solo placeholder
    else:
        creds_path = st.text_input("Ruta al JSON de credenciales", value="credentials.json")

    sheet_name = st.text_input("Nombre del archivo en Drive", value="WK 2026-08")

    if st.button("🔄 Recargar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.markdown(
        "<div style='font-size:.63rem;color:#3a5a48;font-family:monospace'>"
        "Centro Floricultor de Baja California<br>Streamlit App v1.0</div>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load(_creds, _sheet):
    return get_datos(_creds, _sheet)

if not running_in_cloud and not os.path.exists(creds_path):
    st.warning(f"⚠️ No se encontró **{creds_path}**. Coloca el JSON de Service Account en la carpeta del proyecto.")
    st.stop()

with st.spinner("📡 Leyendo hojas desde Google Drive…"):
    DATA = load(creds_path, sheet_name)

if "error" in DATA:
    st.error(f"❌ {DATA['error']}")
    st.stop()

years     = DATA["years"]
cats      = DATA["categories"]
all_weeks = sorted({r["week"] for r in DATA["weekly_detail"]})

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown(
    "<h2 style='color:#00c97d;font-size:1.1rem;letter-spacing:-.5px;margin-bottom:2px'>"
    "📊 EJECUCIÓN SEMANAL — COMPARATIVO</h2>"
    "<div style='font-size:.72rem;color:#5a7a66;font-family:monospace;margin-bottom:14px'>"
    "Centro Floricultor de Baja California &nbsp;·&nbsp; WK 2026-08</div>",
    unsafe_allow_html=True,
)

badges = [f"{years[0]}–{years[-1]}", f"{len(years)} años",
          f"{len(cats)} categorías", f"{len(all_weeks)} semanas"]
st.markdown(
    "".join(f"<span class='badge'>{b}</span>" for b in badges) +
    "<span class='badge badge-muted'>EJECUCIÓN SEMANAL</span>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ─────────────────────────────────────────────
# CONTROLES GLOBALES
# ─────────────────────────────────────────────
c1, c2, c3 = st.columns([2, 1, 3])
with c1:
    cat = st.selectbox("📂 Categoría", cats,
                       index=cats.index("MATERIAL DE EMPAQUE") if "MATERIAL DE EMPAQUE" in cats else 0)
with c2:
    currency = st.radio("💱 Moneda", ["USD", "MXN"], horizontal=True)
with c3:
    active_years = st.multiselect("📅 Años", years, default=years)

if not active_years:
    st.warning("Selecciona al menos un año.")
    st.stop()

sym = "$" + currency

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 Por Año", "📆 Por Semana", "📈 Tendencia & Rango"])


# ══════════════════════════════════════════════════════
# TAB 1 — POR AÑO
# ══════════════════════════════════════════════════════
with tab1:

    # KPIs
    st.markdown("<div class='sec-title'>Totales Anuales</div>", unsafe_allow_html=True)
    kpi_cols = st.columns(len(active_years))
    prev = None
    for col, yr in zip(kpi_cols, active_years):
        val   = annual_val(DATA, cat, yr, currency)
        color = YEAR_COLORS.get(yr, "#888")
        delta = pct_delta(val, prev) if prev else None
        col.markdown(kpi_card(yr, val, delta, color), unsafe_allow_html=True)
        if val > 0: prev = val

    st.markdown("<br>", unsafe_allow_html=True)
    col_bar, col_ranch = st.columns(2)

    # Barra comparativa
    with col_bar:
        st.markdown("<div class='sec-title'>Comparativo Anual</div>", unsafe_allow_html=True)
        vals   = [annual_val(DATA, cat, yr, currency) for yr in active_years]
        colors = [YEAR_COLORS.get(yr, "#888") for yr in active_years]
        fig = go.Figure(go.Bar(
            x=[str(y) for y in active_years], y=vals,
            marker_color=colors,
            text=[fmt(v) for v in vals], textposition="outside",
            textfont=dict(size=10),
        ))
        fig.update_layout(**PLOTLY_BASE, height=300, showlegend=False)
        fig.update_yaxes(title_text=sym)
        st.plotly_chart(fig, use_container_width=True)

    # Desglose ranchos
    with col_ranch:
        st.markdown("<div class='sec-title'>Desglose por Rancho</div>", unsafe_allow_html=True)
        ranch_yr = st.selectbox("Ver año", ["TODOS"] + [str(y) for y in active_years],
                                label_visibility="collapsed", key="ry")
        if ranch_yr == "TODOS":
            totals: dict = {}
            for yr in active_years:
                d = DATA["summary"].get(cat,{}).get(yr,{})
                for rn,v in d.get("ranches" if currency=="USD" else "ranches_mxn",{}).items():
                    totals[rn] = totals.get(rn,0)+v
        else:
            d = DATA["summary"].get(cat,{}).get(int(ranch_yr),{})
            totals = d.get("ranches" if currency=="USD" else "ranches_mxn", {})

        ordered = [(r, totals.get(r,0)) for r in RANCH_ORDER if totals.get(r,0) > 0]
        total_sum = sum(v for _,v in ordered)
        if ordered:
            mx = max(v for _,v in ordered)
            for rn,v in ordered:
                c  = RANCH_COLORS.get(rn,"#888")
                bp = v/mx*100 if mx else 0
                pp = v/total_sum*100 if total_sum else 0
                st.markdown(
                    f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:5px'>"
                    f"<div style='width:88px;font-size:.67rem;font-family:monospace;color:{c};flex-shrink:0'>{rn}</div>"
                    f"<div style='flex:1;height:7px;background:#1a2530;border-radius:4px;overflow:hidden'>"
                    f"<div style='width:{bp:.1f}%;height:100%;background:{c};border-radius:4px'></div></div>"
                    f"<div style='width:82px;text-align:right;font-size:.66rem;font-family:monospace;color:#5a7a66'>{fmt(v)}</div>"
                    f"<div style='width:34px;text-align:right;font-size:.62rem;font-family:monospace;color:#3a5a48'>{pp:.0f}%</div>"
                    f"</div>", unsafe_allow_html=True)
        else:
            st.caption("Sin datos de ranchos para esta selección")

    # Tendencia semanal
    st.markdown("<div class='sec-title'>Tendencia Semanal (USD) — Años Superpuestos</div>", unsafe_allow_html=True)
    fig_l = go.Figure()
    for yr in active_years:
        wk = {}
        for r in get_detail(DATA, cat, year=yr):
            wk[r["week"]] = wk.get(r["week"],0)+r["usd_total"]
        xs = sorted(wk)
        if xs:
            fig_l.add_trace(go.Scatter(
                x=[w_fmt(w) for w in xs], y=[wk[w] for w in xs],
                mode="lines+markers", name=str(yr),
                line=dict(color=YEAR_COLORS.get(yr,"#888"), width=2),
                marker=dict(size=4),
            ))
    fig_l.update_layout(**PLOTLY_BASE, height=300)
    fig_l.update_yaxes(title_text="USD $")
    st.plotly_chart(fig_l, use_container_width=True)

    col_stack, col_tbl = st.columns(2)

    # Barras apiladas
    with col_stack:
        st.markdown("<div class='sec-title'>Barras Apiladas por Rancho</div>", unsafe_allow_html=True)
        fig_s = go.Figure()
        for rn in RANCH_ORDER:
            ys = [DATA["summary"].get(cat,{}).get(yr,{})
                  .get("ranches" if currency=="USD" else "ranches_mxn",{}).get(rn,0)
                  for yr in active_years]
            if any(v > 0 for v in ys):
                fig_s.add_trace(go.Bar(
                    name=rn, x=[str(y) for y in active_years], y=ys,
                    marker_color=RANCH_COLORS.get(rn,"#888"),
                ))
        fig_s.update_layout(**PLOTLY_BASE, barmode="stack", height=270)
        st.plotly_chart(fig_s, use_container_width=True)

    # Tabla resumen
    with col_tbl:
        st.markdown("<div class='sec-title'>Tabla Resumen Anual</div>", unsafe_allow_html=True)
        rows, prev = [], None
        for yr in active_years:
            v = annual_val(DATA, cat, yr, currency)
            d = pct_delta(v, prev) if prev and prev > 0 else None
            rows.append({"Año": yr, f"Total {sym}": fmt(v),
                         "Δ vs ant.": (f"+{d:.1f}%" if d and d>0 else f"{d:.1f}%" if d else "—")})
            if v > 0: prev = v
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 2 — POR SEMANA
# ══════════════════════════════════════════════════════
with tab2:
    if not all_weeks:
        st.warning("No hay semanas disponibles.")
    else:
        nav1, nav2, nav3 = st.columns([1,10,1])
        if "wi" not in st.session_state: st.session_state.wi = len(all_weeks)-1
        with nav1:
            if st.button("◀", key="pw") and st.session_state.wi > 0:
                st.session_state.wi -= 1
        with nav3:
            if st.button("▶", key="nw") and st.session_state.wi < len(all_weeks)-1:
                st.session_state.wi += 1
        with nav2:
            wn = st.select_slider("Semana", options=all_weeks,
                                  value=all_weeks[st.session_state.wi],
                                  format_func=w_fmt, label_visibility="collapsed")
            st.session_state.wi = all_weeks.index(wn)

        dr = next((r["date_range"] for r in sorted(DATA["weekly_detail"],
                   key=lambda r: r["year"], reverse=True)
                   if r["week"] == wn and r["date_range"]), "")
        avail = [yr for yr in years if any(r["week"]==wn and r["year"]==yr for r in DATA["weekly_detail"])]

        st.markdown(
            f"<div style='font-family:monospace;font-size:.85rem;color:#00c97d;font-weight:800'>Semana {w_fmt(wn)}</div>"
            f"<div style='font-family:monospace;font-size:.67rem;color:#5a7a66'>{dr}</div>"
            f"<div style='font-family:monospace;font-size:.62rem;color:#3a5a48;margin-bottom:10px'>"
            f"Disponible en: {', '.join(str(y) for y in avail)}</div>",
            unsafe_allow_html=True,
        )

        rows_sw, prev_sw = [], None
        for yr in active_years:
            recs = get_detail(DATA, cat, week=wn, year=yr)
            if not recs: continue
            a   = agg(recs)
            val = a["usd"] if currency=="USD" else a["mxn"]
            d   = pct_delta(val, prev_sw) if prev_sw and prev_sw > 0 else None
            src = a["ranches"] if currency=="USD" else a["ranches_mxn"]
            row = {"Año":yr, "Semana":w_fmt(wn), "Fecha":a["date_range"],
                   f"Total {sym}":fmt(val),
                   "Δ vs año ant.": (f"+{d:.1f}%" if d and d>0 else f"{d:.1f}%" if d else "—")}
            for rn in KEY_RANCHES:
                row[rn] = fmt(src.get(rn,0))
            rows_sw.append(row)
            if val > 0: prev_sw = val

        if rows_sw:
            st.dataframe(pd.DataFrame(rows_sw), hide_index=True, use_container_width=True)
        else:
            st.caption("Sin datos para esta semana y selección de años.")


# ══════════════════════════════════════════════════════
# TAB 3 — TENDENCIA & RANGO
# ══════════════════════════════════════════════════════
with tab3:
    if not all_weeks:
        st.warning("Sin datos.")
    else:
        rc1, rc2, rc3 = st.columns([3,3,2])
        with rc1: fw = st.select_slider("Desde", options=all_weeks, value=all_weeks[0],  format_func=w_fmt, key="fw")
        with rc2: tw = st.select_slider("Hasta", options=all_weeks, value=all_weeks[-1], format_func=w_fmt, key="tw")
        with rc3: grp = st.radio("Tabla por", ["Año → Semana","Semana → Año"], key="rg")

        rw = [w for w in all_weeks if fw <= w <= tw]
        if not rw:
            st.warning("Rango sin semanas.")
            st.stop()

        st.markdown(
            f"<div style='font-family:monospace;font-size:.72rem;color:#00c97d;margin-bottom:10px'>"
            f"{w_fmt(fw)} → {w_fmt(tw)} · {len(rw)} semanas</div>",
            unsafe_allow_html=True,
        )

        # Stats
        sc = st.columns(len(active_years))
        for col, yr in zip(sc, active_years):
            recs = [r for r in DATA["weekly_detail"]
                    if r["categoria"]==cat and r["year"]==yr and fw<=r["week"]<=tw]
            tot  = round(sum(r["usd_total"] if currency=="USD" else r["mxn_total"] for r in recs),2)
            color = YEAR_COLORS.get(yr,"#888")
            col.markdown(
                f"<div class='metric-card' style='--ac:{color}'>"
                f"<div class='metric-yr'>{yr}</div>"
                f"<div class='metric-val' style='color:{color}'>{fmt(tot)}</div>"
                f"<div class='metric-sub delta-flat'>{len(recs)} registros</div>"
                f"</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Línea tendencia
        st.markdown("<div class='sec-title'>Tendencia Semanal en el Rango — Años Superpuestos</div>", unsafe_allow_html=True)
        fig_rt = go.Figure()
        for yr in active_years:
            wkv = {}
            for r in DATA["weekly_detail"]:
                if r["categoria"]==cat and r["year"]==yr and fw<=r["week"]<=tw:
                    wkv[r["week"]] = wkv.get(r["week"],0)+(r["usd_total"] if currency=="USD" else r["mxn_total"])
            xs = sorted(wkv)
            if xs:
                fig_rt.add_trace(go.Scatter(
                    x=[w_fmt(w) for w in xs], y=[wkv[w] for w in xs],
                    mode="lines+markers", name=str(yr),
                    line=dict(color=YEAR_COLORS.get(yr,"#888"), width=2), marker=dict(size=4),
                ))
        fig_rt.update_layout(**PLOTLY_BASE, height=300)
        st.plotly_chart(fig_rt, use_container_width=True)

        col_rb, col_cu = st.columns(2)

        # Barra totales rango
        with col_rb:
            st.markdown("<div class='sec-title'>Total en el Rango por Año</div>", unsafe_allow_html=True)
            rb_v  = [round(sum(r["usd_total"] if currency=="USD" else r["mxn_total"]
                              for r in DATA["weekly_detail"]
                              if r["categoria"]==cat and r["year"]==yr and fw<=r["week"]<=tw),2)
                     for yr in active_years]
            fig_rb = go.Figure(go.Bar(
                x=[str(y) for y in active_years], y=rb_v,
                marker_color=[YEAR_COLORS.get(y,"#888") for y in active_years],
                text=[fmt(v) for v in rb_v], textposition="outside",
            ))
            fig_rb.update_layout(**PLOTLY_BASE, height=250, showlegend=False)
            st.plotly_chart(fig_rb, use_container_width=True)

        # Acumulado
        with col_cu:
            st.markdown("<div class='sec-title'>Acumulado Semanal</div>", unsafe_allow_html=True)
            fig_cu = go.Figure()
            for yr in active_years:
                wkv = {}
                for r in DATA["weekly_detail"]:
                    if r["categoria"]==cat and r["year"]==yr and fw<=r["week"]<=tw:
                        wkv[r["week"]] = wkv.get(r["week"],0)+(r["usd_total"] if currency=="USD" else r["mxn_total"])
                xs = sorted(wkv); cum = 0; ys_c = []
                for w in xs:
                    cum += wkv[w]; ys_c.append(round(cum,2))
                if xs:
                    fig_cu.add_trace(go.Scatter(
                        x=[w_fmt(w) for w in xs], y=ys_c,
                        mode="lines", name=str(yr),
                        line=dict(color=YEAR_COLORS.get(yr,"#888"), width=2),
                    ))
            fig_cu.update_layout(**PLOTLY_BASE, height=250)
            st.plotly_chart(fig_cu, use_container_width=True)

        # Heatmap
        st.markdown("<div class='sec-title'>Mapa de Calor — Semana × Año (USD)</div>", unsafe_allow_html=True)
        z, y_lbl = [], []
        gmax = 0
        mx = {yr:{w:sum(r["usd_total"] for r in get_detail(DATA,cat,w,yr)) for w in rw} for yr in active_years}
        for yr in active_years:
            gmax = max(gmax, max((mx[yr][w] for w in rw), default=0))
        for yr in active_years:
            z.append([mx[yr][w] for w in rw]); y_lbl.append(str(yr))
        fig_hm = go.Figure(go.Heatmap(
            z=z, x=[w_fmt(w) for w in rw], y=y_lbl,
            colorscale=[[0,"#0b1117"],[0.3,"#003d20"],[1,"#00c97d"]],
            hovertemplate="Año %{y} %{x}: $%{z:,.2f}<extra></extra>",
        ))
        fig_hm.update_layout(**PLOTLY_BASE, height=max(180, len(active_years)*50+80))
        st.plotly_chart(fig_hm, use_container_width=True)

        # Tabla desglose
        st.markdown("<div class='sec-title'>Tabla Desglose</div>", unsafe_allow_html=True)
        rows_r = []
        if grp == "Año → Semana":
            for yr in active_years:
                for w in rw:
                    recs = get_detail(DATA, cat, week=w, year=yr)
                    if not recs: continue
                    a = agg(recs)
                    val = a["usd"] if currency=="USD" else a["mxn"]
                    if val == 0: continue
                    src = a["ranches"] if currency=="USD" else a["ranches_mxn"]
                    row = {"Año":yr,"Semana":w_fmt(w),"Fecha":a["date_range"],f"Total {sym}":fmt(val)}
                    for rn in RANCH_ORDER: row[rn] = fmt(src.get(rn,0)) if src.get(rn,0) else "—"
                    rows_r.append(row)
        else:
            for w in rw:
                for yr in active_years:
                    recs = get_detail(DATA, cat, week=w, year=yr)
                    a    = agg(recs) if recs else {"usd":0,"mxn":0,"ranches":{},"ranches_mxn":{},"date_range":""}
                    val  = a["usd"] if currency=="USD" else a["mxn"]
                    src  = a["ranches"] if currency=="USD" else a["ranches_mxn"]
                    row  = {"Semana":w_fmt(w),"Año":yr,f"Total {sym}":fmt(val)}
                    for rn in RANCH_ORDER: row[rn] = fmt(src.get(rn,0)) if src.get(rn,0) else "—"
                    rows_r.append(row)

        if rows_r:
            st.dataframe(pd.DataFrame(rows_r), hide_index=True, use_container_width=True)
