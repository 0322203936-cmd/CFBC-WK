"""
app.py
Centro Floricultor de Baja California
Vista empresarial enfocada en datos, sin graficas.
"""

from __future__ import annotations

import importlib
import importlib.util
from pathlib import Path

import pandas as pd
import streamlit as st


def _resolve_get_datos():
    """Importa get_datos aunque el archivo tenga sufijo '(n)'."""
    try:
        module = importlib.import_module("data_extractor")
        if hasattr(module, "get_datos"):
            return module.get_datos
    except Exception:
        pass

    base_dir = Path(__file__).resolve().parent
    candidates = sorted(
        base_dir.glob("data_extractor*.py"),
        key=lambda p: (p.name != "data_extractor.py", p.name),
    )

    for candidate in candidates:
        spec = importlib.util.spec_from_file_location("data_extractor_dynamic", candidate)
        if not spec or not spec.loader:
            continue
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        if hasattr(module, "get_datos"):
            return module.get_datos

    raise ImportError("No se encontro get_datos en ningun data_extractor*.py")


get_datos = _resolve_get_datos()

st.set_page_config(
    page_title="CFBC Data Console",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
  #MainMenu, header, footer { display: none !important; }
  .block-container {
    max-width: 100% !important;
    padding-top: 0.6rem !important;
    padding-bottom: 0.5rem !important;
    padding-left: 0.9rem !important;
    padding-right: 0.9rem !important;
  }
  .stMetric {
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.5rem 0.7rem;
    background: #ffffff;
  }
  div[data-testid="stHorizontalBlock"] > div {
    gap: 0.55rem;
  }
  .stTabs [data-baseweb="tab-list"] {
    gap: 0.15rem;
  }
  .stTabs [data-baseweb="tab"] {
    height: 2.1rem;
    padding-top: 0.1rem;
    padding-bottom: 0.1rem;
  }
</style>
""",
    unsafe_allow_html=True,
)


def _to_float(value) -> float:
    try:
        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").strip()
        return float(value)
    except Exception:
        return 0.0


@st.cache_data(ttl=300, show_spinner=False)
def load_data() -> dict:
    return get_datos()


def _build_weekly_df(data: dict) -> pd.DataFrame:
    rows = []
    for item in data.get("weekly_detail", []):
        rows.append(
            {
                "year": int(item.get("year", 0)),
                "week": int(item.get("week", 0)),
                "categoria": str(item.get("categoria", "")),
                "date_range": str(item.get("date_range", "")),
                "usd_total": _to_float(item.get("usd_total", 0)),
                "mxn_total": _to_float(item.get("mxn_total", 0)),
                "usd_ranches": item.get("usd_ranches", {}) or {},
                "mxn_ranches": item.get("mxn_ranches", {}) or {},
            }
        )
    return pd.DataFrame(rows)


def _build_servicios_df(data: dict) -> pd.DataFrame:
    rows = []
    for item in data.get("servicios_data", []):
        rows.append(
            {
                "year": int(item.get("year", 0)),
                "week": int(item.get("week", 0)),
                "subcat": str(item.get("subcat", "")),
                "date_range": str(item.get("date_range", "")),
                "usd_total": _to_float(item.get("usd_total", 0)),
                "mxn_total": _to_float(item.get("mxn_total", 0)),
                "usd_ranches": item.get("usd_ranches", {}) or {},
                "mxn_ranches": item.get("mxn_ranches", {}) or {},
            }
        )
    return pd.DataFrame(rows)


def _flatten_productos(productos: dict, source: str) -> pd.DataFrame:
    rows = []
    for code_raw, ranch_map in (productos or {}).items():
        code = int(code_raw)
        year = 2000 + (code // 100)
        week = code % 100

        for ranch, tipos in (ranch_map or {}).items():
            for tipo, items in (tipos or {}).items():
                for row in (items or []):
                    producto = str(row[0]) if len(row) > 0 else ""
                    unidades = _to_float(row[1]) if len(row) > 1 else 0.0
                    gasto = _to_float(row[2]) if len(row) > 2 else 0.0
                    ubicacion = str(row[3]) if len(row) > 3 else ""
                    rows.append(
                        {
                            "source": source,
                            "year": year,
                            "week": week,
                            "rancho": ranch,
                            "tipo": tipo,
                            "producto": producto,
                            "unidades": unidades,
                            "gasto": gasto,
                            "ubicacion": ubicacion,
                        }
                    )

    return pd.DataFrame(rows)


@st.cache_data(ttl=300, show_spinner=False)
def _build_productos_df(data: dict) -> pd.DataFrame:
    frames = [
        _flatten_productos(data.get("productos", {}), "PR"),
        _flatten_productos(data.get("productos_mp", {}), "MP"),
        _flatten_productos(data.get("productos_me", {}), "ME"),
    ]
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame(
            columns=["source", "year", "week", "rancho", "tipo", "producto", "unidades", "gasto", "ubicacion"]
        )
    return pd.concat(frames, ignore_index=True)


def _value_from_row(row: pd.Series, currency: str, ranch: str) -> float:
    if ranch == "Todos":
        return float(row["usd_total"] if currency == "USD" else row["mxn_total"])

    key = "usd_ranches" if currency == "USD" else "mxn_ranches"
    return _to_float((row.get(key) or {}).get(ranch, 0))


def _apply_value(df: pd.DataFrame, currency: str, ranch: str) -> pd.DataFrame:
    out = df.copy()
    out["valor"] = out.apply(lambda r: _value_from_row(r, currency, ranch), axis=1)
    return out


def _currency_symbol(currency: str) -> str:
    return "$" if currency == "USD" else "MXN$"


def _to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


try:
    DATA = load_data()
except Exception as exc:
    st.error(f"Error cargando datos: {exc}")
    st.stop()

if "error" in DATA:
    st.error(DATA["error"])
    if st.button("Reintentar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.stop()

weekly_df = _build_weekly_df(DATA)
serv_df = _build_servicios_df(DATA)
prod_df = _build_productos_df(DATA)

if weekly_df.empty:
    st.warning("No hay datos en weekly_detail para construir la vista.")
    st.stop()

all_years = sorted(int(y) for y in weekly_df["year"].dropna().unique().tolist())
all_categories = sorted(c for c in weekly_df["categoria"].dropna().unique().tolist() if c)
all_ranches = ["Todos"] + sorted(DATA.get("ranches", []))

latest_year = max(all_years)

top_left, top_mid, top_right = st.columns([1.2, 2.8, 1.1])
top_left.markdown("### CFBC Data Console")
top_mid.caption(
    "Vista empresarial orientada a operacion: sin graficas, sin espacios vacios, con foco en datos accionables."
)
if top_right.button("Recargar datos", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

f1, f2, f3, f4, f5 = st.columns([1.0, 1.0, 1.5, 1.25, 2.0])
currency = f1.radio("Moneda", ["USD", "MXN"], horizontal=True)
selected_year = f2.selectbox("Anio", options=all_years, index=all_years.index(latest_year))
selected_ranch = f3.selectbox("Rancho", options=all_ranches)

weeks_in_year = sorted(weekly_df.loc[weekly_df["year"] == selected_year, "week"].unique().tolist())
week_min = int(min(weeks_in_year))
week_max = int(max(weeks_in_year))
week_range = f4.slider("Semanas", min_value=week_min, max_value=week_max, value=(week_min, week_max))

default_cats = all_categories[:]
selected_categories = f5.multiselect("Categorias", options=all_categories, default=default_cats)

if not selected_categories:
    st.warning("Selecciona al menos una categoria.")
    st.stop()

mask = (
    (weekly_df["year"] == selected_year)
    & (weekly_df["week"] >= week_range[0])
    & (weekly_df["week"] <= week_range[1])
    & (weekly_df["categoria"].isin(selected_categories))
)

base_df = weekly_df.loc[mask].copy()
base_df = _apply_value(base_df, currency, selected_ranch)

prev_year = selected_year - 1
prev_mask = (
    (weekly_df["year"] == prev_year)
    & (weekly_df["week"] >= week_range[0])
    & (weekly_df["week"] <= week_range[1])
    & (weekly_df["categoria"].isin(selected_categories))
)
prev_df = _apply_value(weekly_df.loc[prev_mask].copy(), currency, selected_ranch)

symbol = _currency_symbol(currency)
total_actual = float(base_df["valor"].sum())
total_prev = float(prev_df["valor"].sum())
delta_abs = total_actual - total_prev
delta_pct = (delta_abs / total_prev * 100.0) if total_prev else 0.0
prom_semana = float(base_df.groupby("week")["valor"].sum().mean()) if not base_df.empty else 0.0

cat_totals = base_df.groupby("categoria", as_index=False)["valor"].sum().sort_values("valor", ascending=False)
top_cat = cat_totals.iloc[0]["categoria"] if not cat_totals.empty else "-"
top_cat_val = float(cat_totals.iloc[0]["valor"]) if not cat_totals.empty else 0.0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total periodo", f"{symbol} {total_actual:,.2f}", f"{delta_abs:,.2f}")
k2.metric("Variacion vs anio previo", f"{delta_pct:,.2f}%")
k3.metric("Promedio semanal", f"{symbol} {prom_semana:,.2f}")
k4.metric("Categoria lider", top_cat)
k5.metric("Valor categoria lider", f"{symbol} {top_cat_val:,.2f}")

tab1, tab2, tab3, tab4 = st.tabs(["Operacion", "Servicios", "Productos", "Calidad de datos"])

with tab1:
    st.subheader("Operacion semanal detallada")

    detail = (
        base_df.groupby(["week", "categoria"], as_index=False)["valor"].sum().rename(columns={"valor": "valor_actual"})
    )
    detail_prev = (
        prev_df.groupby(["week", "categoria"], as_index=False)["valor"].sum().rename(columns={"valor": "valor_prev"})
    )
    detail = detail.merge(detail_prev, on=["week", "categoria"], how="left")
    detail["valor_prev"] = detail["valor_prev"].fillna(0.0)
    detail["delta_abs"] = detail["valor_actual"] - detail["valor_prev"]
    detail["delta_pct"] = detail.apply(
        lambda r: ((r["delta_abs"] / r["valor_prev"]) * 100.0) if r["valor_prev"] else 0.0,
        axis=1,
    )
    detail = detail.sort_values(["week", "valor_actual"], ascending=[False, False])

    st.dataframe(
        detail,
        use_container_width=True,
        height=540,
        hide_index=True,
        column_config={
            "week": st.column_config.NumberColumn("Semana", format="%d"),
            "categoria": st.column_config.TextColumn("Categoria"),
            "valor_actual": st.column_config.NumberColumn(f"Actual ({symbol})", format="%.2f"),
            "valor_prev": st.column_config.NumberColumn(f"Previo ({symbol})", format="%.2f"),
            "delta_abs": st.column_config.NumberColumn("Delta abs", format="%.2f"),
            "delta_pct": st.column_config.NumberColumn("Delta %", format="%.2f"),
        },
    )

    c1, c2 = st.columns([1.3, 1])

    with c1:
        st.markdown("#### Matriz categoria x semana")
        matrix = (
            detail.pivot_table(index="categoria", columns="week", values="valor_actual", aggfunc="sum", fill_value=0)
            .sort_index()
        )
        if not matrix.empty:
            matrix["Total"] = matrix.sum(axis=1)
        st.dataframe(matrix, use_container_width=True, height=360)

    with c2:
        st.markdown("#### Ranking por rancho")
        ranch_totals: dict[str, float] = {}
        ranch_map_col = "usd_ranches" if currency == "USD" else "mxn_ranches"

        for _, row in base_df.iterrows():
            buckets = row.get(ranch_map_col) or {}
            for ranch_name, value in buckets.items():
                ranch_totals[ranch_name] = ranch_totals.get(ranch_name, 0.0) + _to_float(value)

        ranch_rank = pd.DataFrame(
            [{"rancho": k, "valor": v} for k, v in ranch_totals.items()]
        ).sort_values("valor", ascending=False)

        if not ranch_rank.empty:
            grand = float(ranch_rank["valor"].sum())
            ranch_rank["participacion_pct"] = ranch_rank["valor"].apply(lambda v: (v / grand * 100.0) if grand else 0.0)
            st.dataframe(
                ranch_rank,
                use_container_width=True,
                height=360,
                hide_index=True,
                column_config={
                    "rancho": st.column_config.TextColumn("Rancho"),
                    "valor": st.column_config.NumberColumn(f"Valor ({symbol})", format="%.2f"),
                    "participacion_pct": st.column_config.NumberColumn("Part. %", format="%.2f"),
                },
            )
        else:
            st.info("No hay datos de ranchos para el filtro actual.")

    st.download_button(
        label="Descargar detalle operativo (CSV)",
        data=_to_csv_bytes(detail),
        file_name=f"operacion_{selected_year}_{week_range[0]}_{week_range[1]}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with tab2:
    st.subheader("Costo de servicios")

    if serv_df.empty:
        st.info("No hay servicios_data disponibles.")
    else:
        serv_filtered = serv_df[
            (serv_df["year"] == selected_year)
            & (serv_df["week"] >= week_range[0])
            & (serv_df["week"] <= week_range[1])
        ].copy()

        if not serv_filtered.empty:
            serv_filtered = _apply_value(serv_filtered, currency, selected_ranch)

            subcats = sorted(serv_filtered["subcat"].dropna().unique().tolist())
            selected_subcats = st.multiselect("Subcategorias", options=subcats, default=subcats)
            if selected_subcats:
                serv_filtered = serv_filtered[serv_filtered["subcat"].isin(selected_subcats)]

            serv_table = (
                serv_filtered.groupby(["week", "subcat"], as_index=False)["valor"].sum().sort_values(["week", "valor"], ascending=[False, False])
            )

            st.dataframe(
                serv_table,
                use_container_width=True,
                height=520,
                hide_index=True,
                column_config={
                    "week": st.column_config.NumberColumn("Semana", format="%d"),
                    "subcat": st.column_config.TextColumn("Subcategoria"),
                    "valor": st.column_config.NumberColumn(f"Valor ({symbol})", format="%.2f"),
                },
            )

            subcat_rank = (
                serv_filtered.groupby("subcat", as_index=False)["valor"].sum().sort_values("valor", ascending=False)
            )
            st.markdown("#### Ranking subcategorias")
            st.dataframe(
                subcat_rank,
                use_container_width=True,
                height=260,
                hide_index=True,
                column_config={
                    "subcat": st.column_config.TextColumn("Subcategoria"),
                    "valor": st.column_config.NumberColumn(f"Valor ({symbol})", format="%.2f"),
                },
            )

            st.download_button(
                label="Descargar servicios (CSV)",
                data=_to_csv_bytes(serv_table),
                file_name=f"servicios_{selected_year}_{week_range[0]}_{week_range[1]}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.info("No hay registros de servicios para ese rango.")

with tab3:
    st.subheader("Productos PR / MP / ME")

    if prod_df.empty:
        st.info("No hay datos de productos disponibles.")
    else:
        p1, p2, p3 = st.columns([1.1, 1.1, 2.2])
        selected_sources = p1.multiselect("Origen", options=["PR", "MP", "ME"], default=["PR", "MP", "ME"])
        selected_tipo = p2.multiselect("Tipo", options=sorted(prod_df["tipo"].dropna().unique().tolist()), default=[])
        search_text = p3.text_input("Buscar producto", value="")

        p_mask = (
            (prod_df["year"] == selected_year)
            & (prod_df["week"] >= week_range[0])
            & (prod_df["week"] <= week_range[1])
        )

        prod_filtered = prod_df.loc[p_mask].copy()
        if selected_sources:
            prod_filtered = prod_filtered[prod_filtered["source"].isin(selected_sources)]
        if selected_ranch != "Todos":
            prod_filtered = prod_filtered[prod_filtered["rancho"] == selected_ranch]
        if selected_tipo:
            prod_filtered = prod_filtered[prod_filtered["tipo"].isin(selected_tipo)]
        if search_text.strip():
            term = search_text.strip().lower()
            prod_filtered = prod_filtered[prod_filtered["producto"].str.lower().str.contains(term, na=False)]

        prod_filtered = prod_filtered.sort_values(["week", "gasto"], ascending=[False, False])

        st.dataframe(
            prod_filtered,
            use_container_width=True,
            height=500,
            hide_index=True,
            column_config={
                "source": st.column_config.TextColumn("Origen"),
                "year": st.column_config.NumberColumn("Anio", format="%d"),
                "week": st.column_config.NumberColumn("Semana", format="%d"),
                "rancho": st.column_config.TextColumn("Rancho"),
                "tipo": st.column_config.TextColumn("Tipo"),
                "producto": st.column_config.TextColumn("Producto"),
                "unidades": st.column_config.NumberColumn("Unidades", format="%.2f"),
                "gasto": st.column_config.NumberColumn(f"Gasto ({symbol})", format="%.2f"),
                "ubicacion": st.column_config.TextColumn("Ubicacion"),
            },
        )

        top_products = (
            prod_filtered.groupby(["source", "producto"], as_index=False)["gasto"]
            .sum()
            .sort_values("gasto", ascending=False)
            .head(25)
        )
        st.markdown("#### Top 25 productos por gasto")
        st.dataframe(
            top_products,
            use_container_width=True,
            height=280,
            hide_index=True,
            column_config={
                "source": st.column_config.TextColumn("Origen"),
                "producto": st.column_config.TextColumn("Producto"),
                "gasto": st.column_config.NumberColumn(f"Gasto ({symbol})", format="%.2f"),
            },
        )

        st.download_button(
            label="Descargar productos (CSV)",
            data=_to_csv_bytes(prod_filtered),
            file_name=f"productos_{selected_year}_{week_range[0]}_{week_range[1]}.csv",
            mime="text/csv",
            use_container_width=True,
        )

with tab4:
    st.subheader("Calidad y cobertura de datos")

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Registros weekly_detail", f"{len(weekly_df):,}")
    q2.metric("Registros servicios_data", f"{len(serv_df):,}")
    q3.metric("Registros productos", f"{len(prod_df):,}")
    max_week = int(weekly_df.loc[weekly_df["year"] == latest_year, "week"].max())
    q4.metric("Ultimo corte", f"{latest_year}-W{max_week:02d}")

    cov_rows = []
    for y in all_years:
        weeks = sorted(weekly_df.loc[weekly_df["year"] == y, "week"].unique().tolist())
        if not weeks:
            continue
        expected = int(max(weeks) - min(weeks) + 1)
        observed = len(weeks)
        missing = expected - observed
        cov_rows.append(
            {
                "anio": y,
                "semana_min": int(min(weeks)),
                "semana_max": int(max(weeks)),
                "esperadas": expected,
                "observadas": observed,
                "faltantes": missing,
                "cobertura_pct": (observed / expected * 100.0) if expected else 0.0,
            }
        )

    cov_df = pd.DataFrame(cov_rows).sort_values("anio", ascending=False)
    st.markdown("#### Cobertura por anio")
    st.dataframe(
        cov_df,
        use_container_width=True,
        height=300,
        hide_index=True,
        column_config={
            "anio": st.column_config.NumberColumn("Anio", format="%d"),
            "semana_min": st.column_config.NumberColumn("Semana min", format="%d"),
            "semana_max": st.column_config.NumberColumn("Semana max", format="%d"),
            "esperadas": st.column_config.NumberColumn("Esperadas", format="%d"),
            "observadas": st.column_config.NumberColumn("Observadas", format="%d"),
            "faltantes": st.column_config.NumberColumn("Faltantes", format="%d"),
            "cobertura_pct": st.column_config.NumberColumn("Cobertura %", format="%.2f"),
        },
    )

    invalid_cats = weekly_df[weekly_df["categoria"].isna() | (weekly_df["categoria"].astype(str).str.strip() == "")]
    invalid_weeks = weekly_df[(weekly_df["week"] <= 0) | (weekly_df["week"] > 53)]

    i1, i2 = st.columns(2)
    i1.metric("Categorias vacias", f"{len(invalid_cats):,}")
    i2.metric("Semanas fuera de rango", f"{len(invalid_weeks):,}")

st.caption("CFBC Data Console - Modo empresarial: informacion densa, accionable y exportable.")
