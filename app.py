"""
app.py
Centro Floricultor de Baja California
Streamlit — tablas HTML estilo tabla dinámica Excel, sin AG Grid
"""

import json
import base64
import os
import streamlit as st
import streamlit.components.v1 as components

from data_extractor import get_datos

st.set_page_config(
    page_title="CFBC WK",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

if "show_auto" not in st.session_state:
    st.session_state.show_auto = False

def toggle_auto():
    st.session_state.show_auto = not st.session_state.show_auto

if st.session_state.show_auto:
    # CSS para el Panel de Automatizacion estilo consola administrativa
    st.markdown('''
    <style>
      #MainMenu, header, footer { display: none !important; }
      .stApp {
        background:
          linear-gradient(180deg, #edf1f5 0%, #e4e9ef 100%);
      }
      .block-container {
        padding: 1rem !important;
        max-width: 100% !important;
      }
      .stMainBlockContainer {
        padding-top: 0 !important;
      }
      [data-testid="stVerticalBlock"] { gap: 0.9rem !important; }
      section[data-testid="stSidebar"] { display: none !important; }

      div[data-testid="stHorizontalBlock"]:has(#auto-topbar-left) {
        background: linear-gradient(180deg, #2f3c4b 0%, #263240 100%) !important;
        border: 1px solid #1f2933 !important;
        border-radius: 16px !important;
        padding: 0.65rem 0.8rem !important;
        align-items: center !important;
        box-shadow: 0 10px 28px rgba(22, 34, 48, 0.20) !important;
        gap: 0.55rem !important;
      }
      div[data-testid="stColumn"]:has(#auto-topbar-left),
      div[data-testid="stColumn"]:has(#auto-topbar-reload),
      div[data-testid="stColumn"]:has(#auto-topbar-back) {
        display: flex !important;
        align-items: center !important;
      }
      div[data-testid="stColumn"]:has(#auto-topbar-left) {
        min-height: 58px !important;
      }
      div[data-testid="stColumn"]:has(#auto-topbar-reload) div[data-testid="stButton"],
      div[data-testid="stColumn"]:has(#auto-topbar-back) div[data-testid="stButton"] {
        width: 100% !important;
      }
      div[data-testid="stColumn"]:has(#auto-topbar-reload) button,
      div[data-testid="stColumn"]:has(#auto-topbar-back) button {
        min-height: 40px !important;
        border-radius: 10px !important;
        border: 1px solid rgba(255,255,255,0.18) !important;
        background: rgba(255,255,255,0.08) !important;
        color: #ffffff !important;
      }
      div[data-testid="stColumn"]:has(#auto-topbar-reload) button:hover,
      div[data-testid="stColumn"]:has(#auto-topbar-back) button:hover {
        background: rgba(255,255,255,0.18) !important;
      }

      .auto-topbar-wrap {
        display: flex;
        align-items: center;
        gap: 0.8rem;
      }
      .auto-topbar-icon {
        width: 38px;
        height: 38px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(255,255,255,0.08);
        color: #ffffff;
        font-size: 18px;
        font-weight: 800;
      }
      .auto-topbar-kicker {
        color: #9eb0c6;
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1.1px;
        text-transform: uppercase;
        margin-bottom: 2px;
      }
      .auto-topbar-title {
        color: #ffffff;
        font-size: 18px;
        font-weight: 800;
        line-height: 1.05;
      }
      .auto-topbar-subtitle {
        color: #b8c6d8;
        font-size: 11px;
        margin-top: 2px;
      }

      div[data-testid="stColumn"]:has(#auto-sidebar-shell) {
        background: linear-gradient(180deg, #2d3745 0%, #24303c 100%) !important;
        border: 1px solid #1f2933 !important;
        border-radius: 22px !important;
        box-shadow: 0 16px 34px rgba(22, 34, 48, 0.18) !important;
        padding: 1.2rem 1rem 1rem 1rem !important;
        min-height: calc(100vh - 128px) !important;
      }
      div[data-testid="stColumn"]:has(#auto-main-shell) {
        min-width: 0 !important;
      }
      div[data-testid="stVerticalBlock"]:has(#auto-hero-shell),
      div[data-testid="stColumn"]:has(#auto-stat-1),
      div[data-testid="stColumn"]:has(#auto-stat-2),
      div[data-testid="stColumn"]:has(#auto-stat-3),
      div[data-testid="stColumn"]:has(#auto-stat-4),
      div[data-testid="stColumn"]:has(#auto-card-download),
      div[data-testid="stColumn"]:has(#auto-card-create),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-mv-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-siembra-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-nomina-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-conteo-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-upload-shell),
      div[data-testid="stColumn"]:has(#auto-upload-pr),
      div[data-testid="stColumn"]:has(#auto-upload-mp),
      div[data-testid="stColumn"]:has(#auto-upload-me),
      div[data-testid="stColumn"]:has(#auto-upload-mv),
      div[data-testid="stColumn"]:has(#auto-upload-tt) {
        background: rgba(255,255,255,0.97) !important;
        border: 1px solid #d5dde7 !important;
        border-radius: 20px !important;
        box-shadow: 0 14px 30px rgba(31, 45, 61, 0.08) !important;
      }
      div[data-testid="stVerticalBlock"]:has(#auto-hero-shell) {
        padding: 1.25rem 1.35rem 1.1rem 1.35rem !important;
        margin-bottom: 0.1rem !important;
      }
      div[data-testid="stColumn"]:has(#auto-stat-1),
      div[data-testid="stColumn"]:has(#auto-stat-2),
      div[data-testid="stColumn"]:has(#auto-stat-3),
      div[data-testid="stColumn"]:has(#auto-stat-4),
      div[data-testid="stColumn"]:has(#auto-card-download),
      div[data-testid="stColumn"]:has(#auto-card-create),
      div[data-testid="stColumn"]:has(#auto-upload-pr),
      div[data-testid="stColumn"]:has(#auto-upload-mp),
      div[data-testid="stColumn"]:has(#auto-upload-me),
      div[data-testid="stColumn"]:has(#auto-upload-mv),
      div[data-testid="stColumn"]:has(#auto-upload-tt) {
        padding: 1rem 1rem 0.95rem 1rem !important;
      }
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-mv-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-siembra-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-nomina-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-card-fill-conteo-shell),
      div[data-testid="stVerticalBlock"]:has(#auto-upload-shell) {
        padding: 1rem 1rem 0.95rem 1rem !important;
      }

      .auto-sidebar-logo {
        width: 62px;
        height: 62px;
        border-radius: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(135deg, #ffffff 0%, #e7edf4 100%);
        color: #7B1C1C;
        font-size: 18px;
        font-weight: 900;
        letter-spacing: 1px;
        margin-bottom: 1rem;
      }
      .auto-sidebar-caption {
        color: #dfe7f0;
        font-size: 17px;
        font-weight: 800;
        line-height: 1.2;
        margin-bottom: 0.3rem;
      }
      .auto-sidebar-note {
        color: #aebccc;
        font-size: 12px;
        line-height: 1.45;
        margin-bottom: 1rem;
      }
      .auto-sidebar-group {
        color: #7f93aa;
        font-size: 10px;
        font-weight: 800;
        letter-spacing: 1.1px;
        text-transform: uppercase;
        margin: 1rem 0 0.45rem 0;
      }
      .auto-nav-item {
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 12px;
        background: rgba(255,255,255,0.04);
        color: #edf3f9;
        font-size: 13px;
        font-weight: 700;
        padding: 0.72rem 0.85rem;
        margin-bottom: 0.45rem;
      }
      .auto-nav-item-active {
        background: linear-gradient(135deg, rgba(123,28,28,0.95) 0%, rgba(164,56,56,0.92) 100%);
        border-color: rgba(255,255,255,0.14);
      }
      .auto-sidebar-metric {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0.6rem 0.05rem;
        border-bottom: 1px solid rgba(255,255,255,0.08);
        color: #dfe7f0;
        font-size: 12px;
      }
      .auto-sidebar-metric strong {
        color: #ffffff;
        font-size: 13px;
      }

      .auto-eyebrow {
        color: #7B1C1C;
        font-size: 10px;
        font-weight: 800;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        margin-bottom: 8px;
      }
      .auto-title {
        color: #1c2c3f;
        font-size: 31px;
        font-weight: 800;
        line-height: 1.08;
        margin: 0 0 0.45rem 0;
      }
      .auto-subtitle {
        color: #5c6c7e;
        font-size: 13px;
        line-height: 1.5;
        margin: 0 0 0.9rem 0;
        max-width: 900px;
      }
      .auto-hero-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
      }
      .auto-hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.42rem 0.7rem;
        border-radius: 999px;
        border: 1px solid #dce4ee;
        background: #f7f9fc;
        color: #3c4d61;
        font-size: 11px;
        font-weight: 700;
      }
      .auto-stat-label {
        color: #7a8798;
        font-size: 10px;
        font-weight: 800;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 8px;
      }
      .auto-stat-value {
        color: #1c2c3f;
        font-size: 24px;
        font-weight: 800;
        line-height: 1;
        margin-bottom: 6px;
      }
      .auto-stat-note {
        color: #6d7c90;
        font-size: 11px;
        line-height: 1.35;
      }
      .auto-card-kicker {
        color: #7B1C1C;
        font-size: 10px;
        font-weight: 800;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 6px;
      }
      .auto-card-title {
        color: #1c2c3f;
        font-size: 19px;
        font-weight: 800;
        margin-bottom: 4px;
      }
      .auto-card-note {
        color: #66768a;
        font-size: 12px;
        line-height: 1.45;
        margin-bottom: 10px;
      }
      .auto-section-title {
        color: #1c2c3f;
        font-size: 22px;
        font-weight: 800;
        margin: 6px 0 2px 0;
      }
      .auto-section-note {
        color: #66768a;
        font-size: 12px;
        line-height: 1.45;
        margin: 0 0 10px 0;
      }
      .auto-mini-title {
        color: #1c2c3f;
        font-size: 15px;
        font-weight: 800;
        margin-bottom: 2px;
      }
      .auto-mini-note {
        color: #6d7c90;
        font-size: 11px;
        margin-bottom: 8px;
      }
      div[data-testid="stColumn"]:has(#auto-stat-1),
      div[data-testid="stColumn"]:has(#auto-stat-2),
      div[data-testid="stColumn"]:has(#auto-stat-3),
      div[data-testid="stColumn"]:has(#auto-stat-4) {
        min-height: 118px !important;
      }
      div[data-testid="stColumn"]:has(#auto-upload-pr),
      div[data-testid="stColumn"]:has(#auto-upload-mp),
      div[data-testid="stColumn"]:has(#auto-upload-me),
      div[data-testid="stColumn"]:has(#auto-upload-mv) {
        min-height: 100% !important;
      }

      div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 0.55rem !important;
        margin-bottom: 1rem !important;
        border-bottom: 1px solid #d7dee7 !important;
        padding-left: 0.15rem !important;
      }
      div[data-testid="stTabs"] button[role="tab"] {
        height: 40px !important;
        border-radius: 12px 12px 0 0 !important;
        border: 1px solid #d7dee7 !important;
        border-bottom: none !important;
        background: #f5f7fa !important;
        color: #4f6277 !important;
        padding: 0 1rem !important;
        font-size: 12px !important;
        font-weight: 700 !important;
      }
      div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: #ffffff !important;
        color: #1c2c3f !important;
        box-shadow: inset 0 3px 0 #7B1C1C !important;
      }

      div[data-testid="stButton"] button,
      div[data-testid="stDownloadButton"] button {
        border-radius: 12px !important;
        min-height: 44px !important;
        font-size: 12px !important;
        font-weight: 700 !important;
      }
      div[data-testid="stButton"] button[kind="primary"],
      div[data-testid="stDownloadButton"] button[kind="primary"] {
        background: linear-gradient(180deg, #6f94bc 0%, #55799f 100%) !important;
        border: 1px solid #4d6d8f !important;
        color: #ffffff !important;
      }
      div[data-testid="stFileUploaderDropzone"] {
        padding: 0.95rem 1rem !important;
        border-radius: 14px !important;
        background: #f8fbff !important;
        border: 1px dashed #bfd0e4 !important;
      }
      div[data-testid="stFileUploaderDropzone"] * {
        font-size: 12px !important;
      }
      div[data-testid="stSelectbox"] label,
      div[data-testid="stTextInput"] label,
      div[data-testid="stFileUploader"] label {
        font-size: 11px !important;
        font-weight: 700 !important;
        color: #516173 !important;
      }
      div[data-testid="stSelectbox"] > div,
      div[data-testid="stTextInput"] > div {
        margin-top: 0.1rem !important;
      }
      div[data-baseweb="select"] > div,
      div[data-testid="stTextInput"] input {
        border-radius: 12px !important;
        border: 1px solid #d5dde7 !important;
        background: #fbfcfe !important;
      }
      div[data-testid="stAlert"] {
        border-radius: 14px !important;
        padding-top: 0.55rem !important;
        padding-bottom: 0.55rem !important;
      }
      hr {
        margin: 0.35rem 0 0.2rem 0 !important;
        border-top: 1px solid #dbe4ef !important;
      }
    </style>
    ''', unsafe_allow_html=True)
else:
    # CSS para el Dashboard — header nativo de Streamlit, sin trucos de iframe
    st.markdown('''
    <style>
      #MainMenu, header, footer { display: none !important; }
      .stApp { background: #ffffff; }
      /* Eliminar TODO espacio blanco sobre el header */
      .block-container          { padding: 0 !important; max-width: 100% !important; margin: 0 !important; }
      .stMainBlockContainer     { padding: 0 !important; margin: 0 !important; }
      [data-testid="stAppViewContainer"] { padding-top: 0 !important; margin-top: 0 !important; }
      [data-testid="stAppViewBlockContainer"] { padding: 0 !important; }
      section.main              { padding-top: 0 !important; }
      /* Primer bloque vertical (donde vive nuestro header de columnas) */
      [data-testid="stVerticalBlock"] { gap: 0 !important; padding: 0 !important; }
      section[data-testid="stSidebar"] { display: none !important; }

      /* ── HEADER NATIVO: columnas de Streamlit que contienen #cfbc-brand ── */
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) {
          background: #7B1C1C !important;
          border-bottom: none !important;
          min-height: 36px !important;
          max-height: 36px !important;
          padding: 0 10px !important;
          align-items: center !important;
          overflow: hidden !important;
          gap: 4px !important;
          margin: 0 !important;
      }
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) > [data-testid="stColumn"] {
          padding: 0 2px !important;
          display: flex !important;
          align-items: center !important;
      }
      /* Texto de marca */
      #cfbc-brand {
          color: #ffffff;
          font-family: Calibri, 'Segoe UI', Arial, sans-serif;
          font-size: 12px;
          font-weight: 700;
          letter-spacing: 1px;
          white-space: nowrap;
      }
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) p {
          margin: 0 !important; padding: 0 !important; line-height: 36px !important;
      }
      /* Botones dentro del header nativo */
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) div[data-testid="stButton"] {
          width: 100% !important;
          display: flex !important;
          justify-content: flex-end !important;
          align-items: center !important;
          padding: 0 !important;
          margin: 0 !important;
      }
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) div[data-testid="stButton"] button {
          background: rgba(255,255,255,0.2) !important;
          color: #ffffff !important;
          border: 1px solid rgba(255,255,255,0.35) !important;
          border-radius: 3px !important;
          height: 24px !important;
          min-height: 24px !important;
          padding: 0 10px !important;
          font-size: 10px !important;
          font-weight: 700 !important;
          white-space: nowrap !important;
          cursor: pointer !important;
          width: auto !important;
      }
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) div[data-testid="stButton"] button:hover {
          background: rgba(255,255,255,0.38) !important;
      }
      div[data-testid="stHorizontalBlock"]:has(#cfbc-brand) div[data-testid="stButton"] button p {
          color: #ffffff !important;
          font-size: 10px !important;
          font-weight: 700 !important;
          margin: 0 !important;
          line-height: 1 !important;
      }
    </style>
    ''', unsafe_allow_html=True)


@st.cache_data(ttl=300, show_spinner=False)
def load_data_conteo_v3():
    return get_datos()


try:
    DATA = load_data_conteo_v3()
except Exception as e:
    st.error(f"❌ Error cargando datos: {e}")
    st.stop()

if "error" in DATA:
    st.error(f"❌ {DATA['error']}")
    if st.button("🔄 Reintentar"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

import math

def _sanitize(obj):
    """Convierte NaN/Inf a 0 recursivamente para JSON válido."""
    if isinstance(obj, float):
        return 0 if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _sanitize(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize(v) for v in obj]
    return obj

data_json = base64.b64encode(
    json.dumps(_sanitize(DATA), ensure_ascii=True, default=str).encode('utf-8')
).decode('ascii')

APP_CSS = """<style>
:root {
  --navy:   #1e3a5f;
  --green:  #16a34a;
  --red:    #dc2626;
  --amber:  #b45309;
  --blue:   #2563eb;
  --border: #d0d0d0;

  /* Pivot-table palette — Excel style */
  --pt-hdr-bg:      #D9E1F2;   /* header de columnas  */
  --pt-hdr-border:  #8EA9C1;
  --pt-grp-bg:      #82A1C9;   /* fila de grupo/año - azul más suavizado */
  --pt-grp-fg:      #ffffff;
  --pt-sub-bg:      #BDD7EE;   /* fila subtotal       */
  --pt-sub-fg:      #000000;
  --pt-tot-bg:      #9DC3E6;   /* fila total general  */
  --pt-tot-fg:      #000000;
  --pt-row-hover:   #EBF3FB;
  --pt-cell-border: #BFBFBF;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: Calibri, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  font-size: 12px;
  background: #ffffff;
#loader {
  position: fixed; inset: 0; background: #fff; z-index: 999;
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; gap: 14px;
}
.spin {
  width: 36px; height: 36px;
  border: 3px solid #e0e0e0; border-top-color: var(--green);
  border-radius: 50%; animation: spin 0.9s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
.load-txt { font-size: 12px; color: #666; letter-spacing: 0.5px; }

/* ── HEADER: ahora es nativo de Streamlit, no del iframe ─────────────── */

/* ── TOOLBAR ─────────────────────────────────── */
.toolbar {
  background: #ebebeb; border-bottom: 1px solid var(--border);
  padding: 2px 8px; display: flex; align-items: center; gap: 6px;
  flex-wrap: nowrap; overflow-x: auto; scrollbar-width: none; height: 28px;
}
.toolbar::-webkit-scrollbar { display: none; }
.tb-label { font-size: 9px; color: #777; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; flex-shrink: 0; }
.tb-sep   { width: 1px; height: 18px; background: #ccc; flex-shrink: 0; }
select.tb-sel {
  font-size: 11px; font-family: inherit;
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 2px 6px; color: #222; cursor: pointer; height: 22px; flex-shrink: 0;
}
select.tb-sel:focus { outline: 2px solid var(--green); outline-offset: -1px; }
.tb-btn {
  font-size: 10px; font-weight: 700; font-family: inherit;
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 2px 8px; cursor: pointer; height: 22px;
  white-space: nowrap; color: #333; transition: background 0.1s; flex-shrink: 0;
}
.tb-btn:hover  { background: #ddd; }
.tb-btn.active { background: var(--navy); color: #fff; border-color: var(--navy); }
.tb-grp { display: flex; flex-shrink: 0; }
.tb-grp .tb-btn { border-radius: 0; border-right-width: 0; }
.tb-grp .tb-btn:first-child { border-radius: 3px 0 0 3px; }
.tb-grp .tb-btn:last-child  { border-radius: 0 3px 3px 0; border-right-width: 1px; }
.week-ctr { display: flex; align-items: center; gap: 4px; flex-shrink: 0; }
.week-ctr span { font-size: 11px; font-weight: 700; color: var(--navy); min-width: 62px; text-align: center; }
.tb-slider  { width: 100px; accent-color: var(--green); cursor: pointer; flex-shrink: 0; }
.yr-chip {
  font-size: 10px; font-weight: 700; padding: 1px 7px; border-radius: 3px;
  cursor: pointer; border: 1px solid transparent; background: transparent;
  transition: all 0.1s; flex-shrink: 0;
}
.yr-chip.on { background: #fff; }

/* ── RANGE BAR ───────────────────────────────── */
.range-bar {
  display: none; background: #f4f4f4; border-bottom: 1px solid var(--border);
  padding: 3px 10px; align-items: center; gap: 8px; height: 26px; overflow: hidden;
}
.range-bar.show { display: flex; }
.range-val   { font-size: 11px; font-weight: 700; color: var(--navy); min-width: 36px; text-align: center; }
.range-badge {
  font-size: 10px; background: #e8f5e9; border: 1px solid #a7d7b4;
  color: var(--green); padding: 1px 8px; border-radius: 3px; white-space: nowrap; flex-shrink: 0;
}

/* ── VIEW TABS ───────────────────────────────── */
.view-tabs {
  background: #f8f8f8; border-bottom: 2px solid #d5d5d5;
  display: flex; padding: 0; height: 28px;
}
.vtab {
  padding: 0 14px; font-size: 10px; font-weight: 700; font-family: inherit;
  cursor: pointer; border: none; background: transparent; color: #888;
  border-bottom: 2px solid transparent; margin-bottom: -2px;
  text-transform: uppercase; letter-spacing: 0.5px;
  transition: color 0.1s; white-space: nowrap; height: 28px;
}
.vtab:hover  { color: #333; background: rgba(0,0,0,0.03); }
.vtab.active { color: var(--green); border-bottom-color: var(--green); background: #fff; }

/* ── TABLE WRAPPER ───────────────────────────── */
#gridWrap {
  background: #fff;
  border: 1px solid #d5d5d5;
  border-top: none;
  overflow: visible;
}
.pt-table-wrap {
  overflow-x: auto;
  overflow-y: visible;
  width: 100%;
}
.pt-table-wrap::-webkit-scrollbar { height: 6px; width: 6px; }
.pt-table-wrap::-webkit-scrollbar-thumb { background: #b0c4d8; border-radius: 3px; }

/* ── PIVOT TABLE ─────────────────────────────── */
.pt-table {
  border-collapse: collapse;
  width: 100%;
  font-size: 12px;
  font-family: Calibri, 'Segoe UI', Arial, sans-serif;
  white-space: nowrap;
}
.pt-table th {
  background: var(--pt-hdr-bg);
  color: #1e3a5f;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.3px;
  padding: 5px 8px;
  border: none;
  border-bottom: 1px solid var(--pt-hdr-border);
  white-space: nowrap;
  position: sticky;
  top: 0;
  z-index: 3;
  user-select: none;
}
.pt-table th.pt-pinned { position: sticky; left: 0; z-index: 4; }
.pt-table td {
  padding: 3px 8px;
  border: none;
  border-bottom: 1px solid #eeeeee;
  white-space: nowrap;
  color: #000000;
  height: 24px;
  line-height: 18px;
}
.pt-table td.pt-pinned {
  position: sticky; left: 0; z-index: 1;
  background: inherit;
}
/* regular row */
.pt-row { background: #fff; }
.pt-row:hover td { background: var(--pt-row-hover) !important; }
/* alternating */
.pt-row:nth-child(even) { background: #F7FBFF; }
/* group header row */
.pt-row-group td {
  background: var(--pt-grp-bg) !important;
  color: var(--pt-grp-fg);
  font-weight: 700;
}
/* subtotal row */
.pt-row-sub td {
  background: var(--pt-sub-bg) !important;
  color: var(--pt-sub-fg);
  font-weight: 700;
}
/* total row */
.pt-row-total td {
  background: var(--pt-tot-bg) !important;
  color: var(--pt-tot-fg);
  font-weight: 700;
}
/* muted/dash cells */
.cell-muted { color: #bbb !important; }
.cell-pos   { color: #16a34a !important; font-weight: 600; }
.cell-neg   { color: #dc2626 !important; font-weight: 600; }
.cell-navy  { color: #1e3a5f !important; font-weight: 600; }
.prod-link  { cursor: pointer; text-decoration: underline dotted; text-underline-offset: 2px; }

/* ── COMPARATIVO TABLE ───────────────────────── */
#comparativoWrap { display: none; background: #fff; border: 1px solid #d5d5d5; border-top: none; }
#comparativoWrap.show { display: block; }
.cmp-stat-strip { display: flex; gap: 8px; flex-wrap: wrap; padding: 8px 10px; background: #f4f4f4; border-bottom: 1px solid #d5d5d5; }
.cmp-tbl-wrap { overflow-x: auto; scrollbar-width: thin; scrollbar-color: #b0c4d8 transparent; }
.cmp-tbl-wrap::-webkit-scrollbar { height: 5px; width: 5px; }
.cmp-tbl-wrap::-webkit-scrollbar-thumb { background: #b0c4d8; border-radius: 3px; }
.cmp-tbl { border-collapse: collapse; width: 100%; font-size: 12px; font-family: Calibri,'Segoe UI',Arial,sans-serif; }
.cmp-tbl th {
  padding: 5px 8px; background: var(--pt-hdr-bg); color: #1e3a5f;
  font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.3px;
  white-space: nowrap; border: none; border-bottom: 1px solid var(--pt-hdr-border);
  position: sticky; top: 0; z-index: 2; text-align: right;
}
.cmp-tbl th:first-child, .cmp-tbl th:nth-child(2) { text-align: left; }
.cmp-tbl td { padding: 3px 8px; border: none; border-bottom: 1px solid #eeeeee; white-space: nowrap; text-align: right; height: 24px; }
.cmp-tbl td:first-child, .cmp-tbl td:nth-child(2) { text-align: left; }
.cmp-grp-hdr td {
  background: var(--pt-grp-bg) !important; color: var(--pt-grp-fg);
  font-weight: 700;
}
.cmp-grp-hdr td:first-child { border-left: 3px solid #4ade80; }
.cmp-row { background: #fff; }
.cmp-row:hover td { background: var(--pt-row-hover) !important; }
.cmp-row:nth-child(even) { background: #F7FBFF; }
.cmp-total-row td { background: var(--pt-sub-bg) !important; font-weight: 700; color: var(--pt-sub-fg); }
.delta-cell { font-size: 10px; white-space: nowrap; }
.delta-amt  { display: block; }
.delta-pct  { display: block; font-size: 9px; opacity: 0.8; }
.chg-pos { color: #16a34a; font-weight: 600; }
.chg-neg { color: #dc2626; font-weight: 600; }
.chg-0   { color: #aaa; }

/* ── PRODUCTOS PANEL ─────────────────────────── */
#prodPanel { 
  display: none; background: #fdfdfd; border: 1px solid #cbd5e1; border-top: 2px solid #7B1C1C;
  box-shadow: 0 4px 12px rgba(0,0,0,0.06);
  margin: 5px 0 0 0; width: 100%; overflow: hidden;
}
#prodPanel.show { display: block; }
#prodTableWrap { overflow: visible; }

/* ── STATUS BAR ──────────────────────────────── */
.statusbar { display: none !important; }
.statusbar b { color: #333; }
</style>"""

APP_HTML_BODY = """
<!-- LOADER -->
<div id="ranchDropdownPanel" style="display:none;position:fixed;z-index:9999;background:#fff;border:1px solid #bbb;border-radius:4px;box-shadow:0 4px 12px rgba(0,0,0,0.15);min-width:150px;max-height:220px;overflow-y:auto;padding:4px 0;"></div>
<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">CFBC &#8212; Cargando datos...</div>
</div>

<!-- APP -->
<div id="app" style="display:none">


  <!-- TOOLBAR -->
  <div class="toolbar">
    <span class="tb-label">Cat</span>
    <select class="tb-sel" id="catSel" onchange="onCatChange(this.value)" style="max-width:200px"></select>
    <div class="tb-sep"></div>
    <span class="tb-label">Rancho</span>
    <button class="tb-btn" id="ranchDropdownBtn" onclick="toggleRanchDropdown(event)" style="flex-shrink:0;min-width:90px;max-width:180px;text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">Todos ▾</button>
    <div class="tb-sep"></div>
    <div class="tb-grp">
      <button class="tb-btn"        id="btnUSD" onclick="setCurrency('usd')">USD</button>
      <button class="tb-btn active" id="btnMXN" onclick="setCurrency('mxn')">MXN</button>
    </div>
    <div class="tb-sep"></div>
    <span class="tb-label">Desde</span>
    <span class="range-val" id="fromWeekLabel">W01</span>
    <input type="range" class="tb-slider" id="fromSlider" min="1" max="52" value="1" oninput="onRangeChange()">
    <span style="color:#aaa;font-size:11px;flex-shrink:0;">→</span>
    <span class="tb-label">Hasta</span>
    <span class="range-val" id="toWeekLabel">W52</span>
    <input type="range" class="tb-slider" id="toSlider" min="1" max="52" value="52" oninput="onRangeChange()">
    <span class="range-badge" id="rangeBadge">W01 → W52</span>
    <div class="tb-sep"></div>
    <span class="tb-label">Años</span>
    <div id="yearChips" style="display:flex;gap:3px"></div>
    <div class="tb-sep" id="sepVerProd" style="display:none"></div>
    <button class="tb-btn" id="btnVerProd" onclick="abrirProdGlobal()" style="display:none; color:#16a34a; border-color:#16a34a;" title="Ver desglose del rango seleccionado">🔍 VER PRODUCTOS</button>
  </div>

  <!-- VIEW TABS -->
  <div class="view-tabs">
    <button class="vtab active" id="vtComparativo"  onclick="setView('comparativo')">Comparativo</button>
    <button class="vtab"        id="vtRancho"       onclick="setView('rancho')">Por Rancho</button>
    <button class="vtab"        id="vtServicios"    onclick="setView('servicios')">Costo Servicios</button>
  </div>

  <!-- RANGE BAR eliminada — controles movidos al toolbar -->

  <!-- MAIN TABLE AREA (todas las vistas excepto comparativo) -->
  <div id="gridWrap">
    <div class="pt-table-wrap" id="tableWrap" style="overflow:visible"></div>
  </div>

  <!-- COMPARATIVO TABLE -->
  <div id="comparativoWrap">
    <div class="cmp-stat-strip" id="cmpStats"></div>
    <div class="cmp-tbl-wrap">
      <table class="cmp-tbl">
        <thead id="cmpHead"></thead>
        <tbody id="cmpBody"></tbody>
      </table>
    </div>
  </div>

  <!-- PRODUCTOS SUB-PANEL -->
  <div id="prodPanel">
    <div id="prodTableWrap" style="display:flex; gap:6px; padding:6px; overflow-x:auto;"></div>
  </div>

  <!-- STATUS BAR -->
  <div class="statusbar" id="statusbar">
    <span>Total: <b id="stTotal">&#8212;</b></span>
  </div>
</div>
"""

APP_JS = """<script>
// =======================================================
// ERROR HANDLER &#8212; primero de todo
// =======================================================
window.onerror = function(msg, src, line, col, err) {
  var loader = document.getElementById('loader');
  if (loader) loader.innerHTML =
    '<div style="color:#dc2626;font-family:monospace;padding:20px;background:#fff;' +
    'border-radius:8px;border:1px solid #fecaca;max-width:600px;margin:20px auto">' +
    '<b>ERROR (línea ' + line + '):</b><br>' + msg +
    (err && err.stack ? '<br><pre style="font-size:10px;color:#999;margin-top:8px;overflow:auto">' + err.stack + '</pre>' : '') +
    '</div>';
  return true;
};

// =======================================================
// DATOS
// =======================================================
var DATA;
try {
  var _raw = atob('__DATA_JSON__');
  DATA = JSON.parse(_raw);
} catch(e) {
  document.getElementById('loader').innerHTML =
    '<div style="color:#dc2626;font-family:monospace;padding:20px;background:#fff;border-radius:8px;border:1px solid #fecaca;max-width:500px;margin:20px auto">' +
    '<b>Error parseando datos:</b> ' + e.message + '</div>';
}

// =======================================================
// CONSTANTES Y CONFIGURACIÓN DINÁMICA
// =======================================================
var RANCH_ORDER  = (DATA && DATA.config) ? DATA.config.ranch_order : [];
var RANCH_COLORS = (DATA && DATA.config) ? DATA.config.ranch_colors : {};
var YEAR_COLORS = {2021:'#0e7490',2022:'#d97706',2023:'#16a34a',2024:'#9333ea',2025:'#f97316',2026:'#dc2626'};
var CAT_MIRFE = 'FERTILIZANTES';
var CAT_MIPE  = 'DESINFECCION / PLAGUICIDAS';

// =======================================================
// ESTADO
// =======================================================
var state = { cat:'', activeRanches:['Todos'], currency:'mxn', activeYears:{}, view:'comparativo', weekIdx:0, fromWeek:1, toWeek:52 };
function getActiveRanches() { return state.activeRanches.indexOf('Todos') > -1 ? RANCH_ORDER : state.activeRanches; }
var allWeeks = [];

// =======================================================
// TABLA PIVOT &#8212; estado global
// =======================================================
var _tableRows    = [];
var _tableColDefs = [];

// =======================================================
// FORMATEO
// =======================================================
function fmt(n) {
  if (n === null || n === undefined || n === 0 || isNaN(n)) return '';
  var neg = n < 0, s = Math.abs(n);
  return (neg ? '-$' : '$') + s.toLocaleString('en-US', {minimumFractionDigits:0, maximumFractionDigits:0});
}
function fmtFull(n) {
  if (!n || isNaN(n)) return '';
  var neg = n < 0, s = Math.abs(n);
  return (neg ? '-$' : '$') + s.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
}
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '';
  var sign = n > 0 ? '+' : '';
  return sign + n.toFixed(1) + '%';
}
function fmtHa(n) {
  if (n === null || n === undefined || n === 0 || isNaN(n)) return '';
  var neg = n < 0, s = Math.abs(n);
  return (neg ? '-' : '') + s.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:4});
}
function fmtHc(n) {
  if (!n || isNaN(n) || n === 0) return '';
  return Math.round(n).toLocaleString('en-US');
}
function fmtHcDiff(n) {
  if (!n || isNaN(n) || n === 0) return '—';
  return Math.abs(Math.round(n)).toLocaleString('en-US');
}
function wFmt(n) { return 'W' + String(n).padStart(2,'0'); }
function recargar() { window.location.reload(); }

// =======================================================
// DATA HELPERS
// =======================================================
function getActiveYears() { return DATA.years.filter(function(y){ return state.activeYears[y]; }); }
function getWeekDetail(cat, wn, yr) {
  return DATA.weekly_detail.filter(function(r){ return r.categoria===cat && r.week===wn && r.year===yr; });
}
function ranchFieldName(ranch) { return 'r_' + ranch.replace(/[^a-zA-Z0-9]/g,'_'); }
function fieldToRanch(fn) {
  if (!fn) return null;
  for (var i=0;i<RANCH_ORDER.length;i++) { if (ranchFieldName(RANCH_ORDER[i])===fn) return RANCH_ORDER[i]; }
  return null;
}
function monthFromRecord(rec) {
  var dr = String(rec.date_range||'').toLowerCase();
  var m  = {enero:1,febrero:2,marzo:3,abril:4,mayo:5,junio:6,julio:7,agosto:8,septiembre:9,setiembre:9,octubre:10,noviembre:11,diciembre:12,jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  for (var k in m) { if (dr.indexOf(k)!==-1) return m[k]; }
  var wk = parseInt(rec.week||1); if (!wk||wk<1) wk=1;
  return Math.max(1, Math.min(12, Math.ceil(wk/4.35)));
}
function sumDetail(recs, currency) {
  var out = {total:0, ranches:{}};
  recs.forEach(function(r){
    var v = currency==='usd' ? r.usd_total : r.mxn_total;
    out.total += v;
    var rsrc = currency==='usd' ? r.usd_ranches : r.mxn_ranches;
    Object.keys(rsrc||{}).forEach(function(rn){ out.ranches[rn]=(out.ranches[rn]||0)+rsrc[rn]; });
    if (r.date_range) out.date_range = r.date_range;
  });
  return out;
}

// =======================================================
// CELL RENDERERS (devuelven HTML string)
// =======================================================
function moneyRenderer(p) {
  var v = p.value;
  if (v===null||v===undefined||v===0||isNaN(v)) return '';
  return '<span class="cell-navy">' + fmt(v) + '</span>';
}
function deltaRenderer(p) {
  var v = p.value;
  if (v===null||v===undefined||isNaN(v)) return '';
  if (Math.abs(v)<0.5) return '<span style="color:#999">~0%</span>';
  var cl = v>0 ? 'cell-pos' : 'cell-neg';
  var ar = v>0 ? '&#9650;' : '&#9660;';
  return '<span class="'+cl+'">'+ar+' '+Math.abs(v).toFixed(1)+'%</span>';
}
function deltaAmtRenderer(p) {
  var v = p.value;
  if (!v||isNaN(v)) return '';
  var cl = v>0?'cell-pos':'cell-neg';
  var sign = v>0?'+':'';
  return '<span class="'+cl+'">'+sign+fmt(v)+'</span>';
}
function catRenderer(p) {
  var v = p.value; if (!v) return '';
  return '<span style="font-weight:700;color:#1e3a5f;font-size:11px">'+v+'</span>';
}
function ranchRenderer(ranch) {
  return function(p) {
    var v = p.value;
    if (!v||isNaN(v)||v===0) return '';
    return '<span style="color:#334155;font-weight:600">'+fmt(v)+'</span>';
  };
}

// =======================================================
// RENDER PIVOT TABLE
// =======================================================
function renderPivotTable(colDefs, rows, statusText) {
  _tableColDefs = colDefs;
  _tableRows    = rows;

  // Detectar qué columnas son pinned (las primeras hasta que termine la racha pinned)
  var pinnedCount = 0;
  for (var pi=0; pi<colDefs.length; pi++) {
    if (colDefs[pi].pinned === 'left') pinnedCount++;
    else break;
  }

  // Head
  var headHtml = '<tr>';
  colDefs.forEach(function(col, ci) {
    var align = col.type==='numericColumn' ? 'text-align:right' : 'text-align:left';
    var pinnedCls = ci < pinnedCount ? ' pt-pinned' : '';
    // Calcular left offset para múltiples columnas pinned
    var leftOff = 0;
    if (ci < pinnedCount) {
      for (var px=0; px<ci; px++) leftOff += (colDefs[px].width || 120);
    }
    var leftStyle = ci < pinnedCount ? ';left:'+leftOff+'px' : '';
    headHtml += '<th class="'+pinnedCls+'" style="'+align+leftStyle+'">'+(col.headerName||'')+'</th>';
  });
  headHtml += '</tr>';

  // Body
  var bodyHtml = '';
  var expandedRow = typeof window._expandedRow === 'number' ? window._expandedRow : null;
  rows.forEach(function(row, ri) {
    var rowCls = 'pt-row';
    if      (row._isGroup) rowCls = 'pt-row-group';
    else if (row._isSub)   rowCls = 'pt-row-sub';
    else if (row._isTotal) rowCls = 'pt-row-total';

    bodyHtml += '<tr class="'+rowCls+'" data-ri="'+ri+'">';
    colDefs.forEach(function(col, ci) {
      var val  = row[col.field];
      var align = col.type==='numericColumn' ? 'text-align:right' : 'text-align:left';
      var pinnedCls = ci < pinnedCount ? ' pt-pinned' : '';
      var leftOff = 0;
      if (ci < pinnedCount) {
        for (var px=0; px<ci; px++) leftOff += (colDefs[px].width || 120);
      }
      var leftStyle = ci < pinnedCount ? ';left:'+leftOff+'px' : '';
      var html;
      if (col.cellRenderer) {
        try { html = col.cellRenderer({value:val, data:row, colDef:col}); }
        catch(e) { html = (val===null||val===undefined)?'':String(val); }
      } else {
        html = (val===null||val===undefined)?'':String(val);
      }
      bodyHtml += '<td class="'+pinnedCls+'" style="'+align+leftStyle+'" data-ci="'+ci+'">'+html+'</td>';
    });
    bodyHtml += '</tr>';

    // Si esta fila está expandida, insertar subfila con detalles
    if (expandedRow === ri) {
      var detailHtml = renderRowDetail(row);
      bodyHtml += '<tr class="pt-row-detail"><td colspan="'+colDefs.length+'">'+detailHtml+'</td></tr>';
    }
  });

  var wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '<table class="pt-table"><thead>'+headHtml+'</thead><tbody>'+bodyHtml+'</tbody></table>';

  if (statusText !== undefined) document.getElementById('stTotal').textContent = statusText;
}

// Delegated click sobre tableWrap para expansión de fila
document.addEventListener('click', function(e) {
  var td = e.target.closest('#tableWrap td');
  if (!td) return;
  var tr = td.closest('tr');
  var ri = parseInt(tr.dataset.ri);
  if (isNaN(ri)) return;
  // Alternar expansión
  if (window._expandedRow === ri) {
    window._expandedRow = null;
  } else {
    window._expandedRow = ri;
  }
  renderPivotTable(_tableColDefs, _tableRows);
});

// Renderiza los detalles de la fila expandida (charolas, metros, etc.)
function renderRowDetail(row) {
  // Aquí puedes personalizar los detalles a mostrar
  var html = '<div style="padding:10px 18px 10px 18px; background:#f8fafc; border-left:4px solid #2563eb; font-size:13px; color:#222;">';
  html += '<b>Detalles:</b><br>';
  if (row.charolas_sembradas !== undefined) html += 'Charolas sembradas: <b>' + row.charolas_sembradas + '</b><br>';
  if (row.metros !== undefined) html += 'Metros: <b>' + row.metros + '</b><br>';
  // Agrega aquí más campos según tu modelo de datos
  // Ejemplo:
  if (row.costo_mano_obra !== undefined) html += 'Costo Mano de Obra: <b>' + fmt(row.costo_mano_obra) + '</b><br>';
  if (row.costo_servicios !== undefined) html += 'Costo Servicios: <b>' + fmt(row.costo_servicios) + '</b><br>';
  html += '</div>';
  return html;
}

// =======================================================
// INICIALIZAR
// =======================================================
function inicializar() {
  // prod-link handler
  if (!window._prodLinkBound) {
    document.addEventListener('click', function(e){
      var el = e.target.closest('.prod-link');
      if (!el) return;
      var row = {
        _cat:      decodeURIComponent(el.dataset.cat||''),
        _year:     parseInt(el.dataset.year||'0',10),
        _week:     parseInt(el.dataset.week||'0',10),
        _fromWeek: parseInt(el.dataset.from||el.dataset.week||'0',10),
        _toWeek:   parseInt(el.dataset.to  ||el.dataset.week||'0',10),
      };
      var ranch = decodeURIComponent(el.dataset.ranch||'');
      showProdPanel(row, {ranch: ranch||null});
      e.stopPropagation(); e.preventDefault();
    });
    window._prodLinkBound = true;
  }

  var prefCat = 'MATERIAL DE EMPAQUE';
  state.cat = DATA.categories.indexOf(prefCat)>-1 ? prefCat : DATA.categories[0];

  // Asegurar que los años del conteo de mano de obra también estén activos
  if (Array.isArray(DATA.mano_obra_data)) {
    DATA.mano_obra_data.forEach(function(r){
      if (DATA.years.indexOf(r.year)<0) DATA.years.push(r.year);
    });
    DATA.years.sort(function(a,b){return a-b;});
  }

  state.activeYears = {};
  var latestYr = DATA.years[DATA.years.length-1];
  var prevYr   = DATA.years[DATA.years.length-2];
  if (latestYr) state.activeYears[latestYr] = true;
  if (prevYr)   state.activeYears[prevYr]   = true;

  var wSet = {};
  DATA.weekly_detail.forEach(function(r){ wSet[r.week]=1; });
  // Incluir semanas del conteo de mano de obra
  if (Array.isArray(DATA.mano_obra_data)) {
    DATA.mano_obra_data.forEach(function(r){ wSet[r.week]=1; });
  }
  allWeeks = Object.keys(wSet).map(Number).sort(function(a,b){return a-b;});

  // Semanas del año más reciente — fusionar weekly_detail + mano_obra_data
  // para que semanas WK≥14 (que sólo existen en mano_obra_data) queden incluidas
  // en state.toWeek y no sean filtradas por renderManoObra.
  var _wkLatestSet = {};
  DATA.weekly_detail
    .filter(function(r){return r.year===latestYr;})
    .forEach(function(r){_wkLatestSet[r.week]=1;});
  if (Array.isArray(DATA.mano_obra_data)) {
    DATA.mano_obra_data
      .filter(function(r){return r.year===latestYr;})
      .forEach(function(r){_wkLatestSet[r.week]=1;});
  }
  var wksLatest = Object.keys(_wkLatestSet).map(Number).sort(function(a,b){return a-b;});
  var curWeek = wksLatest[wksLatest.length-1] || allWeeks[allWeeks.length-1];
  var idx = allWeeks.indexOf(curWeek);
  state.weekIdx = idx>=0 ? idx : allWeeks.length-1;

  state.toWeek   = wksLatest[wksLatest.length-1] || allWeeks[allWeeks.length-1] || 52;
  state.fromWeek = wksLatest[wksLatest.length-2] || wksLatest[0] || state.toWeek;

  buildCatSelect();
  buildRanchCheckboxes();
  buildYearChips();
  updateWeekControls();
  updateRangeSliders();
  // Ocultar tab Costo Servicios si la cat inicial no es de tipo servicio
  var vtSrv = document.getElementById('vtServicios');
  var _isSrvCat = (state.cat === 'COSTO SERVICIOS' || state.cat === 'COSTO MANO DE OBRA');
  if (vtSrv) vtSrv.style.display = _isSrvCat ? '' : 'none';

  // Ocultar/Mostrar Boton "Ver Productos"
  var btnProd = document.getElementById('btnVerProd');
  var sepProd = document.getElementById('sepVerProd');
  if (btnProd && sepProd) {
    if (_isSrvCat) { btnProd.style.display = 'none'; sepProd.style.display = 'none'; }
    else { btnProd.style.display = ''; sepProd.style.display = ''; }
  }

  renderView();

  document.getElementById('loader').style.display = 'none';
  document.getElementById('app').style.display    = 'block';
  setTimeout(resizeTable, 80);
  setTimeout(resizeTable, 300);
}

// =======================================================
// UI BUILDERS
// =======================================================
function buildCatSelect() {
  var el = document.getElementById('catSel');
  el.innerHTML = DATA.categories.map(function(c){
    var hasData = categoryHasData(c);
    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+' style="color:'+(hasData?'#222':'#dc2626')+';font-weight:'+(hasData?'400':'700')+';">'+c+'</option>';
  }).join('');
  el.style.color = categoryHasData(state.cat) ? '#222' : '#dc2626';
}
function buildRanchCheckboxes() {
  var panel = document.getElementById('ranchDropdownPanel');
  if (!panel) return;
  var ranches = ['Todos'].concat(RANCH_ORDER);
  panel.innerHTML = ranches.map(function(r) {
    var checked = state.activeRanches.indexOf(r) > -1 ? 'checked' : '';
    var sep = (r === RANCH_ORDER[0]) ? 'border-top:1px solid #eee;margin-top:2px;padding-top:4px;' : '';
    return '<label style="display:flex;align-items:center;gap:6px;padding:4px 10px;cursor:pointer;font-size:11px;white-space:nowrap;' + sep + '">'
      + '<input type="checkbox" value="' + r + '" ' + checked + ' onchange="toggleRanch(this)" style="margin:0;cursor:pointer;">'
      + r + '</label>';
  }).join('');
  var btn = document.getElementById('ranchDropdownBtn');
  if (btn) {
    var sel = state.activeRanches;
    var label = sel.indexOf('Todos') > -1 ? 'Todos' : (sel.length === 1 ? sel[0] : sel.length + ' ranchos');
    btn.textContent = label + ' \u25BE';
  }
}
function toggleRanchDropdown(e) {
  e.stopPropagation();
  var panel = document.getElementById('ranchDropdownPanel');
  if (!panel) return;
  var isOpen = panel.style.display === 'block';
  if (isOpen) {
    panel.style.display = 'none';
    return;
  }
  // Posicionar el panel usando fixed para escapar del overflow del toolbar
  var btn = document.getElementById('ranchDropdownBtn');
  var rect = btn.getBoundingClientRect();
  panel.style.position = 'fixed';
  panel.style.top  = (rect.bottom + 2) + 'px';
  panel.style.left = rect.left + 'px';
  panel.style.display = 'block';
}
document.addEventListener('click', function(e) {
  var btn = document.getElementById('ranchDropdownBtn');
  var panel = document.getElementById('ranchDropdownPanel');
  if (panel && btn && !btn.contains(e.target) && !panel.contains(e.target)) {
    panel.style.display = 'none';
  }
});
function buildYearChips() {
  var el = document.getElementById('yearChips');
  el.innerHTML = DATA.years.map(function(y){
    var col = YEAR_COLORS[y]||'#888';
    var on  = state.activeYears[y] ? ' on' : '';
    return '<button class="yr-chip'+on+'" id="yrChip'+y+'" style="color:'+col+';border-color:'+(state.activeYears[y]?col:'transparent')+';background:'+(state.activeYears[y]?col+'20':'transparent')+'" onclick="toggleYear('+y+')">'+y+'</button>';
  }).join('');
}
function updateWeekControls() {
  // weekSlider y weekLabel eliminados del toolbar — no-op
}
function rowHasData(rec) {
  if (!rec) return false;
  return Math.abs(rec.mxn_total||0) > 0 || Math.abs(rec.usd_total||0) > 0 || Math.abs(rec.hc_total||0) > 0;
}
function categoryHasData(cat) {
  var years = getActiveYears();
  var fromW = state.fromWeek || 1;
  var toW   = state.toWeek   || 52;
  if (!cat || !years.length) return false;
  function inScope(rec) {
    return years.indexOf(rec.year) > -1 && rec.week >= fromW && rec.week <= toW;
  }
  if (cat === 'COSTO MANO DE OBRA') {
    return (DATA.mano_obra_data || []).some(function(rec){ return inScope(rec) && rowHasData(rec); });
  }
  if (cat === 'COSTO SERVICIOS') {
    if ((DATA.servicios_data || []).some(function(rec){ return inScope(rec) && rowHasData(rec); })) return true;
  }
  return (DATA.weekly_detail || []).some(function(rec){ return rec.categoria === cat && inScope(rec) && rowHasData(rec); });
}

// =======================================================
// EVENTS
// =======================================================
function onRanchChange(val) {
  state.ranch = val;
  buildRanchCheckboxes();
  if (state.view !== 'semana') renderView();
}
function onCatChange(val) {
  state.cat = val;
  buildCatSelect();
  var isSrvCat = (val === 'COSTO SERVICIOS' || val === 'COSTO MANO DE OBRA');
  ['Comparativo','Rancho'].forEach(function(name) {
    var el = document.getElementById('vt' + name);
    if (el) el.style.display = isSrvCat ? 'none' : '';
  });
  var vtSrv = document.getElementById('vtServicios');
  if (vtSrv) vtSrv.style.display = isSrvCat ? '' : 'none';

  var btnProd = document.getElementById('btnVerProd');
  var sepProd = document.getElementById('sepVerProd');
  if (btnProd && sepProd) {
    if (isSrvCat) { btnProd.style.display = 'none'; sepProd.style.display = 'none'; }
    else { btnProd.style.display = ''; sepProd.style.display = ''; }
  }

  if (isSrvCat && state.view !== 'servicios') {
    _prodViews = []; _prodViewsData = [];
    setView('servicios');
  } else if (!isSrvCat && state.view === 'servicios') {
    setView('comparativo');
  } else {
    renderView();
  }
}
function setCurrency(cur) {
  state.currency=cur;
  document.getElementById('btnUSD').className='tb-btn'+(cur==='usd'?' active':'');
  document.getElementById('btnMXN').className='tb-btn'+(cur==='mxn'?' active':'');
  renderView();
  // Re-renderizar paneles de detalle solo si la categoría actual los soporta
  var isSrvCat = (state.cat === 'COSTO SERVICIOS' || state.cat === 'COSTO MANO DE OBRA');
  if (_prodViewsData.length > 0 && !isSrvCat) {
    _prodViews = [];
    var _savedData = _prodViewsData.slice();
    _prodViewsData = [];
    _savedData.forEach(function(d){ showProdPanel(d.rowData, d.opts); });
  }
}
function toggleYear(y) {
  var active = DATA.years.filter(function(yr){return state.activeYears[yr];});
  if (state.activeYears[y]&&active.length>1) delete state.activeYears[y];
  else state.activeYears[y]=true;
  buildYearChips();
  buildCatSelect();
  renderView();
}
function toggleRanch(cb) {
  var val = cb.value;
  var idx = state.activeRanches.indexOf(val);
  if (val === 'Todos') {
    // Si se marca "Todos", desmarcar todos los individuales
    if (cb.checked) { state.activeRanches = ['Todos']; }
    else            { state.activeRanches = RANCH_ORDER.slice(); }
  } else {
    // Si se marca un rancho individual, quitar "Todos"
    var todosIdx = state.activeRanches.indexOf('Todos');
    if (todosIdx > -1) state.activeRanches.splice(todosIdx, 1);
    if (cb.checked) {
      if (idx === -1) state.activeRanches.push(val);
    } else {
      if (idx > -1) state.activeRanches.splice(idx, 1);
      // Si no queda ninguno seleccionado, volver a "Todos"
      if (state.activeRanches.length === 0) state.activeRanches = ['Todos'];
    }
  }
  buildRanchCheckboxes();
  renderView();
}
function prevWeek() { if (state.weekIdx>0){state.weekIdx--;updateWeekControls();renderView();} }
function nextWeek() { if (state.weekIdx<allWeeks.length-1){state.weekIdx++;updateWeekControls();renderView();} }
function onWeekSlider(val) {
  var wn=parseInt(val), idx=allWeeks.indexOf(wn);
  if (idx<0){ idx=0; var mn=Math.abs(allWeeks[0]-wn); allWeeks.forEach(function(w,i){var d=Math.abs(w-wn);if(d<mn){mn=d;idx=i;}});}
  state.weekIdx=idx; updateWeekControls(); renderView();
}
function setView(v) {
  state.view=v;
  ['comparativo','rancho','servicios'].forEach(function(name){
    var el=document.getElementById('vt'+name.charAt(0).toUpperCase()+name.slice(1));
    if(el) el.className='vtab'+(v===name?' active':'');
  });
  var rb=document.getElementById('rangeBar');
  if (rb) rb.style.display='none';
  var gw =document.getElementById('gridWrap');
  var cmp=document.getElementById('comparativoWrap');
  if (v==='comparativo') { if(gw)gw.style.display='none'; if(cmp)cmp.className='show'; }
  else                   { if(gw)gw.style.display='';     if(cmp)cmp.className=''; }
  closeProdPanel();
  renderView();
}
function exportCSV() {
  if (!_tableColDefs.length||!_tableRows.length) return;
  var cols = _tableColDefs;
  var lines = [cols.map(function(c){return '"'+(c.headerName||'').replace(/"/g,'""')+'"';}).join(',')];
  _tableRows.forEach(function(row){
    lines.push(cols.map(function(c){
      var v = row[c.field];
      if (v===null||v===undefined) return '';
      return '"'+String(v).replace(/"/g,'""')+'"';
    }).join(','));
  });
  var blob = new Blob([lines.join('\\n')], {type:'text/csv;charset=utf-8;'});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href=url; a.download='CFBC_'+state.view+'_'+new Date().toISOString().slice(0,10)+'.csv';
  a.click(); URL.revokeObjectURL(url);
}
function updateRangeSliders() {
  var f=state.fromWeek, t=state.toWeek;
  var fEl=document.getElementById('fromSlider'), tEl=document.getElementById('toSlider');
  var mn=allWeeks[0]||1, mx=allWeeks[allWeeks.length-1]||52;
  if(fEl){fEl.min=mn;fEl.max=mx;fEl.value=f;}
  if(tEl){tEl.min=mn;tEl.max=mx;tEl.value=t;}
  var yy=String(DATA.years[DATA.years.length-1]).slice(2);
  var fLbl=document.getElementById('fromWeekLabel'), tLbl=document.getElementById('toWeekLabel'), badge=document.getElementById('rangeBadge');
  if(fLbl)fLbl.textContent=String(f).padStart(2,'0');
  if(tLbl)tLbl.textContent=String(t).padStart(2,'0');
  var count=allWeeks.filter(function(w){return w>=f&&w<=t;}).length;
  if(badge)badge.textContent=String(f).padStart(2,'0')+' → '+String(t).padStart(2,'0')+' ('+count+' sem)';
}
function onRangeChange() {
  var f=parseInt(document.getElementById('fromSlider').value);
  var t=parseInt(document.getElementById('toSlider').value);
  if (f>t){var tmp=f;f=t;t=tmp;}
  state.fromWeek=f; state.toWeek=t;
  updateRangeSliders();
  buildCatSelect();
  renderView();
}
function resetRange() {
  var latestYr=DATA.years[DATA.years.length-1];
  var wks=DATA.weekly_detail.filter(function(r){return r.year===latestYr;}).map(function(r){return r.week;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort(function(a,b){return a-b;});
  state.toWeek   = wks[wks.length-1]||allWeeks[allWeeks.length-1]||52;
  state.fromWeek = wks[wks.length-2]||wks[0]||state.toWeek;
  updateRangeSliders();
  buildCatSelect();
  if (state.view==='comparativo') renderComparativo();
}

// =======================================================
// VIEW ROUTER
// =======================================================
function renderView() {
  document.getElementById('prodPanel').className='';
  if (state.view==='comparativo') renderComparativo();
  else if (state.view==='rancho')      renderRancho();
  else if (state.view==='servicios') {
    if (state.cat==='COSTO MANO DE OBRA') renderManoObra();
    else renderServicios();
  }
}

// =======================================================
// VIEW 1: SEMANA
// =======================================================
function renderSemana() {
  var yrs=getActiveYears(), wn=allWeeks[state.weekIdx]||1, sym=state.currency.toUpperCase();
  var cols = [
    { field:'year', headerName:'AÑO', width:70, pinned:'left',
      cellRenderer:function(p){ var c=YEAR_COLORS[p.value]||'#888'; return '<span style="color:'+c+';font-weight:700">'+p.value+'</span>'; }},
    { field:'week', headerName:'SEM', width:60, type:'numericColumn', pinned:'left',
      cellRenderer:function(p){ return wFmt(p.value); }},
    { field:'cat_label', headerName:'CATEGORÍA', width:170, pinned:'left', cellRenderer:catRenderer },
    { field:'total',    headerName:'TOTAL '+sym, width:110, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'deltaAmt', headerName:'Δ $',        width:90,  type:'numericColumn', cellRenderer:deltaAmtRenderer },
    { field:'deltaPct', headerName:'Δ %',        width:72,  type:'numericColumn', cellRenderer:deltaRenderer },
  ];
  RANCH_ORDER.forEach(function(r){ cols.push({field:ranchFieldName(r),headerName:r,width:100,type:'numericColumn',cellRenderer:ranchRenderer(r)}); });

  var rows=[]; var grandTotal=0;
  yrs.forEach(function(yr,i){
    var prevYr=i>0?yrs[i-1]:null;
    var recs=getWeekDetail(state.cat,wn,yr), agg=sumDetail(recs,state.currency);
    var row={year:yr,week:wn,cat_label:state.cat,_cat:state.cat,_year:yr,_week:wn,_fromWeek:wn,_toWeek:wn};
    row.total=agg.total;
    if (prevYr){ var aggP=sumDetail(getWeekDetail(state.cat,wn,prevYr),state.currency); row.deltaAmt=agg.total-aggP.total; row.deltaPct=aggP.total>0?(agg.total-aggP.total)/aggP.total*100:null; }
    RANCH_ORDER.forEach(function(r){ row[ranchFieldName(r)]=agg.ranches[r]||0; });
    rows.push(row);
    if (yr===yrs[yrs.length-1]) grandTotal+=agg.total;
  });
  renderPivotTable(cols, rows, fmt(grandTotal)+' '+sym+' · AÑO '+yrs[yrs.length-1]);
}

// =======================================================
// VIEW 2: ANUAL
// =======================================================
function renderAnual() {
  var yrs=getActiveYears(), sym=state.currency.toUpperCase();

  // Columnas: CATEGORÍA fija + una columna de total por año
  var cols=[
    { field:'cat_label', headerName:'CATEGORÍA', width:220, pinned:'left', cellRenderer:catRenderer },
  ];
  yrs.forEach(function(yr){
    var col=YEAR_COLORS[yr]||'#888';
    cols.push({
      field:'v'+yr, headerName:String(yr)+' '+sym, width:130, type:'numericColumn',
      cellRenderer:(function(c){ return function(p){
        var v=p.value; if(!v||isNaN(v)||v===0) return '';
        return '<span style="color:#1e3a5f;font-weight:600">'+fmt(v)+'</span>';
      };})(col)
    });
  });
  if(yrs.length>=2){
    cols.push({field:'deltaPct',headerName:'Δ %',width:72,type:'numericColumn',cellRenderer:deltaRenderer});
  }

  var getYrTotal=function(cat,yr){
    var d=(DATA.summary[cat]||{})[yr]||{usd:0,mxn:0,ranches:{},ranches_mxn:{}};
    return state.currency==='usd'?d.usd:d.mxn;
  };

  var cats=DATA.categories;
  var rows=[];
  var grandByYr={};
  yrs.forEach(function(yr){ grandByYr[yr]=0; });

  cats.forEach(function(cat){
    var row={cat_label:cat,_cat:cat};
    var hasAny=false;
    yrs.forEach(function(yr){
      var v=getYrTotal(cat,yr);
      row['v'+yr]=v;
      grandByYr[yr]+=v;
      if(v) hasAny=true;
    });
    if(!hasAny) return;
    if(yrs.length>=2){
      var cur=getYrTotal(cat,yrs[yrs.length-1]);
      var prev=getYrTotal(cat,yrs[yrs.length-2]);
      row.deltaPct=prev>0?(cur-prev)/prev*100:null;
    }
    rows.push(row);
  });

  // Fila TOTAL GENERAL
  var totRow={cat_label:'TOTAL GENERAL',_isTotal:true};
  yrs.forEach(function(yr){ totRow['v'+yr]=grandByYr[yr]; });
  if(yrs.length>=2){
    var tCur=grandByYr[yrs[yrs.length-1]];
    var tPrev=grandByYr[yrs[yrs.length-2]];
    totRow.deltaPct=tPrev>0?(tCur-tPrev)/tPrev*100:null;
  }
  rows.push(totRow);

  renderPivotTable(cols, rows);
}

// =======================================================
// VIEW 3: COMPARATIVO
// =======================================================
// (Botones de grupo removidos)
function fmtMes(dr) {
  if (!dr) return '';
  var MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var lower=dr.toLowerCase();
  for (var i=0;i<MESES.length;i++){
    if (lower.indexOf(MESES[i])>-1){
      var m=MESES[i].charAt(0).toUpperCase()+MESES[i].slice(1);
      var yrMatch=dr.match(/\\b(20\\d{2})\\b/);
      return m+(yrMatch?' '+yrMatch[1]:'');
    }
  }
  return dr;
}
function aggregateRecs(recs) {
  var out={usd:0,mxn:0,ranches:{},ranches_mxn:{},date_range:''};
  recs.forEach(function(r){
    out.usd+=r.usd_total; out.mxn+=r.mxn_total;
    if (r.date_range) out.date_range=r.date_range;
    Object.keys(r.usd_ranches||{}).forEach(function(rn){out.ranches[rn]=(out.ranches[rn]||0)+r.usd_ranches[rn];});
    Object.keys(r.mxn_ranches||{}).forEach(function(rn){out.ranches_mxn[rn]=(out.ranches_mxn[rn]||0)+r.mxn_ranches[rn];});
  });
  out.usd=Math.round(out.usd*100)/100; out.mxn=Math.round(out.mxn*100)/100;
  return out;
}
function getRangeByYear(cat,fromW,toW) {
  var res={};
  getActiveYears().forEach(function(yr){
    var recs=DATA.weekly_detail.filter(function(r){return r.categoria===cat&&r.year===yr&&r.week>=fromW&&r.week<=toW;});
    if (!recs.length) return;
    var ag=aggregateRecs(recs);
    ag.weekly={};
    recs.forEach(function(r){ ag.weekly[r.week]=(ag.weekly[r.week]||0)+(state.currency==='usd'?r.usd_total:r.mxn_total); });
    res[yr]=ag;
  });
  return res;
}
function deltaCellHtml(val,prev) {
  if (prev===null||prev===undefined||prev===0) return '<td class="delta-cell chg-0">&#8212;</td>';
  var diff=val-prev, p=((diff/prev)*100).toFixed(1);
  var cls=diff>0?'chg-pos':diff<0?'chg-neg':'chg-0';
  var sign=diff>0?'+':'';
  return '<td class="delta-cell '+cls+'"><span class="delta-amt">'+sign+fmt(diff)+'</span><span class="delta-pct">'+sign+p+'%</span></td>';
}
function renderComparativo() {
  var f=state.fromWeek,t=state.toWeek,yrs=getActiveYears(),sym=state.currency.toUpperCase();
  var byYear=getRangeByYear(state.cat,f,t);
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});
  var ranchCols=getActiveRanches();
  document.getElementById('cmpStats').innerHTML='';

  var weekData={};
  yrs.forEach(function(yr){
    weekData[yr]={};
    rangeWeeks.forEach(function(w){
      var recs=DATA.weekly_detail.filter(function(r){return r.categoria===state.cat&&r.year===yr&&r.week===w;});
      if (recs.length) weekData[yr][w]=aggregateRecs(recs);
    });
  });

  var head='<tr><th>Semana</th><th>Total '+sym+'</th><th>Δ$ vs sem ant.</th>'+ranchCols.map(function(r){return '<th>'+r+'</th>';}).join('')+'</tr>';
  var body=yrs.map(function(yr,yi){
    var col=YEAR_COLORS[yr]||'#888';
    var prevWkVal=null;
    return rangeWeeks.map(function(w){
      var d=weekData[yr][w];
      var val=0;
      if(d){
        if(state.activeRanches.indexOf('Todos')>-1) { val=state.currency==='usd'?d.usd:d.mxn; } else { val=0; var srcA=state.currency==='usd'?d.ranches:d.ranches_mxn; getActiveRanches().forEach(function(rn){ val += (srcA[rn]||0); }); }
      }
      var dCell=deltaCellHtml(val,prevWkVal);
      if (val>0) prevWkVal=val;
      var ranchCells=ranchCols.map(function(r){
        if (!d) return '<td></td>';
        var src=state.currency==='usd'?d.ranches:d.ranches_mxn;
        var v=src[r]||0;
        var style='color:'+(v>0?'#334155':'#ddd')+(v>0?';cursor:pointer;font-weight:600':'');
        var attrs=v>0?' class="cmp-clickable" data-yr="'+yr+'" data-wk="'+w+'" data-ranch="'+r+'"':'';
        return '<td style="'+style+'"'+attrs+'>'+(v>0?fmt(v):'')+'</td>';
      }).join('');
      var totalStyle='color:'+(val>0?'#1e3a5f':'#bbb')+';font-weight:'+(val>0?'700':'400')+(val>0?';cursor:pointer':'');
      var totalAttrs=val>0?' class="cmp-clickable" data-yr="'+yr+'" data-wk="'+w+'" data-ranch=""':'';
      return '<tr class="cmp-row">'+
        '<td style="color:'+col+';font-weight:600">'+String(yr).slice(2)+String(w).padStart(2,'0')+'</td>'+
        '<td style="'+totalStyle+'"'+totalAttrs+'>'+fmt(val)+'</td>'+
        dCell+ranchCells+'</tr>';
    }).join('');
  }).join('');

  document.getElementById('cmpHead').innerHTML=head;
  document.getElementById('cmpBody').innerHTML=body;
  var grandTotal=yrs.reduce(function(s,yr){
    var d=byYear[yr];
    if(!d) return s;
    if(state.activeRanches.indexOf('Todos')>-1) { return s+(state.currency==='usd'?d.usd:d.mxn); } var stot=0; var srcB=state.currency==='usd'?d.ranches:d.ranches_mxn; getActiveRanches().forEach(function(rn){ stot += (srcB[rn]||0); }); return s+stot;
  },0);
  document.getElementById('stTotal').textContent=fmt(grandTotal)+' '+sym;
}

// Delegated click para comparativo clickeable
document.addEventListener('click', function(e){
  var td=e.target.closest('td.cmp-clickable');
  if (!td) return;
  showProdFromCmp(parseInt(td.dataset.yr), parseInt(td.dataset.wk), td.dataset.ranch||null);
});

// =======================================================
// VIEW 4: POR RANCHO
// =======================================================
function renderRancho() {
  var yrs=getActiveYears(), sym=state.currency.toUpperCase();
  var f=state.fromWeek, t=state.toWeek;
  var cur=state.currency;
  var activeRanches = getActiveRanches();
  var showTotal = state.activeRanches.indexOf('Todos')>-1;

  var matCats=DATA.categories.filter(function(c){
    return c!=='COSTO SERVICIOS'&&c!=='COSTO MANO DE OBRA';
  });
  var matRecs=(DATA.weekly_detail||[]).filter(function(r){ return matCats.indexOf(r.categoria)>-1; });
  var moRecs=(DATA.weekly_detail||[]).filter(function(r){ return r.categoria==='COSTO MANO DE OBRA'; });
  var svRecs=DATA.servicios_data||[];

  // Semanas en el rango que tienen datos en algún año activo
  var rangeWeeks=allWeeks.filter(function(w){ return w>=f&&w<=t; });
  var nWk=rangeWeeks.length;
  var nYrs=yrs.length;

  // Sumar para un (año, semana) específico
  function sumForYW(records, yr, wk){
    var out={total:0};
    activeRanches.forEach(function(rn){ out[rn]=0; });
    records.forEach(function(r){
      if(r.year!==yr||r.week!==wk) return;
      var ranches=cur==='usd'?r.usd_ranches:r.mxn_ranches;
      var rowTot = 0;
      activeRanches.forEach(function(rn){
        var v = ranches[rn]||0;
        out[rn] += v;
        rowTot += v;
      });
      out.total += rowTot;
    });
    return out;
  }

  // ywData[yr][wk] = {mat, mo, sv, cpv}
  var ywData={};
  yrs.forEach(function(yr){
    ywData[yr]={};
    rangeWeeks.forEach(function(wk){
      var mat=sumForYW(matRecs,yr,wk);
      var mo =sumForYW(moRecs, yr,wk);
      var sv =sumForYW(svRecs, yr,wk);
      var cpv={total:mat.total+mo.total+sv.total};
      activeRanches.forEach(function(rn){ cpv[rn]=(mat[rn]||0)+(mo[rn]||0)+(sv[rn]||0); });
      ywData[yr][wk]={mat:mat,mo:mo,sv:sv,cpv:cpv};
    });
  });

  var CATS=[
    {key:'mat', label:'COSTO DE MATERIALES', fmt:fmt},
    {key:'mo',  label:'COSTO DE MANO DE OBRA', fmt:fmt},
    {key:'sv',  label:'COSTO DE SERVICIOS', fmt:fmt},
    {key:'cpv', label:'TOTAL CPV', fmt:fmt},
  ];

  // Columnas por grupo de rancho/total:
  //   Por cada año: nWk columnas de semana + 1 DIF (si nWk>=2)
  //   + 1 DIF entre años (si nYrs>=2)
  var perYrCols=nWk+(nWk>=2?1:0);
  var nCols=nYrs*perYrCols+(nYrs>=2?1:0);

  var thB='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid #ddd;white-space:nowrap;position:sticky;top:0;z-index:3;';
  var thPin=thB+'left:0;z-index:4;';

  // ── Header nivel 1: CATEGORÍA | rancho (colspan) | TOTAL (colspan) ────────
  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'min-width:220px;text-align:left">CATEGORÍA</th>';
  activeRanches.forEach(function(rn){
    var col=RANCH_COLORS[rn]||'#888';
    h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;color:'+col+';border-left:2px solid #8EA9C1">'+rn+'</th>';
  });
  if(showTotal) h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;border-left:3px solid #7B1C1C;background:#9DC3E6">TOTAL '+sym+'</th>';
  h1+='</tr>';

  // ── Header nivel 2: por año → semanas + DIF, luego DIF años ───────────────
  function subHeaders(){
    var s='';
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      var yy=String(yr).slice(2);
      rangeWeeks.forEach(function(wk,wi){
        var lbl=yy+String(wk).padStart(2,'0');
        var lb=wi===0?'2px solid '+col:'1px solid #ddd';
        s+='<th style="'+thB+'font-size:9px;color:'+col+';min-width:72px;text-align:right;border-left:'+lb+'">'+lbl+'</th>';
      });
      if(nWk>=2){
        s+='<th style="'+thB+'font-size:9px;min-width:62px;text-align:right;border-left:1px solid #aaa;background:#BDD7EE">DIF'+yy+'</th>';
      }
    });
    if(nYrs>=2){
      var y0=String(yrs[0]).slice(2), yn=String(yrs[nYrs-1]).slice(2);
      s+='<th style="'+thB+'font-size:9px;min-width:68px;text-align:right;border-left:2px solid #7B1C1C;background:#9DC3E6">'+yn+'−'+y0+'</th>';
    }
    return s;
  }
  var h2='<tr>';
  activeRanches.forEach(function(){ h2+=subHeaders(); });
  if(showTotal) h2+=subHeaders(); // TOTAL
  h2+='</tr>';

  // ── Celda ─────────────────────────────────────────────────────────────────
  function cell(v, isDif, bold, lb, fmtFn){
    var format = fmtFn || fmt;
    var s='padding:3px 5px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;';
    if(lb) s+='border-left:'+lb+';';
    if(bold) s+='font-weight:700;';
    if(v===null||v===undefined||isNaN(v)||v===0) return '<td style="'+s+'color:#ccc">—</td>';
    var text = format(v);
    if(isDif){
      var cl=v>0?'#16a34a':'#dc2626';
      return '<td style="'+s+'color:'+cl+'">'+(v>0?'+':'')+text+'</td>';
    }
    return '<td style="'+s+'color:#1e3a5f">'+text+'</td>';
  }

  // ── Generar celdas para un rancho o TOTAL (rnKey=null) ────────────────────
  function groupCells(rnKey){
    var s='';
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      rangeWeeks.forEach(function(wk,wi){
        var v=rnKey!==null ? (ywData[yr][wk][cat.key][rnKey]||0) : (ywData[yr][wk][cat.key].total||0);
        var lb=wi===0?'2px solid '+col:'';
        s+=cell(v, false, isCpv, lb, cat.fmt);
      });
      if(nWk>=2){
        var first=rnKey!==null?(ywData[yr][rangeWeeks[0]][cat.key][rnKey]||0):(ywData[yr][rangeWeeks[0]][cat.key].total||0);
        var last =rnKey!==null?(ywData[yr][rangeWeeks[nWk-1]][cat.key][rnKey]||0):(ywData[yr][rangeWeeks[nWk-1]][cat.key].total||0);
        s+=cell(last-first, true, isCpv, '1px solid #aaa', cat.fmt);
      }
    });
    if(nYrs>=2){
      var v0=rnKey!==null?(ywData[yrs[0]][rangeWeeks[nWk-1]][cat.key][rnKey]||0):(ywData[yrs[0]][rangeWeeks[nWk-1]][cat.key].total||0);
      var vn=rnKey!==null?(ywData[yrs[nYrs-1]][rangeWeeks[nWk-1]][cat.key][rnKey]||0):(ywData[yrs[nYrs-1]][rangeWeeks[nWk-1]][cat.key].total||0);
      s+=cell(vn-v0, true, true, '2px solid #7B1C1C', cat.fmt);
    }
    return s;
  }

  // ── Cuerpo ────────────────────────────────────────────────────────────────
  var bodyHtml='';
  var cat, isCpv; // shared with groupCells via closure
  CATS.forEach(function(c,ci){
    cat=c; isCpv=cat.key==='cpv';
    var bgRow=isCpv?'var(--pt-tot-bg)':(ci%2===0?'#fff':'#F7FBFF');
    var catColor=isCpv?'#1e3a5f':'#333';
    var catFw=isCpv?'800':'700';
    bodyHtml+='<tr style="background:'+bgRow+'">';
    bodyHtml+='<td style="padding:3px 8px;position:sticky;left:0;z-index:1;background:'+bgRow+';border-bottom:1px solid #eee;border-right:1px solid #ddd;white-space:nowrap"><span style="color:'+catColor+';font-weight:'+catFw+'">'+cat.label+'</span></td>';
    activeRanches.forEach(function(rn){ bodyHtml+=groupCells(rn); });
    if(showTotal) bodyHtml+=groupCells(null); // TOTAL
    bodyHtml+='</tr>';
  });

  // ── Inyectar en DOM ───────────────────────────────────────────────────────
  var html='<div class="pt-table-wrap" style="overflow-x:auto;overflow-y:visible;"><table class="pt-table"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
  var gw=document.getElementById('gridWrap');
  gw.style.display='';
  gw.innerHTML=html;
  document.getElementById('comparativoWrap').className='';
  
  // ── AGREGAR TABLAS DE COSTOS UNITARIOS ─────────────────────────────────
  renderUnitCostosTallo(ywData, yrs, rangeWeeks, nWk, nYrs, nCols, activeRanches, showTotal);
  renderUnitCostosHa(ywData, yrs, rangeWeeks, nWk, nYrs, nCols, activeRanches, showTotal);
}

// =======================================================
// TABLA: COSTOS UNITARIOS $ / TALLO PROCESADO
// =======================================================
function renderUnitCostosTallo(ywData, yrs, rangeWeeks, nWk, nYrs, nCols, activeRanches, showTotal) {
  var TALLO_CATS=[
    {key:'materiales_tallo', label:'Materiales'},
    {key:'mano_obra_tallo',  label:'Mano de Obra'},
    {key:'servicios_tallo',  label:'Servicios (Fletes)'},
    {key:'cpv_tallo',        label:'Costo de Producción y Ventas'},
    {key:'empaque_tallo',    label:'Material de Empaque'},
    {key:'sanidad_tallo',    label:'Sanidad Vegetal'},
    {key:'fertilizacion_tallo', label:'Fertilización'},
    {key:'mano_obra_prod_tallo', label:'Mano de Obra Prod'},
  ];

  var thB='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid #ddd;white-space:nowrap;position:sticky;top:0;z-index:3;';
  var thPin=thB+'left:0;z-index:4;';
  var fmt2=fmtFull;

  // Header nivel 1
  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'min-width:220px;text-align:left">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=RANCH_COLORS[rn]||'#888';
    h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;color:'+col+';border-left:2px solid #8EA9C1">'+rn+'</th>';
  });
  if(showTotal) h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;border-left:3px solid #7B1C1C;background:#9DC3E6">TOTAL</th>';
  h1+='</tr>';

  // Header nivel 2
  function subHeaders(){
    var s='';
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      var yy=String(yr).slice(2);
      rangeWeeks.forEach(function(wk,wi){
        var lbl=yy+String(wk).padStart(2,'0');
        var lb=wi===0?'2px solid '+col:'1px solid #ddd';
        s+='<th style="'+thB+'font-size:9px;color:'+col+';min-width:72px;text-align:right;border-left:'+lb+'">'+lbl+'</th>';
      });
      if(nWk>=2){
        s+='<th style="'+thB+'font-size:9px;min-width:62px;text-align:right;border-left:1px solid #aaa;background:#BDD7EE">DIF'+yy+'</th>';
      }
    });
    if(nYrs>=2){
      var y0=String(yrs[0]).slice(2), yn=String(yrs[nYrs-1]).slice(2);
      s+='<th style="'+thB+'font-size:9px;min-width:68px;text-align:right;border-left:2px solid #7B1C1C;background:#9DC3E6">'+yn+'−'+y0+'</th>';
    }
    return s;
  }
  var h2='<tr>';
  activeRanches.forEach(function(){ h2+=subHeaders(); });
  if(showTotal) h2+=subHeaders();
  h2+='</tr>';

  // Función celda
  function cell(v, isDif, lb, fmtFn){
    var format = fmtFn || fmt2;
    var s='padding:3px 5px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;';
    if(lb) s+='border-left:'+lb+';';
    if(v===null||v===undefined||isNaN(v)||v===0) return '<td style="'+s+'color:#ccc">—</td>';
    var text = format(v);
    if(isDif){
      var cl=v>0?'#16a34a':'#dc2626';
      return '<td style="'+s+'color:'+cl+'">'+(v>0?'+':'')+text+'</td>';
    }
    return '<td style="'+s+'color:#1e3a5f">'+text+'</td>';
  }

  // Función para generar celdas por rancho
  // Lee de DATA.unit_costs_data (estructura: { code: { ranch|"TOTAL": { key: val } } })
  function groupCells(rnKey, cat){
    var s='';
    var ucData = DATA.unit_costs_data || {};
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      rangeWeeks.forEach(function(wk,wi){
        var code = (yr - 2000) * 100 + wk;
        var wkUC = ucData[code] || ucData[String(code)] || {};
        var rData = rnKey !== null ? (wkUC[rnKey]||{}) : (wkUC['TOTAL']||{});
        var v = rData[cat.key] || 0;
        var lb=wi===0?'2px solid '+col:'';
        s+=cell(v, false, lb, cat.fmt);
      });
      if(nWk>=2){
        var code0 = (yr-2000)*100 + rangeWeeks[0];
        var codeN = (yr-2000)*100 + rangeWeeks[nWk-1];
        var uc0 = ucData[code0] || ucData[String(code0)] || {};
        var ucN = ucData[codeN] || ucData[String(codeN)] || {};
        var r0 = rnKey!==null ? (uc0[rnKey]||{}) : (uc0['TOTAL']||{});
        var rN = rnKey!==null ? (ucN[rnKey]||{}) : (ucN['TOTAL']||{});
        var first = r0[cat.key]||0;
        var last  = rN[cat.key]||0;
        s+=cell(last-first, true, '1px solid #aaa', cat.fmt);
      }
    });
    if(nYrs>=2){
      var c0yr = (yrs[0]-2000)*100 + rangeWeeks[nWk-1];
      var cNyr = (yrs[nYrs-1]-2000)*100 + rangeWeeks[nWk-1];
      var uc0yr = ucData[c0yr] || ucData[String(c0yr)] || {};
      var ucNyr = ucData[cNyr] || ucData[String(cNyr)] || {};
      var r0yr = rnKey!==null ? (uc0yr[rnKey]||{}) : (uc0yr['TOTAL']||{});
      var rNyr = rnKey!==null ? (ucNyr[rnKey]||{}) : (ucNyr['TOTAL']||{});
      var v0 = r0yr[cat.key]||0;
      var vn = rNyr[cat.key]||0;
      s+=cell(vn-v0, true, '2px solid #7B1C1C', cat.fmt);
    }
    return s;
  }

  // Cuerpo de tabla
  var bodyHtml='';
  TALLO_CATS.forEach(function(c,ci){
    var bgRow=ci%2===0?'#fff':'#F7FBFF';
    bodyHtml+='<tr style="background:'+bgRow+'">';
    bodyHtml+='<td style="padding:3px 8px;position:sticky;left:0;z-index:1;background:'+bgRow+';border-bottom:1px solid #eee;border-right:1px solid #ddd;white-space:nowrap"><span style="color:#333;font-weight:700">'+c.label+'</span></td>';
    activeRanches.forEach(function(rn){ bodyHtml+=groupCells(rn, c); });
    if(showTotal) bodyHtml+=groupCells(null, c);
    bodyHtml+='</tr>';
  });

  var html='<div style="margin-top:20px"><h3 style="color:#1e3a5f;font-size:14px;font-weight:800;margin-bottom:10px">COSTOS UNITARIOS $ / TALLO PROCESADO</h3><div class="pt-table-wrap" style="overflow-x:auto;overflow-y:visible;"><table class="pt-table"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div></div>';
  var gw=document.getElementById('gridWrap');
  gw.innerHTML+=html;
}

// =======================================================
// TABLA: COSTOS UNITARIOS $ / HECTÁREA
// =======================================================
function renderUnitCostosHa(ywData, yrs, rangeWeeks, nWk, nYrs, nCols, activeRanches, showTotal) {
  var HA_CATS=[
    {key:'hectareas_ha',     label:'$ / Hectárea (Ha totales)', fmt:fmtHa},
    {key:'materiales_ha',    label:'Materiales'},
    {key:'mano_obra_ha',     label:'Mano de Obra'},
    {key:'servicios_ha',     label:'Servicios (Fletes)'},
    {key:'cpv_ha',           label:'Costo de Producción y Ventas'},
    {key:'empaque_ha',       label:'Material de Empaque'},
    {key:'fertilizacion_ha', label:'Fertilización'},
    {key:'mano_obra_prod_ha',label:'Mano de Obra Prod'},
  ];

  var thB='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid #ddd;white-space:nowrap;position:sticky;top:0;z-index:3;';
  var thPin=thB+'left:0;z-index:4;';
  var fmt2=fmt;

  // Header nivel 1
  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'min-width:220px;text-align:left">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=RANCH_COLORS[rn]||'#888';
    h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;color:'+col+';border-left:2px solid #8EA9C1">'+rn+'</th>';
  });
  if(showTotal) h1+='<th colspan="'+nCols+'" style="'+thB+'text-align:center;border-left:3px solid #7B1C1C;background:#9DC3E6">TOTAL</th>';
  h1+='</tr>';

  // Header nivel 2
  function subHeaders(){
    var s='';
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      var yy=String(yr).slice(2);
      rangeWeeks.forEach(function(wk,wi){
        var lbl=yy+String(wk).padStart(2,'0');
        var lb=wi===0?'2px solid '+col:'1px solid #ddd';
        s+='<th style="'+thB+'font-size:9px;color:'+col+';min-width:72px;text-align:right;border-left:'+lb+'">'+lbl+'</th>';
      });
      if(nWk>=2){
        s+='<th style="'+thB+'font-size:9px;min-width:62px;text-align:right;border-left:1px solid #aaa;background:#BDD7EE">DIF'+yy+'</th>';
      }
    });
    if(nYrs>=2){
      var y0=String(yrs[0]).slice(2), yn=String(yrs[nYrs-1]).slice(2);
      s+='<th style="'+thB+'font-size:9px;min-width:68px;text-align:right;border-left:2px solid #7B1C1C;background:#9DC3E6">'+yn+'−'+y0+'</th>';
    }
    return s;
  }
  var h2='<tr>';
  activeRanches.forEach(function(){ h2+=subHeaders(); });
  if(showTotal) h2+=subHeaders();
  h2+='</tr>';

  // Función celda
  function cell(v, isDif, lb, fmtFn){
    var format = fmtFn || fmt2;
    var s='padding:3px 5px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;';
    if(lb) s+='border-left:'+lb+';';
    if(v===null||v===undefined||isNaN(v)||v===0) return '<td style="'+s+'color:#ccc">—</td>';
    var text = format(v);
    if(isDif){
      var cl=v>0?'#16a34a':'#dc2626';
      return '<td style="'+s+'color:'+cl+'">'+(v>0?'+':'')+text+'</td>';
    }
    return '<td style="'+s+'color:#1e3a5f">'+text+'</td>';
  }

  // Función para generar celdas por rancho
  // Lee de DATA.unit_costs_data (estructura: { code: { ranch|"TOTAL": { key: val } } })
  function groupCells(rnKey, cat){
    var s='';
    var ucData = DATA.unit_costs_data || {};
    yrs.forEach(function(yr,yi){
      var col=YEAR_COLORS[yr]||'#888';
      rangeWeeks.forEach(function(wk,wi){
        var code = (yr - 2000) * 100 + wk;
        var wkUC = ucData[code] || ucData[String(code)] || {};
        var rData = rnKey !== null ? (wkUC[rnKey]||{}) : (wkUC['TOTAL']||{});
        var v = rData[cat.key] || 0;
        var lb=wi===0?'2px solid '+col:'';
        s+=cell(v, false, lb, cat.fmt);
      });
      if(nWk>=2){
        var code0 = (yr-2000)*100 + rangeWeeks[0];
        var codeN = (yr-2000)*100 + rangeWeeks[nWk-1];
        var uc0 = ucData[code0] || ucData[String(code0)] || {};
        var ucN = ucData[codeN] || ucData[String(codeN)] || {};
        var r0 = rnKey!==null ? (uc0[rnKey]||{}) : (uc0['TOTAL']||{});
        var rN = rnKey!==null ? (ucN[rnKey]||{}) : (ucN['TOTAL']||{});
        var first = r0[cat.key]||0;
        var last  = rN[cat.key]||0;
        s+=cell(last-first, true, '1px solid #aaa', cat.fmt);
      }
    });
    if(nYrs>=2){
      var c0yr = (yrs[0]-2000)*100 + rangeWeeks[nWk-1];
      var cNyr = (yrs[nYrs-1]-2000)*100 + rangeWeeks[nWk-1];
      var uc0yr = ucData[c0yr] || ucData[String(c0yr)] || {};
      var ucNyr = ucData[cNyr] || ucData[String(cNyr)] || {};
      var r0yr = rnKey!==null ? (uc0yr[rnKey]||{}) : (uc0yr['TOTAL']||{});
      var rNyr = rnKey!==null ? (ucNyr[rnKey]||{}) : (ucNyr['TOTAL']||{});
      var v0 = r0yr[cat.key]||0;
      var vn = rNyr[cat.key]||0;
      s+=cell(vn-v0, true, '2px solid #7B1C1C', cat.fmt);
    }
    return s;
  }

  // Cuerpo de tabla
  var bodyHtml='';
  HA_CATS.forEach(function(c,ci){
    var bgRow=ci%2===0?'#fff':'#F7FBFF';
    bodyHtml+='<tr style="background:'+bgRow+'">';
    bodyHtml+='<td style="padding:3px 8px;position:sticky;left:0;z-index:1;background:'+bgRow+';border-bottom:1px solid #eee;border-right:1px solid #ddd;white-space:nowrap"><span style="color:#333;font-weight:700">'+c.label+'</span></td>';
    activeRanches.forEach(function(rn){ bodyHtml+=groupCells(rn, c); });
    if(showTotal) bodyHtml+=groupCells(null, c);
    bodyHtml+='</tr>';
  });

  var html='<div style="margin-top:20px"><h3 style="color:#1e3a5f;font-size:14px;font-weight:800;margin-bottom:10px">COSTOS UNITARIOS $ / HECTÁREA</h3><div class="pt-table-wrap" style="overflow-x:auto;overflow-y:visible;"><table class="pt-table"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div></div>';
  var gw=document.getElementById('gridWrap');
  gw.innerHTML+=html;
}
function renderDetalle() {
  var sym=state.currency.toUpperCase();
  var activeRanches = getActiveRanches();
  var cols=[
    { field:'year',      headerName:'AÑO',      width:60,  type:'numericColumn', pinned:'left' },
    { field:'week',      headerName:'SEM',       width:55,  type:'numericColumn', pinned:'left', cellRenderer:function(p){return wFmt(p.value);} },
    { field:'categoria', headerName:'CATEGORÍA', width:220, pinned:'left', cellRenderer:catRenderer },
    { field:'usd_total', headerName:'USD',       width:100, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'mxn_total', headerName:'MXN',       width:110, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'date_range',headerName:'PERÍODO',   width:160,
      cellRenderer:function(p){return '<span style="color:#888;font-size:11px">'+(p.value||'')+'</span>';}},
  ];
  activeRanches.forEach(function(r){
    var c=RANCH_COLORS[r]||'#888';
    cols.push((function(color){return {field:'rn_'+r.replace(/[^a-zA-Z0-9]/g,'_'),headerName:r,width:100,type:'numericColumn',
      cellRenderer:function(p){var v=p.value;if(!v||v<0.01)return '<span class="cell-muted">&#8212;</span>';return '<span style="color:'+color+'">'+fmt(v)+'</span>';}};})(c));
  });
  var rows=[],grandTotal=0;
  DATA.weekly_detail.forEach(function(r){
    if (!state.activeYears[r.year]) return;
    if (r.categoria!==state.cat) return;
    var row={year:r.year,week:r.week,categoria:r.categoria,usd_total:r.usd_total,mxn_total:r.mxn_total,date_range:r.date_range||''};
    activeRanches.forEach(function(rn){var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;row['rn_'+rn.replace(/[^a-zA-Z0-9]/g,'_')]=src[rn]||0;});
    grandTotal+=state.currency==='usd'?r.usd_total:r.mxn_total;
    rows.push(row);
  });
  rows.sort(function(a,b){return b.year!==a.year?b.year-a.year:b.week-a.week;});
  renderPivotTable(cols,rows,fmt(grandTotal)+' '+sym+' ('+rows.length+' registros) · '+state.cat);
}

// =======================================================
// VIEW 6: PRODUCTOS
// =======================================================
function renderProductosFull() {
  var cols=[
    { field:'tipo',      headerName:'TIPO',     width:60,  pinned:'left' },
    { field:'cat',       headerName:'CAT',       width:55,  pinned:'left',
      cellRenderer:function(p){var m={'PR':'#16a34a','MP':'#7c3aed','ME':'#0369a1','MV':'#b45309'};return '<span style="color:'+(m[p.value]||'#666')+';font-weight:700">'+(p.value||'')+'</span>';}},
    { field:'week_code', headerName:'WK',        width:72 },
    { field:'rancho',    headerName:'RANCHO',    width:110,
      cellRenderer:function(p){return '<span style="color:'+(RANCH_COLORS[p.value]||'#666')+';font-weight:600">'+(p.value||'')+'</span>';}},
    { field:'producto',  headerName:'PRODUCTO',  width:260,
      cellRenderer:function(p){return '<span style="color:#1e3a5f">'+(p.value||'')+'</span>';}},
    { field:'unidades',  headerName:'UNID.',     width:80,
      cellRenderer:function(p){return '<span style="color:#555">'+(p.value||'')+'</span>';}},
    { field:'gasto',     headerName:'GASTO',     width:100, type:'numericColumn', cellRenderer:moneyRenderer },
  ];
  var rows=[];
  function flattenProd(dataSet,label){
    if (!dataSet) return;
    Object.keys(dataSet).forEach(function(wkCode){
      var byRanch=dataSet[wkCode];
      Object.keys(byRanch).forEach(function(ranch){
        var byTipo=byRanch[ranch];
        Object.keys(byTipo).forEach(function(tipo){
          var items=byTipo[tipo];
          if (!Array.isArray(items)) return;
          items.forEach(function(item){rows.push({cat:label,tipo:tipo,week_code:parseInt(wkCode)||wkCode,rancho:ranch,producto:item[0]||'',unidades:item[1]||'',gasto:parseFloat(item[2])||0});});
        });
      });
    });
  }
  flattenProd(DATA.productos,'PR'); flattenProd(DATA.productos_mp,'MP'); flattenProd(DATA.productos_me,'ME'); flattenProd(DATA.productos_mv,'MV');
  rows.sort(function(a,b){if(b.week_code!==a.week_code)return (b.week_code||0)-(a.week_code||0);return (a.rancho||'').localeCompare(b.rancho||'');});
  var total=rows.reduce(function(s,r){return s+(r.gasto||0);},0);
  renderPivotTable(cols,rows,fmt(total)+' · '+rows.length+' registros');
}

// =======================================================
// VIEW 7: COSTO SERVICIOS  (con desglose por rancho)
// =======================================================
var SV_SUBCATS=['Electricidad','Fletes y Acarreos','Gastos de Exportación','Certificado Fitosanitario','Transporte de Personal','Compra de Flor a Terceros','Comida para el Personal','RO, TEL, RTA.Alim'];
function renderServicios() {
  var sym=state.currency.toUpperCase();
  var f=state.fromWeek, t=state.toWeek;
  var yrs=getActiveYears();
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});

  // ── Acumular datos ─────────────────────────────────────
  var weekMap={};
  var src=Array.isArray(DATA.servicios_data)&&DATA.servicios_data.length ? DATA.servicios_data : DATA.weekly_detail;
  src.forEach(function(r){
    if (!state.activeYears[r.year]) return;
    if (r.week<f||r.week>t) return;
    var key=r.year+'-'+r.week;
    if (!weekMap[key]) weekMap[key]={_year:r.year,_week:r.week,date_range:r.date_range||''};
    var subcat, val, ranches;
    if (r.subcat) {
      subcat=r.subcat;
      val=state.currency==='usd'?r.usd_total:r.mxn_total;
      ranches=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
    } else if (r.categoria&&r.categoria.startsWith('SV:')) {
      subcat=r.categoria.replace('SV:','');
      val=state.currency==='usd'?r.usd_total:r.mxn_total;
      ranches=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
    } else return;
    if (!subcat) return;
    weekMap[key][subcat]=(weekMap[key][subcat]||0)+(val||0);
    Object.keys(ranches||{}).forEach(function(rn){
      weekMap[key][subcat+'__r__'+rn]=(weekMap[key][subcat+'__r__'+rn]||0)+(ranches[rn]||0);
    });
  });

  // ── Semanas con datos ─────────────────────────────────
  var weekKeys=[];
  yrs.forEach(function(yr){
    rangeWeeks.forEach(function(w){
      var key=yr+'-'+w;
      var hasSome=false;
      Object.keys(weekMap[key]||{}).forEach(function(k){
        if(k[0]!=='_'&&!k.includes('__r__')&&weekMap[key][k]>0) hasSome=true;
      });
      if(hasSome) weekKeys.push(key);
    });
  });

  // ── Subcats ordenadas ─────────────────────────────────
  var subcatsSet={};
  weekKeys.forEach(function(key){
    Object.keys(weekMap[key]||{}).forEach(function(k){
      if(k[0]!=='_'&&!k.includes('__r__')) subcatsSet[k]=1;
    });
  });
  var orderedSubcats=SV_SUBCATS.filter(function(sc){return subcatsSet[sc];});
  Object.keys(subcatsSet).forEach(function(sc){if(orderedSubcats.indexOf(sc)===-1)orderedSubcats.push(sc);});

  // ── Ranchos activos ───────────────────────────────────
  var allowedRanches=getActiveRanches();
  var activeRanches=allowedRanches.filter(function(rn){
    return weekKeys.some(function(key){
      return Object.keys(weekMap[key]||{}).some(function(k){return k.endsWith('__r__'+rn)&&weekMap[key][k]>0;});
    });
  });
  var showTotal = state.activeRanches.indexOf('Todos')>-1;

  // ── Sin datos ─────────────────────────────────────────
  if (!weekKeys.length || !orderedSubcats.length) {
    document.getElementById('gridWrap').style.display='';
    document.getElementById('gridWrap').innerHTML='<div style="padding:20px;color:#888;font-size:12px">Sin datos para el rango seleccionado.</div>';
    document.getElementById('comparativoWrap').className='';
    document.getElementById('stTotal').textContent='—';
    return;
  }

  // ─────────────────────────────────────────────────────
  // NUEVO LAYOUT: columnas = RANCHO, sub-columnas = semanas
  // CONCEPTO | [Isabela: wk1 wk2 ... SUB] | [PosCo: wk1 wk2 ... SUB] | TOTAL
  // ─────────────────────────────────────────────────────
  var nWeeks = weekKeys.length;
  var nColsPerRanch = nWeeks + 1; // semanas + SUB

  // TOTAL: una columna por año (suma del rango) + DIF si hay 2 años
  var totYears = [];
  weekKeys.forEach(function(key){ var yr=weekMap[key]._year; if(totYears.indexOf(yr)<0) totYears.push(yr); });
  totYears.sort(function(a,b){return a-b;});
  var nColsTotal = totYears.length + (totYears.length >= 2 ? 1 : 0);

  var thBase='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid var(--pt-hdr-border);white-space:nowrap;';
  var thPin =thBase+'position:sticky;top:0;z-index:4;';
  var thScroll=thBase+'position:sticky;top:0;z-index:3;text-align:right;';

  // ── Header nivel 1: CONCEPTO | [Ranch colspan=nColsPerRanch] ... | TOTAL group ──
  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'left:0;min-width:190px">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=RANCH_COLORS[rn]||'#888';
    h1+='<th colspan="'+nColsPerRanch+'" style="'+thScroll+'border-left:2px solid #8EA9C1;text-align:center;color:'+col+'">'+rn+'</th>';
  });
  if(showTotal) h1+='<th colspan="'+nColsTotal+'" style="'+thScroll+'border-left:3px solid #7B1C1C;text-align:center;background:#9DC3E6;color:#1e3a5f">TOTAL</th>';
  h1+='</tr>';

  // ── Header nivel 2: por cada rancho → [wk labels... | DIF] + TOTAL group (por año) ──
  var h2='<tr>';
  activeRanches.forEach(function(){
    weekKeys.forEach(function(key){
      var d=weekMap[key];
      var lbl=String(d._year).slice(2)+String(d._week).padStart(2,'0');
      var col=YEAR_COLORS[d._year]||'#888';
      h2+='<th style="'+thScroll+'border-left:1px solid var(--pt-hdr-border);font-size:9px;color:'+col+';min-width:60px">'+lbl+'</th>';
    });
    h2+='<th style="'+thScroll+'border-left:1px solid #aaa;font-size:9px;min-width:70px;background:#BDD7EE">DIF</th>';
  });
  if(showTotal) {
    totYears.forEach(function(yr, i){
      var col=YEAR_COLORS[yr]||'#888';
      var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : 'border-left:1px solid var(--pt-hdr-border);';
      h2+='<th style="'+thScroll+lb+'font-size:9px;color:'+col+';min-width:90px;background:#EAF3FF">'+yr+'</th>';
    });
    if(totYears.length >= 2) {
      h2+='<th style="'+thScroll+'border-left:1px solid #aaa;font-size:9px;min-width:70px;background:#9DC3E6">DIF</th>';
    }
  }
  h2+='</tr>';

  // ── Acumuladores totales ──────────────────────────────
  var grandByRnWk={}, grandByRn={}, grandTotal=0;
  var grandTotWk={};
  activeRanches.forEach(function(rn){
    grandByRnWk[rn]={}; grandByRn[rn]=0;
    weekKeys.forEach(function(k){ grandByRnWk[rn][k]=0; });
  });
  weekKeys.forEach(function(k){ grandTotWk[k]=0; });

  function cell(v,bold,color){
    if(!v||isNaN(v)||v===0) return '<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ccc">—</td>';
    var s='padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;';
    if(bold) s+='font-weight:700;';
    s+='color:'+(color||'#1e3a5f')+';';
    return '<td style="'+s+'">'+fmt(v)+'</td>';
  }

  // ── Filas por subcat ──────────────────────────────────
  var bodyHtml='';
  orderedSubcats.forEach(function(sc){
    var scByRnWk={}, scByRn={}, scTotal=0, scByWk={};
    weekKeys.forEach(function(k){ scByWk[k]=0; });
    activeRanches.forEach(function(rn){
      scByRnWk[rn]={}; scByRn[rn]=0;
      weekKeys.forEach(function(k){
        var v=(weekMap[k]&&weekMap[k][sc+'__r__'+rn])?weekMap[k][sc+'__r__'+rn]:0;
        scByRnWk[rn][k]=v;
        scByRn[rn]+=v;
        scTotal+=v;
        scByWk[k]+=v;
        grandByRnWk[rn][k]+=v;
        grandByRn[rn]+=v;
        grandTotal+=v;
        grandTotWk[k]+=v;
      });
    });
    if(scTotal===0) return;

    var tdPin='padding:3px 8px;position:sticky;z-index:1;background:#fff;border-bottom:1px solid #eee;border-right:1px solid #eee;white-space:nowrap;font-size:11px;';
    bodyHtml+='<tr class="pt-row">';
    var displayLabel = (sc === 'RO, TEL, RTA.Alim') ? 'COMBUSTIBLE' : sc;
    bodyHtml+='<td style="'+tdPin+'left:0;color:#1e3a5f;font-weight:700">'+displayLabel+'</td>';
    activeRanches.forEach(function(rn){
      weekKeys.forEach(function(key){
        var v=scByRnWk[rn][key];
        if(!v||v===0){bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ddd">—</td>';}
        else{bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#334155;font-weight:600">'+fmt(v)+'</td>';}
      });
      var rnDif=(scByRnWk[rn][weekKeys[weekKeys.length-1]]||0)-(scByRnWk[rn][weekKeys[0]]||0);
      bodyHtml+=cell(rnDif||0,true,'#1e3a5f');
    });
    // Celdas TOTAL por año (suma del rango)
    var totCellStyle='padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;background:#EAF3FF;font-weight:700;color:#1e3a5f;';
    if(showTotal) {
      var scByYr={};
      totYears.forEach(function(yr){ scByYr[yr]=0; });
      weekKeys.forEach(function(k){ var yr=weekMap[k]._year; scByYr[yr]=(scByYr[yr]||0)+(scByWk[k]||0); });
      totYears.forEach(function(yr, i){
        var v=scByYr[yr]||0;
        var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
        bodyHtml+='<td style="'+totCellStyle+lb+'">'+( v?fmt(v):'—')+'</td>';
      });
      if(totYears.length >= 2) {
        var scDifYr=(scByYr[totYears[totYears.length-1]]||0)-(scByYr[totYears[0]]||0);
        bodyHtml+='<td style="'+totCellStyle+'border-left:1px solid #aaa;background:#9DC3E6">'+(scDifYr?(scDifYr>0?'+':'')+fmt(scDifYr):'—')+'</td>';
      }
    }
    bodyHtml+='</tr>';
  });

  // ── Fila TOTAL GENERAL ────────────────────────────────
  var totStyle='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;';
  var totPin='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;z-index:2;white-space:nowrap;';
  bodyHtml+='<tr>';
  bodyHtml+='<td style="'+totPin+'left:0;text-align:left">TOTAL GENERAL</td>';
  activeRanches.forEach(function(rn){
    weekKeys.forEach(function(key){
      var v=grandByRnWk[rn][key];
      bodyHtml+='<td style="'+totStyle+'color:#1e3a5f">'+(v?fmt(v):'—')+'</td>';
    });
    var rnTotDif=(grandByRnWk[rn][weekKeys[weekKeys.length-1]]||0)-(grandByRnWk[rn][weekKeys[0]]||0);
    bodyHtml+='<td style="'+totStyle+'color:#1e3a5f;border-left:1px solid #aaa">'+(rnTotDif?fmt(rnTotDif):'—')+'</td>';
  });
  var totTotStyle='padding:4px 8px;background:#9DC3E6;font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;color:#1e3a5f;';
  if(showTotal) {
    var grandTotYr={};
    totYears.forEach(function(yr){ grandTotYr[yr]=0; });
    weekKeys.forEach(function(k){ var yr=weekMap[k]._year; grandTotYr[yr]=(grandTotYr[yr]||0)+(grandTotWk[k]||0); });
    totYears.forEach(function(yr, i){
      var v=grandTotYr[yr]||0;
      var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
      bodyHtml+='<td style="'+totTotStyle+lb+'">'+( v?fmt(v):'—')+'</td>';
    });
    if(totYears.length >= 2) {
      var gtDifYr=(grandTotYr[totYears[totYears.length-1]]||0)-(grandTotYr[totYears[0]]||0);
      bodyHtml+='<td style="'+totTotStyle+'border-left:1px solid #aaa;background:#7EB3D4">'+(gtDifYr?(gtDifYr>0?'+':'')+fmt(gtDifYr):'—')+'</td>';
    }
  }
  bodyHtml+='</tr>';

  // ── Inyectar en el DOM ────────────────────────────────
  var html='<div class="pt-table-wrap" id="tableWrap" style="overflow-x:auto;overflow-y:visible;"><table class="pt-table" style="border-collapse:collapse;width:100%"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
  var gw=document.getElementById('gridWrap');
  if(gw){ gw.style.display=''; gw.innerHTML=html; }
  document.getElementById('comparativoWrap').className='';
  document.getElementById('stTotal').textContent=fmt(grandTotal)+' '+sym;
  setTimeout(resizeTable,80);
}

// =======================================================
// VIEW 8: COSTO MANO DE OBRA
// =======================================================
var MO_SUBCATS = [
  'Ing. Y Admon.','Supervisores','Corte','Trasplante','Manejo P.',
  'Consolidacion','Siembra','Mov. Charolas','Riego','Phlox','Hoops','MIPE Y MIRFE',
  'Tract. Y Cameros','Veladores','Soldadores','Transporte',
  'Admon Posco','Alm.upc y empaq','Contratista y com.',
  'Prod. Patina y rec','IMSS,INFO Y RCV','Imp. 1.8%'
];

// Colores para ranchos del conteo de personal (distintos a los ranchos físicos)
var MO_RANCH_COLORS = {
  'Administracion':  '#374151',
  'Propagacion':     '#047857',
  'Poscosecha':      '#0369a1',
  'Ramona':          '#b45309',
  'Isabela':         '#7c3aed',
  'Christina':       '#0369a1',
  'Cecilia':         '#be185d',
  'Cecilia 25':      '#047857',
};

// Orden preferido de ranchos en mano de obra
var MO_RANCH_ORDER = [
  'Administracion','Propagacion','Poscosecha','Ramona',
  'Isabela','Christina','Cecilia','Cecilia 25'
];

function renderManoObra() {
  var sym=state.currency.toUpperCase();
  var f=state.fromWeek, t=state.toWeek;
  var yrs=getActiveYears();

  // ── Acumular datos ─────────────────────────────────────
  var weekMap={};
  var src=Array.isArray(DATA.mano_obra_data)&&DATA.mano_obra_data.length?DATA.mano_obra_data:[];
  src.forEach(function(r){
    if (!state.activeYears[r.year]) return;
    if (r.week<f||r.week>t) return;
    var key=r.year+'-'+r.week;
    if (!weekMap[key]) weekMap[key]={_year:r.year,_week:r.week,date_range:r.date_range||''};
    var subcat=(r.subcat||'').trim(); if (!subcat) return;
    var val=state.currency==='usd'?r.usd_total:r.mxn_total;
    var ranches=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
    var valHc=r.hc_total||0;
    var hcRanches=r.hc_ranches||{};
    weekMap[key][subcat]=(weekMap[key][subcat]||0)+(val||0);
    Object.keys(ranches||{}).forEach(function(rn){
      weekMap[key][subcat+'__r__'+rn]=(weekMap[key][subcat+'__r__'+rn]||0)+(ranches[rn]||0);
    });
    Object.keys(hcRanches||{}).forEach(function(rn){
      weekMap[key][subcat+'__hc_r__'+rn]=(weekMap[key][subcat+'__hc_r__'+rn]||0)+(hcRanches[rn]||0);
    });
  });

  // MO_GROUPS: un subcat por grupo — igual tanto en conteo.xlsx (sem 1-13)
  // como en el WK Excel (sem 14+), ya que norm_cat() ahora colapsa
  // Nómina + H.Extra + Bonos al mismo nombre unificado del BD CONTEO.
  var MO_GROUPS=[
    {label:'ING. Y ADMON.',      subcats:['Ing. Y Admon.']},
    {label:'SUPERVISORES',       subcats:['Supervisores']},
    {label:'CORTE',              subcats:['Corte']},
    {label:'TRASPLANTE',         subcats:['Trasplante']},
    {label:'MANEJO PLANTA',      subcats:['Manejo P.']},
    {label:'CONSOLIDACIÓN',      subcats:['Consolidacion']},
    {label:'SIEMBRA',            subcats:['Siembra']},
    {label:'MOV. CHAROLAS',      subcats:['Mov. Charolas']},
    {label:'RIEGO',              subcats:['Riego']},
    {label:'PHLOX',              subcats:['Phlox']},
    {label:'HOOPS',              subcats:['Hoops']},
    {label:'MIPE / MIRFE',       subcats:['MIPE Y MIRFE']},
    {label:'TRACTORES/CAMEROS',  subcats:['Tract. Y Cameros']},
    {label:'VELADORES',          subcats:['Veladores']},
    {label:'SOLDADORES',         subcats:['Soldadores']},
    {label:'TRANSPORTE',         subcats:['Transporte']},
    {label:'ADMON POSCO',        subcats:['Admon Posco']},
    {label:'ALM. UPC Y EMPAQUE', subcats:['Alm.upc y empaq']},
    {label:'CONTRATISTA',        subcats:['Contratista y com.']},
    {label:'PROD. PÁTINA Y REC', subcats:['Prod. Patina y rec']},
    {label:'IMSS/INFO/RCV',      subcats:['IMSS,INFO Y RCV']},
    {label:'IMP. 1.8%',          subcats:['Imp. 1.8%']},
  ];

  // Semanas con datos — directamente del weekMap (orden año-semana)
  var weekKeys=Object.keys(weekMap).filter(function(key){
    var hasSome=false;
    Object.keys(weekMap[key]).forEach(function(k){if(k[0]!=='_'&&!k.includes('__r__')&&weekMap[key][k]>0)hasSome=true;});
    return hasSome;
  }).sort(function(a,b){
    var pa=a.split('-'), pb=b.split('-');
    return (parseInt(pa[0])-parseInt(pb[0]))||( parseInt(pa[1])-parseInt(pb[1]));
  });

  // Ranchos con datos: usar los ranchos del conteo (no RANCH_ORDER físico)
  // Primero recolectar todos los ranchos que aparecen en los datos
  var ranchesEnDatos={};
  weekKeys.forEach(function(key){
    Object.keys(weekMap[key]||{}).forEach(function(k){
      if(!k.includes('__r__')) return;
      var rn=k.split('__r__')[1];
      if(weekMap[key][k]>0) ranchesEnDatos[rn]=true;
    });
  });
  // Mapeo de ranchos para Mano de Obra
  var moMap = {
    'Isabela': ['Isabela'],
    'Cecilia': ['Cecilia'],
    'Cecilia 25': ['Cecilia 25'],
    'Prop-RM': ['Propagacion'],
    'Campo-RM': ['Ramona'],
    'PosCo-RM': ['Poscosecha']
  };
  var _ar=getActiveRanches(); var allowedRanches=[]; if(state.activeRanches.indexOf('Todos')>-1){allowedRanches=MO_RANCH_ORDER;}else{_ar.forEach(function(rn){var m=moMap[rn]||[rn]; m.forEach(function(nx){if(allowedRanches.indexOf(nx)<0)allowedRanches.push(nx);});});}
  var activeRanches = allowedRanches.filter(function(rn){ return ranchesEnDatos[rn]; });
  Object.keys(ranchesEnDatos).forEach(function(rn){
    if (state.activeRanches.indexOf('Todos')<0 && allowedRanches.indexOf(rn) < 0) return;
    if (activeRanches.indexOf(rn) < 0) activeRanches.push(rn);
  });
  var showTotal = true;

  function shortLabel(sc){
    return sc.replace('Nómina Prod. ','').replace('Nómina Op. ','')
      .replace('H.Extra Dom. y Festivos ','H.Extra ').replace('H.Extra Dom. y Fest. ','H.Extra ')
      .replace('Bonos Asist./Puntualidad ','Bonos ');
  }
  function cell(v,bold,color){
    if(!v||isNaN(v)||v===0) return '<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ccc">—</td>';
    var s='padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;';
    if(bold) s+='font-weight:700;';
    s+='color:'+(color||'#1e3a5f')+';';
    return '<td style="'+s+'">'+fmt(v)+'</td>';
  }
  function cellGrp(v){
    if(!v||isNaN(v)||v===0) return '<td style="padding:3px 6px;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;background:var(--pt-grp-bg);color:#fff;font-weight:700">—</td>';
    return '<td style="padding:3px 6px;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;background:var(--pt-grp-bg);color:#fff;font-weight:700">'+fmt(v)+'</td>';
  }

  if (!weekKeys.length || !activeRanches.length) {
    var diagHtml = '<div style="padding:16px;font-family:monospace;font-size:11px;background:#fff3cd;border:1px solid #ffc107;margin:10px">';
    diagHtml += '<b>DEBUG</b><br>';
    diagHtml += 'mano_obra_data.length: ' + (Array.isArray(DATA.mano_obra_data) ? DATA.mano_obra_data.length : 'NO ARRAY') + '<br>';
    if (Array.isArray(DATA.mano_obra_data) && DATA.mano_obra_data.length > 0) {
      var s=DATA.mano_obra_data[0];
      diagHtml += 'sample: year='+s.year+' week='+s.week+' subcat="'+s.subcat+'" mxn='+s.mxn_total+'<br>';
      diagHtml += 'mxn_ranches='+JSON.stringify(s.mxn_ranches)+'<br>';
    }
    diagHtml += 'activeYears='+JSON.stringify(state.activeYears)+'<br>';
    diagHtml += 'weekMap keys='+JSON.stringify(Object.keys(weekMap))+'<br>';
    diagHtml += 'activeRanches='+JSON.stringify(activeRanches)+'<br>';
    diagHtml += '</div>';
    var gw=document.getElementById('gridWrap');
    if(gw){gw.style.display='';gw.innerHTML=diagHtml;}
    document.getElementById('comparativoWrap').className='';
    document.getElementById('stTotal').textContent='—';
    return;
  }

  // ── Construir HTML ────────────────────────────────────
  // LAYOUT: CONCEPTO (grupo+subcat colapsable) | [Rancho: wk1 wk2 ... DIF] | TOTAL
  var nWeeks = weekKeys.length;
  var nColsPerRanch = nWeeks + 1; // semanas + SUB

  // TOTAL: una columna por año (suma del rango) + DIF si hay 2 años
  var totYears = [];
  weekKeys.forEach(function(key){ var yr=weekMap[key]._year; if(totYears.indexOf(yr)<0) totYears.push(yr); });
  totYears.sort(function(a,b){return a-b;});
  var nColsTotal = totYears.length + (totYears.length >= 2 ? 1 : 0);

  var thBase='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid var(--pt-hdr-border);white-space:nowrap;';
  var thPin=thBase+'position:sticky;top:0;z-index:4;';
  var thScroll=thBase+'position:sticky;top:0;z-index:3;text-align:right;';

  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'left:0;min-width:200px">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=MO_RANCH_COLORS[rn]||RANCH_COLORS[rn]||'#374151';
    h1+='<th colspan="'+nColsPerRanch+'" style="'+thScroll+'border-left:2px solid #8EA9C1;text-align:center;color:'+col+'">'+rn+'</th>';
  });
  if(showTotal) h1+='<th colspan="'+nColsTotal+'" style="'+thScroll+'border-left:3px solid #7B1C1C;text-align:center;background:#9DC3E6;color:#1e3a5f">TOTAL</th>';
  h1+='</tr>';

  // HEADER nivel 2: [por cada rancho: wk labels... | DIF] + TOTAL group (por año)
  var h2='<tr>';
  activeRanches.forEach(function(){
    weekKeys.forEach(function(key){
      var d=weekMap[key];
      var lbl=String(d._year).slice(2)+String(d._week).padStart(2,'0');
      var col=YEAR_COLORS[d._year]||'#888';
      h2+='<th style="'+thScroll+'border-left:1px solid var(--pt-hdr-border);font-size:9px;color:'+col+';min-width:60px">'+lbl+'</th>';
    });
    h2+='<th style="'+thScroll+'border-left:1px solid #aaa;font-size:9px;min-width:70px;background:#BDD7EE">DIF</th>';
  });
  // Sub-headers del grupo TOTAL (un encabezado por año)
  if(showTotal) {
    totYears.forEach(function(yr, i){
      var col=YEAR_COLORS[yr]||'#888';
      var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : 'border-left:1px solid var(--pt-hdr-border);';
      h2+='<th style="'+thScroll+lb+'font-size:9px;color:'+col+';min-width:90px;background:#EAF3FF">'+yr+'</th>';
    });
    if(totYears.length >= 2) {
      h2+='<th style="'+thScroll+'border-left:1px solid #aaa;font-size:9px;min-width:70px;background:#9DC3E6">DIF</th>';
    }
  }
  h2+='</tr>';

  // ── FILAS ─────────────────────────────────────────────
  var grandByRnWk={}, grandByRn={}, grandTotal=0;
  var grandHcByRnWk={};
  activeRanches.forEach(function(rn){
    grandByRnWk[rn]={}; grandByRn[rn]=0;
    grandHcByRnWk[rn]={};
    weekKeys.forEach(function(k){ grandByRnWk[rn][k]=0; grandHcByRnWk[rn][k]=0; });
  });

  var bodyHtml='';

  window.togglePtGroup = window.togglePtGroup || function(grpClass) {
    var rows = document.getElementsByClassName(grpClass);
    if (!rows.length) return;
    var isHidden = rows[0].style.display === 'none';
    for(var i=0; i<rows.length; i++){
      rows[i].style.display = isHidden ? '' : 'none';
    }
  };

  MO_GROUPS.forEach(function(grp, grpIdx){
    var grpByRnWk={}, grpByRn={}, grpTotal=0, grpHcByRn={};
    activeRanches.forEach(function(rn){
      grpByRnWk[rn]={}; grpByRn[rn]=0; grpHcByRn[rn]=0;
      weekKeys.forEach(function(k){ grpByRnWk[rn][k]=0; });
    });
    var scRows=[];

    grp.subcats.forEach(function(sc){
      var scByRnWk={}, scByRn={}, scTotal=0;
      var hcByRnWk={}, hcByRn={}, hcTotal=0;
      activeRanches.forEach(function(rn){
        scByRnWk[rn]={}; scByRn[rn]=0;
        hcByRnWk[rn]={}; hcByRn[rn]=0;
        weekKeys.forEach(function(k){
          var val=(weekMap[k]&&weekMap[k][sc+'__r__'+rn])?weekMap[k][sc+'__r__'+rn]:0;
          var hc_val=(weekMap[k]&&weekMap[k][sc+'__hc_r__'+rn])?weekMap[k][sc+'__hc_r__'+rn]:0;
          scByRnWk[rn][k]=val;
          scByRn[rn]+=val;
          scTotal+=val;
          hcByRnWk[rn][k]=hc_val;
          grandHcByRnWk[rn][k]=(grandHcByRnWk[rn][k]||0)+hc_val;
          grpByRnWk[rn][k]+=val;
          grpByRn[rn]+=val;
          grpTotal+=val;
          grandByRnWk[rn][k]+=val;
          grandByRn[rn]+=val;
          grandTotal+=val;
        });
        // SUB HC = ultima semana - primera semana del rango
        var firstKey=weekKeys[0], lastKey=weekKeys[weekKeys.length-1];
        hcByRn[rn]=(hcByRnWk[rn][lastKey]||0)-(hcByRnWk[rn][firstKey]||0);
      });
      // hcTotal = suma de diferencias por rancho
      hcTotal=activeRanches.reduce(function(s,rn){return s+(hcByRn[rn]||0);},0);
      activeRanches.forEach(function(rn){ grpHcByRn[rn]+=(hcByRn[rn]||0); });
      if(scTotal>0) scRows.push({label:shortLabel(sc),byRnWk:scByRnWk,byRn:scByRn,total:scTotal, hcByRnWk:hcByRnWk, hcByRn:hcByRn, hcTotal:hcTotal});
    });
    if(grpTotal===0) return;

    // Fila grupo
    var gTdPin='padding:4px 8px;background:var(--pt-grp-bg);color:#fff;font-weight:700;font-size:11px;position:sticky;z-index:2;border-bottom:1px solid #ddd;border-right:1px solid rgba(255,255,255,0.2);white-space:nowrap;';
    bodyHtml+='<tr style="cursor:pointer;" onclick="togglePtGroup(`mo_grp_'+grpIdx+'`)" title="Clic para expandir/contraer">';
    bodyHtml+='<td style="'+gTdPin+'left:0">'+grp.label+'</td>';
    activeRanches.forEach(function(rn){
      weekKeys.forEach(function(key){
        bodyHtml+=cellGrp(grpByRnWk[rn][key]);
      });
      var firstKey=weekKeys[0], lastKey=weekKeys[weekKeys.length-1];
      var costDif=(grpByRnWk[rn][lastKey]||0)-(grpByRnWk[rn][firstKey]||0);
      var grpDifStyle='padding:3px 6px;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;background:var(--pt-grp-bg);color:#fff;font-weight:700';
      bodyHtml+='<td style="'+grpDifStyle+'">'+(costDif?(costDif>0?'+':'')+fmt(costDif):'—')+'</td>';
    });
    // Celdas TOTAL por año (suma del rango, todos los ranchos)
    var grpTotStyle='padding:3px 6px;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;background:#9DC3E6;color:#1e3a5f;font-weight:700;border-left:1px solid #8EA9C1;';
    if(showTotal) {
      var grpByYr={};
      totYears.forEach(function(yr){ grpByYr[yr]=0; });
      weekKeys.forEach(function(key){
        var yr=weekMap[key]._year;
        var v=activeRanches.reduce(function(s,rn){return s+(grpByRnWk[rn][key]||0);},0);
        grpByYr[yr]=(grpByYr[yr]||0)+v;
      });
      totYears.forEach(function(yr, i){
        var v=grpByYr[yr]||0;
        var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
        bodyHtml+='<td style="'+grpTotStyle+lb+'">'+( v?fmt(v):'—')+'</td>';
      });
      if(totYears.length >= 2) {
        var grpTotDif=(grpByYr[totYears[totYears.length-1]]||0)-(grpByYr[totYears[0]]||0);
        bodyHtml+='<td style="'+grpTotStyle+'border-left:1px solid #aaa;background:#7EB3D4">'+(grpTotDif?(grpTotDif>0?'+':'')+fmt(grpTotDif):'—')+'</td>';
      }
    }
    bodyHtml+='</tr>';

    // Filas subcat
    scRows.forEach(function(sc){
      var tdPin='padding:3px 8px;position:sticky;z-index:1;background:#fff;border-bottom:1px solid #eee;border-right:1px solid #eee;white-space:nowrap;';
      bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="Número de personas (Headcount)">';
      bodyHtml+='<td style="'+tdPin+'left:0;padding-left:20px;color:#334155;font-size:11px"><span style="color:#888;font-size:9px;margin-right:4px;">👤</span>'+sc.label+'</td>';
      activeRanches.forEach(function(rn){
        weekKeys.forEach(function(key){
          var v=sc.hcByRnWk[rn][key];
          if(!v||v===0){bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ddd">—</td>';}
          else{bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#475569;font-weight:600">'+fmtHc(v)+'</td>';}
        });
        var cellStyle = 'padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#1e3a5f;font-weight:700;';
        bodyHtml+='<td style="'+cellStyle+'">'+(sc.hcByRn[rn]?fmtHcDiff(sc.hcByRn[rn]):'—')+'</td>';
      });
      // Celdas TOTAL HC por año (última semana del año = snapshot de headcount)
      var scTotHcStyle='padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;background:#EAF3FF;color:#1e3a5f;font-weight:700;';
      if(showTotal) {
        var scHcByYr={};
        totYears.forEach(function(yr){
          var lastKeyOfYr=null;
          weekKeys.forEach(function(k){ if(weekMap[k]._year===yr) lastKeyOfYr=k; });
          scHcByYr[yr]=lastKeyOfYr ? activeRanches.reduce(function(s,rn){return s+(sc.hcByRnWk[rn][lastKeyOfYr]||0);},0) : 0;
        });
        totYears.forEach(function(yr, i){
          var v=scHcByYr[yr]||0;
          var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
          bodyHtml+='<td style="'+scTotHcStyle+lb+'">'+( v?fmtHc(v):'—')+'</td>';
        });
        if(totYears.length >= 2) {
          var scHcDifYr=(scHcByYr[totYears[totYears.length-1]]||0)-(scHcByYr[totYears[0]]||0);
          bodyHtml+='<td style="'+scTotHcStyle+'border-left:1px solid #aaa;background:#9DC3E6">'+(scHcDifYr?fmtHcDiff(scHcDifYr):'—')+'</td>';
        }
      }
      bodyHtml+='</tr>';
    });

    // ── Fila $/TALLO, $/METRO o $/CHAROLA — Costo Unitario ─────────────────
    // Solo para grupos que tienen una métrica real como denominador
    var _utGroupMap = {
      'CORTE':          {key:'tallos_cos', label:'$/UNIT',  title:'Costo Unitario por Tallo Cosechado: '},
      'TRASPLANTE':     {key:'charolas',   label:'$/UNIT',  title:'Costo Unitario por Charola Sembrada: '},
      'MANEJO PLANTA':  {key:'metros',     label:'$/UNIT',  title:'Costo Unitario por Metro de Siembra: '},
      'MIPE / MIRFE':   {key:'metros',     label:'$/UNIT',  title:'Costo Unitario por Metro de Siembra: '}
    };
    var _utConf = _utGroupMap[grp.label];
    if (_utConf && DATA.siembra_data) {
      var _utDenomKey = _utConf.key;
      var _utLabel = _utConf.label;
      var _utTitle = _utConf.title;
      var utPin='padding:3px 8px;position:sticky;z-index:1;background:#fffbeb;border-bottom:1px solid #fde68a;border-right:1px solid #fde68a;white-space:nowrap;';
      var utStyle='padding:3px 6px;border-bottom:1px solid #fde68a;border-right:1px solid #fde68a;text-align:right;color:#92400e;font-weight:600;background:#fffbeb;';
      bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="'+_utTitle+grp.label+'">';
      bodyHtml+='<td style="'+utPin+'left:0;padding-left:20px;color:#92400e;font-size:11px"><span style="color:#f59e0b;font-size:9px;margin-right:4px;">💲</span>'+_utLabel+'</td>';
      var _utCostByWk={}, _utDenomByWk={};
      weekKeys.forEach(function(k){ _utCostByWk[k]=0; _utDenomByWk[k]=0; });
      activeRanches.forEach(function(rn){
        var _firstCpt=0, _lastCpt=0;
        weekKeys.forEach(function(key, wi){
          var cost=grpByRnWk[rn][key]||0;
          var pw=key.split('-');
          var wkk=((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
          var sData=DATA.siembra_data[wkk]||DATA.siembra_data[String(wkk)]||{};
          var targetRn=({'Ramona':'Campo-RM','Poscosecha':'PosCo-RM','Propagacion':'Prop-RM'})[rn]||rn;
          var sRow=sData[targetRn]||sData[rn]||{};
          var denom=sRow[_utDenomKey]||0;
          var cpt=(denom>0)?cost/denom:0;
          _utCostByWk[key]+=cost; _utDenomByWk[key]+=denom;
          if(wi===0) _firstCpt=cpt;
          _lastCpt=cpt;
          if(!cpt){bodyHtml+='<td style="'+utStyle+'color:#fcd34d">—</td>';}
          else{bodyHtml+='<td style="'+utStyle+'">$'+cpt.toFixed(2)+'</td>';}
        });
        var _utDif=_lastCpt-_firstCpt;
        var _utDifStyle=utStyle+'border-left:1px solid #fbbf24;';
        if(!_utDif||Math.abs(_utDif)<0.005){bodyHtml+='<td style="'+_utDifStyle+'color:#fcd34d">—</td>';}
        else{var _utCl=_utDif>0?'#dc2626':'#16a34a'; bodyHtml+='<td style="'+_utDifStyle+'color:'+_utCl+'">'+(_utDif>0?'+':'')+_utDif.toFixed(2)+'</td>';}
      });
      if(showTotal){
        var utTotStyle='padding:3px 6px;border-bottom:1px solid #fde68a;border-right:1px solid #fde68a;text-align:right;background:#fef3c7;color:#92400e;font-weight:700;';
        var _utCostYr={}, _utDenomYr={};
        totYears.forEach(function(yr){ _utCostYr[yr]=0; _utDenomYr[yr]=0; });
        weekKeys.forEach(function(key){ var yr=weekMap[key]._year; _utCostYr[yr]+=(_utCostByWk[key]||0); _utDenomYr[yr]+=(_utDenomByWk[key]||0); });
        totYears.forEach(function(yr, i){
          var cpt=(_utDenomYr[yr]>0)?_utCostYr[yr]/_utDenomYr[yr]:0;
          var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
          bodyHtml+='<td style="'+utTotStyle+lb+'">'+( cpt?'$'+cpt.toFixed(2):'—')+'</td>';
        });
        if(totYears.length >= 2) {
          var cpt0=(_utDenomYr[totYears[0]]>0)?_utCostYr[totYears[0]]/_utDenomYr[totYears[0]]:0;
          var cptN=(_utDenomYr[totYears[totYears.length-1]]>0)?_utCostYr[totYears[totYears.length-1]]/_utDenomYr[totYears[totYears.length-1]]:0;
          var _totUtDif=cptN-cpt0;
          var _totUtDifStyle=utTotStyle+'border-left:1px solid #fbbf24;background:#fde68a;';
          if(!_totUtDif||Math.abs(_totUtDif)<0.005){bodyHtml+='<td style="'+_totUtDifStyle+'color:#fcd34d">—</td>';}
          else{var _totUtCl=_totUtDif>0?'#dc2626':'#16a34a'; bodyHtml+='<td style="'+_totUtDifStyle+'color:'+_totUtCl+'">'+(_totUtDif>0?'+':'')+_totUtDif.toFixed(2)+'</td>';}
        }
      }
      bodyHtml+='</tr>';
    }

    if (grp.label === 'CORTE' && DATA.siembra_data) {
      var tcPin='padding:3px 8px;position:sticky;z-index:1;background:#f0fdf4;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;white-space:nowrap;';
      var tcStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;color:#15803d;font-weight:600;background:#f0fdf4;';
      bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="TALLOS COSECHADOS">';
      bodyHtml+='<td style="'+tcPin+'left:0;padding-left:20px;color:#166534;font-size:11px"><span style="color:#22c55e;font-size:9px;margin-right:4px;">🌱</span>TALLOS COSECHADOS</td>';
      var wkTcTotal = {};
      var wKeyOrdered = [];
      activeRanches.forEach(function(rn){
        weekKeys.forEach(function(key){
          var pw=key.split('-');
          var wkk = ((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
          var sData = DATA.siembra_data[wkk] || DATA.siembra_data[String(wkk)] || {};
          var revMap = {'Ramona':'Campo-RM', 'Poscosecha':'PosCo-RM', 'Propagacion':'Prop-RM'};
          var targetRn = revMap[rn] || rn;
          var sRow = sData[targetRn] || sData[rn] || {};
          var v = sRow['tallos_cos'] || 0;
          wkTcTotal[key] = (wkTcTotal[key]||0) + v;
          if(wKeyOrdered.indexOf(key)===-1) wKeyOrdered.push(key);
          if(!v||v===0) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
          else { bodyHtml+='<td style="'+tcStyle+'">'+Number(v).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        });
        var wkTcDiff = 0;
        if (weekKeys.length > 0) {
          var fw = weekKeys[0], lw = weekKeys[weekKeys.length-1];
          var getV = function(key) {
            var pw=key.split('-');
            var wkk = ((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
            var sData = DATA.siembra_data[wkk] || DATA.siembra_data[String(wkk)] || {};
            var revMap = {'Ramona':'Campo-RM', 'Poscosecha':'PosCo-RM', 'Propagacion':'Prop-RM'};
            var targetRn = revMap[rn] || rn;
            var sRow = sData[targetRn] || sData[rn] || {};
            return sRow['tallos_cos'] || 0;
          };
          wkTcDiff = getV(lw) - getV(fw);
        }
        if(!wkTcDiff||wkTcDiff===0) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
        else { bodyHtml+='<td style="'+tcStyle+'">'+(wkTcDiff>0?'+':'')+Number(wkTcDiff).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
      });
      if(showTotal) {
        var tcTotHcStyle = 'padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#dcfce7;color:#166534;font-weight:700;';
        var gtStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#bbf7d0;color:#166534;font-weight:700;border-left:1px solid #86efac;';
        var tcTotByYr={};
        totYears.forEach(function(yr){ tcTotByYr[yr]=0; });
        wKeyOrdered.forEach(function(key){ var yr=weekMap[key]._year; tcTotByYr[yr]=(tcTotByYr[yr]||0)+(wkTcTotal[key]||0); });
        totYears.forEach(function(yr, i){
          var v=tcTotByYr[yr]||0;
          var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
          if(!v) { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">—</td>'; }
          else { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">'+Number(v).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        });
        if(totYears.length >= 2) {
          var gtDif=(tcTotByYr[totYears[totYears.length-1]]||0)-(tcTotByYr[totYears[0]]||0);
          if(!gtDif) { bodyHtml+='<td style="'+gtStyle+'">—</td>'; }
          else { bodyHtml+='<td style="'+gtStyle+'">'+(gtDif>0?'+':'')+Number(gtDif).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        }
      }
      bodyHtml+='</tr>';
    }

    if (grp.label === 'TRASPLANTE' && DATA.siembra_data) {
      var tcPin='padding:3px 8px;position:sticky;z-index:1;background:#f0fdf4;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;white-space:nowrap;';
      var tcStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;color:#15803d;font-weight:600;background:#f0fdf4;';
      bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="NUMERO DE CHAROLAS SEMBRADAS">';
      bodyHtml+='<td style="'+tcPin+'left:0;padding-left:20px;color:#166534;font-size:11px"><span style="color:#22c55e;font-size:9px;margin-right:4px;">🌱</span>NUMERO DE CHAROLAS SEMBRADAS</td>';
      var wkTcTotal = {};
      var wKeyOrdered = [];
      activeRanches.forEach(function(rn){
        weekKeys.forEach(function(key){
          var pw=key.split('-');
          var wkk = ((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
          var sData = DATA.siembra_data[wkk] || DATA.siembra_data[String(wkk)] || {};
          var revMap = {'Ramona':'Campo-RM', 'Poscosecha':'PosCo-RM', 'Propagacion':'Prop-RM'};
          var targetRn = revMap[rn] || rn;
          var sRow = sData[targetRn] || sData[rn] || {};
          var v = sRow['charolas'] || 0;
          wkTcTotal[key] = (wkTcTotal[key]||0) + v;
          if(wKeyOrdered.indexOf(key)===-1) wKeyOrdered.push(key);
          if(!v||v===0) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
          else { bodyHtml+='<td style="'+tcStyle+'">'+Number(v).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        });
        var wkTcDiff = 0;
        if (weekKeys.length > 0) {
          var fw = weekKeys[0], lw = weekKeys[weekKeys.length-1];
          var getV = function(key) {
            var pw=key.split('-');
            var wkk = ((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
            var sData = DATA.siembra_data[wkk] || DATA.siembra_data[String(wkk)] || {};
            var revMap = {'Ramona':'Campo-RM', 'Poscosecha':'PosCo-RM', 'Propagacion':'Prop-RM'};
            var targetRn = revMap[rn] || rn;
            var sRow = sData[targetRn] || sData[rn] || {};
            return sRow['charolas'] || 0;
          };
          wkTcDiff = getV(lw) - getV(fw);
        }
        if(!wkTcDiff||wkTcDiff===0) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
        else { bodyHtml+='<td style="'+tcStyle+'">'+(wkTcDiff>0?'+':'')+Number(wkTcDiff).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
      });
      if(showTotal) {
        var tcTotHcStyle = 'padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#dcfce7;color:#166534;font-weight:700;';
        var gtStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#bbf7d0;color:#166534;font-weight:700;border-left:1px solid #86efac;';
        var cTotByYr={};
        totYears.forEach(function(yr){ cTotByYr[yr]=0; });
        wKeyOrdered.forEach(function(key){ var yr=weekMap[key]._year; cTotByYr[yr]=(cTotByYr[yr]||0)+(wkTcTotal[key]||0); });
        totYears.forEach(function(yr, i){
          var v=cTotByYr[yr]||0;
          var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
          if(!v) { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">—</td>'; }
          else { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">'+Number(v).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        });
        if(totYears.length >= 2) {
          var gtDif=(cTotByYr[totYears[totYears.length-1]]||0)-(cTotByYr[totYears[0]]||0);
          if(!gtDif) { bodyHtml+='<td style="'+gtStyle+'">—</td>'; }
          else { bodyHtml+='<td style="'+gtStyle+'">'+(gtDif>0?'+':'')+Number(gtDif).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>'; }
        }
      }
      bodyHtml+='</tr>';
    }

    if ((grp.label === 'MANEJO PLANTA' || grp.label === 'MIPE / MIRFE') && DATA.siembra_data) {
      var tcPin='padding:3px 8px;position:sticky;z-index:1;background:#f0fdf4;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;white-space:nowrap;';
      var tcStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;color:#15803d;font-weight:600;background:#f0fdf4;';
      
      var metrics = [
        { key: 'metros', title: 'METROS DE SIEMBRA', maxDecimals: 0, minDecimals: 0 },
        { key: 'hectareas', title: 'HECTAREAS EN SIEMBRA', maxDecimals: 4, minDecimals: 2 }
      ];

      metrics.forEach(function(m) {
        bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="'+m.title+'">';
        bodyHtml+='<td style="'+tcPin+'left:0;padding-left:20px;color:#166534;font-size:11px"><span style="color:#22c55e;font-size:9px;margin-right:4px;">🌱</span>'+m.title+'</td>';
        
        var wkTcTotal = {};
        var fw = weekKeys[0], lw = weekKeys[weekKeys.length-1];
        
        var getV = function(rn, k) {
            var pw = k.split('-');
            var wkk = ((parseInt(pw[0])%100)*100)+parseInt(pw[1]);
            var sData = DATA.siembra_data[wkk] || DATA.siembra_data[String(wkk)] || {};
            var targetRn = ({'Ramona':'Campo-RM', 'Poscosecha':'PosCo-RM', 'Propagacion':'Prop-RM'})[rn] || rn;
            return (sData[targetRn] || sData[rn] || {})[m.key] || 0;
        };

        activeRanches.forEach(function(rn){
          weekKeys.forEach(function(key){
            var v = getV(rn, key);
            wkTcTotal[key] = (wkTcTotal[key]||0) + v;
            if(!v) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
            else { bodyHtml+='<td style="'+tcStyle+'">'+Number(v).toLocaleString('es-MX',{minimumFractionDigits:m.minDecimals, maximumFractionDigits:m.maxDecimals})+'</td>'; }
          });
          var wkTcDiff = 0;
          if (weekKeys.length > 0) wkTcDiff = getV(rn, lw) - getV(rn, fw);
          if(!wkTcDiff) { bodyHtml+='<td style="'+tcStyle+'color:#86efac">—</td>'; }
          else { bodyHtml+='<td style="'+tcStyle+'">'+(wkTcDiff>0?'+':'')+Number(wkTcDiff).toLocaleString('es-MX',{minimumFractionDigits:m.minDecimals, maximumFractionDigits:m.maxDecimals})+'</td>'; }
        });

        if(showTotal) {
          var tcTotHcStyle = 'padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#dcfce7;color:#166534;font-weight:700;';
          var gtStyle='padding:3px 6px;border-bottom:1px solid #dcfce7;border-right:1px solid #dcfce7;text-align:right;background:#bbf7d0;color:#166534;font-weight:700;border-left:1px solid #86efac;';
          var mTotByYr={};
          totYears.forEach(function(yr){ mTotByYr[yr]=0; });
          weekKeys.forEach(function(key){ var yr=weekMap[key]._year; mTotByYr[yr]=(mTotByYr[yr]||0)+(wkTcTotal[key]||0); });
          totYears.forEach(function(yr, i){
            var v=mTotByYr[yr]||0;
            var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
            if(!v) { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">—</td>'; }
            else { bodyHtml+='<td style="'+tcTotHcStyle+lb+'">'+Number(v).toLocaleString('es-MX',{minimumFractionDigits:m.minDecimals, maximumFractionDigits:m.maxDecimals})+'</td>'; }
          });
          if(totYears.length >= 2) {
            var gtDif=(mTotByYr[totYears[totYears.length-1]]||0)-(mTotByYr[totYears[0]]||0);
            if(!gtDif) { bodyHtml+='<td style="'+gtStyle+'">—</td>'; }
            else { bodyHtml+='<td style="'+gtStyle+'">'+(gtDif>0?'+':'')+Number(gtDif).toLocaleString('es-MX',{minimumFractionDigits:m.minDecimals, maximumFractionDigits:m.maxDecimals})+'</td>'; }
          }
        }
        bodyHtml+='</tr>';
      });
    }

  });

  // Fila total general
  // ── Totales por semana (filas al fondo) ───────────────
  var _fk=weekKeys[0], _lk=weekKeys[weekKeys.length-1];
  var grandCostWk={}, grandHcWk={};
  weekKeys.forEach(function(k){
    grandCostWk[k]=activeRanches.reduce(function(s,rn){return s+(grandByRnWk[rn][k]||0);},0);
    grandHcWk[k]=activeRanches.reduce(function(s,rn){return s+(grandHcByRnWk[rn][k]||0);},0);
  });

  var totStyle='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;color:#1e3a5f;';
  var totPin='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;left:0;z-index:2;white-space:nowrap;color:#1e3a5f;';
  var totHcStyle='padding:4px 8px;background:#ddeedd;font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;color:#1e3a5f;';
  var totHcPin='padding:4px 8px;background:#ddeedd;font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;left:0;z-index:2;white-space:nowrap;color:#1e3a5f;';
  var dash='<td style="'+totStyle+'color:#bbb">—</td>';
  var dashHc='<td style="'+totHcStyle+'color:#bbb">—</td>';

  // Fila TOTAL $
  bodyHtml+='<tr>';
  bodyHtml+='<td style="'+totPin+'">TOTAL $</td>';
  activeRanches.forEach(function(rn){
    weekKeys.forEach(function(key){
      var v=grandByRnWk[rn][key]||0;
      bodyHtml+='<td style="'+totStyle+'">'+(v?fmt(v):'—')+'</td>';
    });
    var rnDif=(grandByRnWk[rn][weekKeys[weekKeys.length-1]]||0)-(grandByRnWk[rn][weekKeys[0]]||0);
    bodyHtml+='<td style="'+totStyle+'border-left:1px solid #aaa">'+(rnDif?(rnDif>0?'+':'')+fmt(rnDif):'—')+'</td>';
  });
  var totTotStyle='padding:4px 8px;background:#9DC3E6;font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;color:#1e3a5f;';
  if(showTotal) {
    var grandCostYr={};
    totYears.forEach(function(yr){ grandCostYr[yr]=0; });
    weekKeys.forEach(function(k){ var yr=weekMap[k]._year; grandCostYr[yr]=(grandCostYr[yr]||0)+(grandCostWk[k]||0); });
    totYears.forEach(function(yr, i){
      var v=grandCostYr[yr]||0;
      var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
      bodyHtml+='<td style="'+totTotStyle+lb+'">'+( v?fmt(v):'—')+'</td>';
    });
    if(totYears.length >= 2) {
      var gtDifYr=(grandCostYr[totYears[totYears.length-1]]||0)-(grandCostYr[totYears[0]]||0);
      bodyHtml+='<td style="'+totTotStyle+'border-left:1px solid #aaa;background:#7EB3D4">'+(gtDifYr?(gtDifYr>0?'+':'')+fmt(gtDifYr):'—')+'</td>';
    }
  }
  bodyHtml+='</tr>';

  // Fila TOTAL 👤
  bodyHtml+='<tr>';
  bodyHtml+='<td style="'+totHcPin+'">TOTAL 👤</td>';
  activeRanches.forEach(function(rn){
    weekKeys.forEach(function(key){
      var v=grandHcByRnWk[rn][key]||0;
      bodyHtml+='<td style="'+totHcStyle+'">'+(v?fmtHc(v):'—')+'</td>';
    });
    bodyHtml+=dashHc;
  });
  var totHcTotStyle='padding:4px 8px;background:#9DC3E6;font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;color:#1e3a5f;';
  if(showTotal) {
    var grandHcYr={};
    totYears.forEach(function(yr){
      var lastKeyOfYr=null;
      weekKeys.forEach(function(k){ if(weekMap[k]._year===yr) lastKeyOfYr=k; });
      grandHcYr[yr]=lastKeyOfYr ? grandHcWk[lastKeyOfYr]||0 : 0;
    });
    totYears.forEach(function(yr, i){
      var v=grandHcYr[yr]||0;
      var lb = i===0 ? 'border-left:3px solid #7B1C1C;' : '';
      bodyHtml+='<td style="'+totHcTotStyle+lb+'">'+( v?fmtHc(v):'—')+'</td>';
    });
    if(totYears.length >= 2) {
      var gtHcDifYr=(grandHcYr[totYears[totYears.length-1]]||0)-(grandHcYr[totYears[0]]||0);
      bodyHtml+='<td style="'+totHcTotStyle+'border-left:1px solid #aaa;background:#7EB3D4">'+(gtHcDifYr?fmtHcDiff(gtHcDifYr):'—')+'</td>';
    }
  }
  bodyHtml+='</tr>';

  // ── Inyectar en el DOM ────────────────────────────────
  var html='<div class="pt-table-wrap" id="tableWrap" style="overflow-x:auto;overflow-y:visible;"><table class="pt-table" style="border-collapse:collapse;width:100%"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
  var gw=document.getElementById('gridWrap');
  if(gw){ gw.style.display=''; gw.innerHTML=html; }
  document.getElementById('comparativoWrap').className='';
  document.getElementById('stTotal').textContent=fmt(grandTotal)+' '+sym;
  setTimeout(resizeTable,80);
}

function onMainCellClick(evt) {
  if (!evt||!evt.data||!evt.colDef) return;
  var data=evt.data, clickedField=evt.colDef.field||'';
  var clickedRanch=fieldToRanch(clickedField);
  if (state.view==='semana') { showProdPanel(data,{ranch:clickedRanch||null}); return; }
  if (state.view==='rancho') { if (clickedField==='rancho'||clickedRanch) showProdPanel(data,{ranch:data.rancho||null}); }
}

// =======================================================
// PRODUCTOS SUBPANEL
// =======================================================
var _prodViews = [];
var _prodViewsData = []; // guarda {rowData, opts} para re-renderizar al cambiar moneda

function abrirProdGlobal() {
  if (state.cat === 'COSTO SERVICIOS' || state.cat === 'COSTO MANO DE OBRA') return;
  
  _prodViews = [];
  _prodViewsData = [];

  var activeYrs = DATA.years.filter(function(y) { return state.activeYears[y]; }).sort(function(a, b) { return b - a; }); 

  activeYrs.forEach(function(yr) {
    showProdPanel({_cat:state.cat, _year:yr, _fromWeek:state.fromWeek, _toWeek:state.toWeek}, {ranch: 'Multi'});
  });
}

window.filterProdRows = function(el) {
  if (!el) return;
  var term = el.value ? el.value.toLowerCase() : '';
  var panel = el.closest('.prod-panel-wrapper');
  if (!panel) return;
  var rows = panel.querySelectorAll('.pt-prod-row');
  var nuevoTotal = 0;
  for(var i=0; i<rows.length; i++){
    var p = rows[i].getAttribute('data-prod') || '';
    if (term === '' || p.indexOf(term) > -1) {
      if(rows[i].style.display === 'none'){
        rows[i].style.display = '';
        nuevoTotal += parseFloat(rows[i].getAttribute('data-costo')||0);
      } else {
        nuevoTotal += parseFloat(rows[i].getAttribute('data-costo')||0);
      }
    } else {
      if(rows[i].style.display === ''){
        rows[i].style.display = 'none';
      }
    }
  }
  // recalcular el TOTAL (MXN u USD)
  var metaSpan = panel.querySelector('.gasto-total-val');
  if (metaSpan) {
     var _isUSD = window._prodIsUSD;
     var _TC = window._prodTC || 19;
     var convTotal = _isUSD ? nuevoTotal/_TC : nuevoTotal;
     if(convTotal === 0){
        metaSpan.innerHTML = _isUSD ? '$0.00' : '$0';
     } else {
        var neg = convTotal < 0, s = Math.abs(convTotal);
        var txt = _isUSD 
           ? (neg?'-$':'$') + s.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})
           : fmt(nuevoTotal);
        metaSpan.innerHTML = txt;
     }
  }
};

function showProdPanel(rowData, opts) {
  opts=opts||{};
  var cat=rowData._cat, yr=rowData._year, wn=rowData._week;
  var fromW=rowData._fromWeek||wn, toW=rowData._toWeek||wn;
  var ranchFilter=opts.ranch||null;
  if (ranchFilter === 'Multi') {
    if (state.activeRanches.indexOf('Todos') >= 0) ranchFilter = null;
    else if (state.activeRanches.length === 1) ranchFilter = state.activeRanches[0];
    else ranchFilter = null;
  } else if (!ranchFilter && state.activeRanches.indexOf('Todos')<0) {
    ranchFilter = state.activeRanches[0];
  }
  if (!cat||!yr) return;
  var hideProducts = false;

  var isRanchAllowed = function(rn) {
    var rawRn = rn.replace('Prop-RM','Propagacion').replace('PosCo-RM','Poscosecha').replace('Campo-RM','Ramona');
    if (ranchFilter) return rn === ranchFilter || rawRn === ranchFilter;
    if (state.activeRanches.indexOf('Todos') > -1) return true;
    return state.activeRanches.indexOf(rn) > -1 || state.activeRanches.indexOf(rawRn) > -1;
  };

  var isMant=cat==='MANTENIMIENTO', isMatEmp=cat==='MATERIAL DE EMPAQUE', isMatVeg=cat==='MATERIAL VEGETAL';
  var isMirfe=cat===CAT_MIRFE, isMipe=cat===CAT_MIPE;
  var src=isMant?'mp':(isMatEmp?'me':(isMatVeg?'mv':'pr'));
  var tipoFilter=null;
  if (src==='pr'){ if(isMirfe)tipoFilter='MIRFE'; else if(isMipe)tipoFilter='MIPE'; }
  var dsMap={pr:DATA.productos,mp:DATA.productos_mp,me:DATA.productos_me,mv:DATA.productos_mv};
  var ds=dsMap[src]||{};

  var wkStart=parseInt(fromW||wn||0), wkEnd=parseInt(toW||wn||0);
  if (!wkStart||!wkEnd) return;
  if (wkStart>wkEnd){var tmp=wkStart;wkStart=wkEnd;wkEnd=tmp;}

  var rows=[];
  if (!hideProducts) {
    for (var wk=wkStart;wk<=wkEnd;wk++){
      var wkCodeShort=((yr%100)*100)+wk, wkCodeLong=(yr*100)+wk;
      var weekD=ds[wkCodeShort]||ds[String(wkCodeShort)]||ds[wkCodeLong]||ds[String(wkCodeLong)];
      if (!weekD) continue;
      Object.keys(weekD).forEach(function(ranch){
        if (!isRanchAllowed(ranch)) return;
        var byTipo=weekD[ranch];
        Object.keys(byTipo).forEach(function(tipo){
          if (tipoFilter&&tipo!==tipoFilter) return;
          (byTipo[tipo]||[]).forEach(function(item){
            rows.push({week_code:wkCodeShort,rancho:ranch,tipo:tipo,producto:item[0]||'',unidades:item[1]||'',gasto:parseFloat(item[2])||0,ubicacion:item[3]||''});
          });
        });
      });
    }
  }

  var rangeText=wkStart===wkEnd?(wFmt(wkStart)+' · '+yr):(wFmt(wkStart)+'→'+wFmt(wkEnd)+' · '+yr);
  // Formato: YYSS - CATEGORÍA - RANCHO  (ej: 2615 - MATERIAL VEGETAL - PROP-RM)
  var _yyss = String(yr).slice(2) + String(wkStart).padStart(2,'0');
  var _wkLabel = wkStart===wkEnd ? _yyss : (String(yr).slice(2)+String(wkStart).padStart(2,'0')+'→'+String(yr).slice(2)+String(wkEnd).padStart(2,'0'));
  var panelTitle = _wkLabel + ' - ' + cat + (ranchFilter ? ' - ' + ranchFilter.toUpperCase() : '');
  
  var panelHtml = '';

  // ── Resumen de siembra (se integra dentro de la tabla) ───────────
  var siembraRowsHtml = '';
  var _allMetas = [
    {k:'inv_inicial',  lbl:'INV. INICIAL'},
    {k:'tallos_cos',   lbl:'TALLOS COSECHADOS'},
    {k:'tallos_des',   lbl:'TALLOS DESECHADOS'},
    {k:'tallos_des_sf',lbl:'TALLOS DESECHADOS SF'},
    {k:'tallos_comp',  lbl:'TALLOS COMPRADOS'},
    {k:'tallos_bouq',  lbl:'TALLOS BOUQUETS/PROC.'},
    {k:'tallos_desp',  lbl:'TALLOS DESPACHADOS'},
    {k:'libras_alb',   lbl:'LIBRAS ALBAHACA'},
    {k:'tallos_mues',  lbl:'TALLOS MUESTRA'},
    {k:'inv_final',    lbl:'INV. FINAL'},
    {k:'tallos_proc',  lbl:'TALLOS PROC. TOTALES'},
    {k:'charolas_288', lbl:'CHAROLAS *288'},
    {k:'charolas',     lbl:'NUMERO DE CHAROLAS SEMBRADAS'},
    {k:'esquejes',     lbl:'N\u00ba ESQUEJES'},
    {k:'metros',       lbl:'METROS SIEMBRA'},
    {k:'hectareas',    lbl:'HECT\u00c1REAS'},
  ];
  if (DATA.siembra_data) {
    // ── Agregar siembra_data de TODAS las semanas en el rango ──────────
    var _yyBase = (yr % 100) * 100;
    var _aggSiembra = {};  // {metricKey: valorSumado}
    var _wkKeysInRange = [];
    for (var _wk = wkStart; _wk <= wkEnd; _wk++) {
      var _wkKey = _yyBase + _wk;
      _wkKeysInRange.push(_wkKey);
      var _wkSrc = DATA.siembra_data[_wkKey] || DATA.siembra_data[String(_wkKey)] || null;
      if (!_wkSrc) continue;
      var _sRow = ranchFilter ? (_wkSrc[ranchFilter] || _wkSrc['TOTAL'] || {}) : (_wkSrc['TOTAL'] || {});
      _allMetas.forEach(function(m) {
        var v = _sRow[m.k];
        if (v !== undefined && v !== null && v !== '' && v !== 0) {
          _aggSiembra[m.k] = (_aggSiembra[m.k] || 0) + Number(v);
        }
      });
    }
    var _isRange = (wkStart !== wkEnd);
    var _activos = _allMetas.filter(function(m){ var v=_aggSiembra[m.k]; return v!==undefined&&v!==null&&v!==0; });
    if (_activos.length > 0) {
      _activos.forEach(function(m, i){
        var bg = (i % 2 === 0) ? '#ffffff' : '#f8fafc';
        var isMetros = (m.k === 'metros');
        var isCharolas = (m.k === 'charolas');
        var _WEEKLY_KEYS = ['inv_inicial','tallos_cos','tallos_des','tallos_comp','tallos_desp','inv_final','tallos_proc'];
        var isWeeklyDetail = _WEEKLY_KEYS.indexOf(m.k) >= 0;
        var isExpandible = isMetros || isCharolas || isWeeklyDetail;
        var lblStyle = isExpandible ? 'cursor:pointer; color:#1e293b; transition: color 0.15s;' : 'color:#475569;';
        var expandIcon = isExpandible ? '<span style="color:#94a3b8; font-size:9px; margin-left:6px; user-select:none;">▼</span>' : '';
        var toggleClick = isExpandible ? 'onclick="var n=this.closest(\\\'tr\\\').nextElementSibling; if(n && n.className.indexOf(\\\'detail\\\')>0) { var isHidden = n.style.display===\\\'none\\\'; n.style.display=(isHidden?\\\'table-row\\\':\\\'none\\\'); this.querySelector(\\\'span\\\').innerHTML=isHidden?\\\'▲\\\':\\\'▼\\\'; }"' : '';
        
        siembraRowsHtml +=
          '<tr style="background:'+bg+';border-bottom:1px solid #e2e8f0; transition: background 0.2s;">' +
            '<td style="padding:6px 10px;white-space:nowrap;width:1%;"><span style="background:#e2e8f0;color:#475569;font-size:9px;font-weight:700;padding:2px 8px;border-radius:4px;letter-spacing:0.5px;">SIEMBRA</span></td>' +
            '<td colspan="2" style="padding:6px 10px;font-weight:600;"><span style="'+lblStyle+'" '+toggleClick+'>'+m.lbl+expandIcon+'</span></td>' +
            '<td style="padding:6px 10px;text-align:right;font-weight:700;color:#0f172a;font-size:11px;">'+Number(_aggSiembra[m.k]).toLocaleString('es-MX',{maximumFractionDigits:2})+'</td>' +
          '</tr>';
          
        if (isMetros && DATA.metros_acumulados) {
          var mRows = DATA.metros_acumulados.filter(function(r){ 
            var sf = parseInt(r.semana_fin);
            return _wkKeysInRange.indexOf(sf) >= 0 && isRanchAllowed(r.rancho); 
          });
          // Si es rango, agrupar por rancho+flor y sumar metros/plantas
          if (_isRange && mRows.length > 0) {
            var _mAgr = {};
            mRows.forEach(function(mr) {
              var k = (mr.rancho||'') + '||' + (mr.flor||'');
              if (!_mAgr[k]) _mAgr[k] = {rancho:mr.rancho, flor:mr.flor, metros:0, pla_acum:0};
              _mAgr[k].metros += mr.metros || 0;
              _mAgr[k].pla_acum += mr.pla_acum || 0;
            });
            mRows = Object.keys(_mAgr).map(function(k){ return _mAgr[k]; });
          }
          mRows.sort(function(a,b) {
            var rCmp = (a.rancho||'').localeCompare(b.rancho||'');
            if (rCmp !== 0) return rCmp;
            return (a.flor||'').localeCompare(b.flor||'');
          });
          if (mRows.length > 0) {
            var tbl = '<table style="width:100%; border-collapse:collapse; font-size:10px; font-family:var(--font-sans, sans-serif);">' +
              '<thead><tr>' +
                (!ranchFilter ? '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Rancho</th>' : '') +
                '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Variedad (Flor)</th>' +
                '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Metros</th>' +
                '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Plantas Acum.</th>' +
              '</tr></thead><tbody>';
            mRows.forEach(function(mr){
              tbl += '<tr style="border-bottom:1px solid #f1f5f9; background:#ffffff;">' +
                (!ranchFilter ? '<td style="color:#64748b;padding:6px 8px;font-weight:500;">'+mr.rancho+'</td>' : '') +
                '<td style="color:#334155;padding:6px 8px;font-weight:600;">'+mr.flor+'</td>' +
                '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:500;">'+mr.metros.toLocaleString('es-MX',{maximumFractionDigits:2})+'</td>' +
                '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:500;">'+mr.pla_acum.toLocaleString('es-MX',{maximumFractionDigits:2})+'</td>' +
              '</tr>';
            });
            tbl += '</tbody></table>';
            siembraRowsHtml += '<tr class="metros-detail" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:0;"><div style="border-left:3px solid #3b82f6; margin-left:24px; padding:8px 12px; background:#f8fafc; box-shadow:inset 0 1px 3px rgba(0,0,0,0.02);">'+tbl+'</div></td></tr>';
          } else {
            var availableWks = [];
            if (DATA.metros_acumulados) {
              var wks = {};
              DATA.metros_acumulados.forEach(function(x){
                if (isRanchAllowed(x.rancho)) wks[x.semana_fin] = true;
              });
              availableWks = Object.keys(wks).sort();
            }
            var wksTxt = availableWks.length > 0 ? availableWks.join(', ') : 'Ninguna';
            siembraRowsHtml += '<tr class="metros-detail" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:12px 24px; color:#64748b; font-size:10px; font-style:italic;">Sin datos para semanas ' + _wkKeysInRange.join(',') + '. Semanas registradas en Excel para ' + (ranchFilter||'Todos') + ': ' + wksTxt + '</td></tr>';
          }
        }

        if (isCharolas && DATA.plantas_metros) {
          var cRows = DATA.plantas_metros.filter(function(r){ 
            var sf = parseInt(r.semana_fin);
            return _wkKeysInRange.indexOf(sf) >= 0 && isRanchAllowed(r.rancho); 
          });
          // Si es rango, agrupar por rancho+flor y sumar
          if (_isRange && cRows.length > 0) {
            var _cAgr = {};
            cRows.forEach(function(cr) {
              var k = (cr.rancho||'') + '||' + (cr.flor||'');
              if (!_cAgr[k]) _cAgr[k] = {rancho:cr.rancho, flor:cr.flor, plantas:0, metros:0};
              _cAgr[k].plantas += cr.plantas || 0;
              _cAgr[k].metros += cr.metros || 0;
            });
            cRows = Object.keys(_cAgr).map(function(k){ return _cAgr[k]; });
          }
          cRows.sort(function(a,b) {
            var rCmp = (a.rancho||'').localeCompare(b.rancho||'');
            if (rCmp !== 0) return rCmp;
            return (a.flor||'').localeCompare(b.flor||'');
          });
          if (cRows.length > 0) {
            var tbl2 = '<table style="width:100%; border-collapse:collapse; font-size:10px; font-family:var(--font-sans, sans-serif);">' +
              '<thead><tr>' +
                (!ranchFilter ? '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Rancho</th>' : '') +
                '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Variedad (Flor)</th>' +
                '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Plantas</th>' +
                '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Metros</th>' +
              '</tr></thead><tbody>';
            cRows.forEach(function(cr){
              tbl2 += '<tr style="border-bottom:1px solid #f1f5f9; background:#ffffff;">' +
                (!ranchFilter ? '<td style="color:#64748b;padding:6px 8px;font-weight:500;">'+cr.rancho+'</td>' : '') +
                '<td style="color:#334155;padding:6px 8px;font-weight:600;">'+cr.flor+'</td>' +
                '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:500;">'+cr.plantas.toLocaleString('es-MX',{maximumFractionDigits:2})+'</td>' +
                '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:500;">'+cr.metros.toLocaleString('es-MX',{maximumFractionDigits:2})+'</td>' +
              '</tr>';
            });
            tbl2 += '</tbody></table>';
            siembraRowsHtml += '<tr class="charolas-detail" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:0;"><div style="border-left:3px solid #10b981; margin-left:24px; padding:8px 12px; background:#f8fafc; box-shadow:inset 0 1px 3px rgba(0,0,0,0.02);">'+tbl2+'</div></td></tr>';
          } else {
            var availableWks2 = [];
            if (DATA.plantas_metros) {
              var wks2 = {};
              DATA.plantas_metros.forEach(function(x){
                if (isRanchAllowed(x.rancho)) wks2[x.semana_fin] = true;
              });
              availableWks2 = Object.keys(wks2).sort();
            }
            var wksTxt2 = availableWks2.length > 0 ? availableWks2.join(', ') : 'Ninguna';
            siembraRowsHtml += '<tr class="charolas-detail" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:12px 24px; color:#64748b; font-size:10px; font-style:italic;">Sin datos para semanas ' + _wkKeysInRange.join(',') + '. Semanas registradas en Excel para ' + (ranchFilter||'Todos') + ': ' + wksTxt2 + '</td></tr>';
          }
        }

        // ── Detalle WEEKLY#### (inv/tallos por variedad de flor) ──────────
        // Agregar datos de TODAS las semanas en el rango
        if (isWeeklyDetail && DATA.detalle_weekly) {
          var _wdRows = [];
          _wkKeysInRange.forEach(function(_wkK) {
            var _wkSrcW = DATA.detalle_weekly[_wkK] || DATA.detalle_weekly[String(_wkK)] || null;
            if (_wkSrcW && _wkSrcW[m.k]) {
              _wkSrcW[m.k].forEach(function(r) {
                if (r.valor && r.valor !== 0) _wdRows.push(r);
              });
            }
          });
          // tallos_cos: filtrar por rancho seleccionado; si es Todos, agrupar por flor
          if (m.k === 'tallos_cos') {
            if (ranchFilter) {
              _wdRows = _wdRows.filter(function(r){ return r.rancho === ranchFilter; });
            }
            // Siempre agrupar por flor (sumar valores de múltiples semanas)
            var _cosMap = {};
            _wdRows.forEach(function(r){ _cosMap[r.flor] = (_cosMap[r.flor]||0) + r.valor; });
            _wdRows = Object.keys(_cosMap).map(function(f){ return {flor:f, valor:_cosMap[f]}; });
          } else if (_isRange) {
            // Para otros campos en rango, agrupar por flor sumando valores
            var _agrMap = {};
            _wdRows.forEach(function(r) {
              var k = r.flor + (r.proveedor ? '||' + r.proveedor : '');
              if (!_agrMap[k]) _agrMap[k] = {flor:r.flor, valor:0, proveedor:r.proveedor||null};
              _agrMap[k].valor += r.valor;
            });
            _wdRows = Object.keys(_agrMap).map(function(k){ return _agrMap[k]; });
          }
          _wdRows.sort(function(a,b) {
            var fCmp = (a.flor||'').localeCompare(b.flor||'');
            if (fCmp !== 0) return fCmp;
            return (a.proveedor||'').localeCompare(b.proveedor||'');
          });
          var _isByProveedor = (m.k === 'tallos_comp');
          // Etiquetas legibles para el encabezado de la columna de valor
          var _wdColLabels = {
            inv_inicial: 'Tallos',
            tallos_cos:  'Tallos Cosechados',
            tallos_proc: 'Tallos Recibidos',
            tallos_comp: 'Tallos',
            tallos_desp: 'Tallos',
            tallos_des:  'Tallos',
            inv_final:   'Tallos'
          };
          var _wdColLbl = _wdColLabels[m.k] || 'Valor';
          if (_wdRows.length > 0) {
            var _wdTotal = _wdRows.reduce(function(s,r){ return s + (r.valor||0); }, 0);
            var tblW = '<table style="width:100%; border-collapse:collapse; font-size:10px; font-family:var(--font-sans, sans-serif);">' +
              '<thead><tr>' +
                '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Variedad (Flor)</th>' +
                '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">'+_wdColLbl+'</th>' +
                (_isByProveedor
                  ? '<th style="text-align:left;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">Proveedor</th>'
                  : '<th style="text-align:right;color:#64748b;padding:6px 8px;border-bottom:2px solid #cbd5e1;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:0.5px;">%</th>'
                ) +
              '</tr></thead><tbody>';
            _wdRows.forEach(function(wr){
              var terceraCelda = _isByProveedor
                ? '<td style="color:#64748b;padding:6px 8px;font-size:9px;font-weight:500;">'+(wr.proveedor||'—')+'</td>'
                : (function(){ var pct = _wdTotal > 0 ? ((wr.valor/_wdTotal)*100).toFixed(1) : '—'; return '<td style="text-align:right;color:#94a3b8;padding:6px 8px;font-size:9px;">'+pct+'%</td>'; })();
              tblW += '<tr style="border-bottom:1px solid #f1f5f9; background:#ffffff;">' +
                '<td style="color:#334155;padding:6px 8px;font-weight:600;">'+(wr.flor||'')+'</td>' +
                '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:500;">'+Number(wr.valor).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>' +
                terceraCelda +
              '</tr>';
            });
            tblW += '<tr style="border-top:2px solid #cbd5e1; background:#f8fafc;">' +
              '<td style="color:#475569;padding:6px 8px;font-weight:700;font-size:9px;text-transform:uppercase;">Total</td>' +
              '<td style="text-align:right;color:#1e293b;padding:6px 8px;font-weight:700;font-size:11px;">'+Number(_wdTotal).toLocaleString('es-MX',{maximumFractionDigits:0})+'</td>' +
              '<td></td>' +
            '</tr>';
            tblW += '</tbody></table>';
            siembraRowsHtml += '<tr class="weekly-detail weekly-detail-'+m.k+'" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:0;"><div style="border-left:3px solid #6366f1; margin-left:24px; padding:8px 12px; background:#f8fafc; box-shadow:inset 0 1px 3px rgba(0,0,0,0.02);">'+tblW+'</div></td></tr>';
          } else {
            siembraRowsHtml += '<tr class="weekly-detail weekly-detail-'+m.k+'" style="display:none; background:#ffffff; border-bottom:1px solid #e2e8f0;"><td colspan="4" style="padding:12px 24px; color:#64748b; font-size:10px; font-style:italic;">Sin detalle disponible para semanas '+_wkKeysInRange.join(',')+' en Excel WEEKLY.</td></tr>';
          }
        }
      });
      siembraRowsHtml += 
        '<tr style="background:#f8fafc; border-bottom:2px solid #e2e8f0;">' +
          '<td colspan="4" style="padding:4px 0;"></td>' +
        '</tr>';
    }
  }

  // ── Zona 2: Tabla de productos ────────────────────────────────────
  var productSection = '';
  if (hideProducts && !siembraRowsHtml) {
    productSection = '<div style="padding:12px 10px; color:#94a3b8; font-size:11px; text-align:center;">Detalle de productos no disponible para MATERIAL VEGETAL.</div>';
  } else if (rows.length === 0 && !siembraRowsHtml) {
    productSection = '<div style="padding:12px 10px; color:#94a3b8; font-size:11px; text-align:center;">EN PROCESO DE SER CARGADO</div>';
  } else {
    if (!hideProducts) rows.sort(function(a,b){return Math.abs(b.gasto) - Math.abs(a.gasto);});
    var _wkCode = ((yr % 100) * 100) + wkStart;
    var _TC = (_wkCode >= 2502 && _wkCode <= 2520) ? 20 : 19; // tipo de cambio MXN → USD
    var _isUSD = state.currency === 'usd';
    window._prodTC = _TC;
    window._prodIsUSD = _isUSD;
    var _conv = function(mxn){ return _isUSD ? mxn / _TC : mxn; };
    var _fmtG = function(mxn){
      var v = _conv(mxn);
      if (!v || isNaN(v)) return '$0';
      var neg = v < 0, s = Math.abs(v);
      var str = _isUSD
        ? (neg?'-$':'$') + s.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})
        : fmt(mxn);
      return str;
    };
    var total = hideProducts ? 0 : rows.reduce(function(s,r){return s+r.gasto;},0);
    var panelMeta = hideProducts
      ? 'Solo resumen de siembra'
      : ('Gasto: <b style="color:#16a34a" class="gasto-total-val">'+_fmtG(total)+'</b>');
      
    var searchHtml = hideProducts ? '' : '<input type="text" class="prod-filter-input" placeholder="Buscar producto..." style="padding:3px 8px; border:1px solid #cbd5e1; border-radius:4px; font-size:11px; width:180px; margin-left:20px" oninput="filterProdRows(this)">';

    productSection =
      '<div style="flex-shrink:0; background:#f1f5f9; border-bottom:1px solid #e2e8f0; padding:6px 10px; display:flex; justify-content:space-between; align-items:center;">' +
        '<span style="font-size:11px; color:#475569;">'+panelMeta+'</span>' +
        searchHtml +
      '</div>' +
      '<div style="overflow-x:auto; scrollbar-width:thin;">' +
        '<table style="font-size:10px; width:100%; border-collapse:collapse;">' +
          '<thead><tr style="position:sticky;top:0;z-index:1;">' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600; white-space:nowrap;">UBICACI\u00d3N</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">PRODUCTO</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">UNID.</th>' +
            '<th style="text-align:right; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">GASTO '+(_isUSD?'USD':'MXN')+'</th>' +
          '</tr></thead><tbody>' +
          siembraRowsHtml;
    if (!hideProducts && rows.length > 0) {
      rows.forEach(function(r,i){
        var rowBg = (i%2===0)?'#ffffff':'#f8fafc';
        var safeProdName = r.producto ? r.producto.replace(/"/g,'&quot;').toLowerCase() : '';
        productSection += '<tr class="pt-prod-row" data-prod="'+safeProdName+'" data-costo="'+r.gasto+'" style="background:'+rowBg+'; border-bottom:1px solid #f1f5f9;">' +
          '<td style="padding:3px 6px; white-space:nowrap; font-weight:600; color:#0f172a;">'+r.ubicacion+'</td>' +
          '<td style="padding:3px 6px; color:#0f172a;">'+r.producto+'</td>' +
          '<td style="padding:3px 6px; color:#94a3b8; font-size:9px;">'+r.unidades+'</td>' +
          '<td style="padding:3px 6px; text-align:right; font-weight:700; color:#0f172a;">'+_fmtG(r.gasto)+'</td>' +
          '</tr>';
      });
    } else {
      productSection +=
        '<tr style="background:#ffffff;border-bottom:1px solid #f1f5f9;">' +
          '<td colspan="4" style="padding:8px 6px;color:#94a3b8;font-size:10px;text-align:center;">EN PROCESO DE SER CARGADO</td>' +
        '</tr>';
    }
    productSection += '</tbody></table></div>';
  }

  panelHtml =
    '<div class="prod-panel-wrapper" style="flex:1; min-width:340px; border:1px solid #cbd5e1; border-top:3px solid #7B1C1C; display:block; background:#fff;">' +
      '<div style="background:#7B1C1C; color:#fff; padding:5px 10px; flex-shrink:0; display:flex; justify-content:space-between; align-items:center;">' +
        '<div style="font-weight:700; font-size:11px; text-transform:uppercase; letter-spacing:0.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="'+panelTitle+'">'+panelTitle+'</div>' +
      '</div>' +
      productSection +
    '</div>';
  
  if (_prodViews.indexOf(panelHtml) === -1) {
    _prodViews.push(panelHtml);
    _prodViewsData.push({rowData:rowData, opts:opts||{}});
    if (_prodViews.length > 2) {
      _prodViews.shift(); // keep max 2 side-by-side
      _prodViewsData.shift();
    }
  }
  
  document.getElementById('prodPanel').className='show';
  document.getElementById('prodTableWrap').innerHTML = _prodViews.join('');
  setTimeout(resizeTable,80);
}
function closeProdPanel() { _prodViews = []; document.getElementById('prodPanel').className=''; setTimeout(resizeTable,60); }
function showProdFromCmp(yr,wk,ranch) { showProdPanel({_cat:state.cat,_year:yr,_week:wk,_fromWeek:wk,_toWeek:wk},{ranch:ranch||null}); }

// =======================================================
// RESIZE
// =======================================================
function resizeTable() {
  reportHeight();
}
window.addEventListener('resize', resizeTable);

// =======================================================
// HEIGHT REPORTING
// =======================================================
function reportHeight() {
  var body = document.body || {};
  var docEl = document.documentElement || {};
  var appEl = document.getElementById('app');
  var h = Math.max(
    window.innerHeight || 0,
    docEl.clientHeight || 0,
    docEl.scrollHeight || 0,
    docEl.offsetHeight || 0,
    body.scrollHeight || 0,
    body.offsetHeight || 0,
    appEl ? appEl.scrollHeight : 0,
    appEl ? appEl.offsetHeight : 0,
    700
  );
  window.parent.postMessage({type:'streamlit:setFrameHeight',height:h},'*');
}
var ro=new ResizeObserver(reportHeight);
ro.observe(document.getElementById('app')||document.body);
reportHeight();
// Reportar varias veces al inicio para que Streamlit capture la altura real
setTimeout(reportHeight,100);
setTimeout(reportHeight,300);
setTimeout(reportHeight,700);
setTimeout(reportHeight,1500);

// =======================================================
// ARRANCAR &#8212; diferido con protección
// =======================================================
if (!DATA || !DATA.weekly_series) {
  if (DATA) {
    DATA.weekly_series={};
    DATA.categories.forEach(function(cat){DATA.weekly_series[cat]={};});
    DATA.weekly_detail.forEach(function(r){
      if (r.usd_total>0){
        if (!DATA.weekly_series[r.categoria]) DATA.weekly_series[r.categoria]={};
        var key=r.year+'-W'+String(r.week).padStart(2,'0');
        DATA.weekly_series[r.categoria][key]=(DATA.weekly_series[r.categoria][key]||0)+r.usd_total;
      }
    });
  }
}

setTimeout(function() {
  if (!DATA) return;
  try {
    inicializar();
  } catch(e) {
    var loader = document.getElementById('loader');
    if (loader) loader.innerHTML =
      '<div style="color:#dc2626;font-family:monospace;padding:20px;background:#fff;' +
      'border-radius:8px;border:1px solid #fecaca;max-width:600px;margin:20px auto">' +
      '<b>Error en inicializar:</b> ' + e.message +
      (e.stack ? '<br><pre style="font-size:10px;color:#999;margin-top:8px;overflow:auto">' + e.stack + '</pre>' : '') +
      '</div>';
  }
}, 100);
</script>"""

HTML = f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CFBC &#8212; Control Operativo</title>
{APP_CSS}
</head>
<body>
{APP_HTML_BODY}
{APP_JS}
</body>
</html>'''

html_final = HTML.replace('__DATA_JSON__', data_json)


# ─── LÓGICA DE RENDERIZADO (DASHBOARD VS AUTOMATIZACIÓN) ─────────────────────
available_weeks = sorted(
    {str(r["year"] % 100).zfill(2) + str(r["week"]).zfill(2) for r in DATA.get("weekly_detail", [])},
    reverse=True
)

from data_extractor import get_sheet_xlsx
try:
    from data_extractor import crear_hoja_wk
    _crear_disponible = True
except ImportError:
    _crear_disponible = False

try:
    from data_extractor import insertar_hojas_pr_me_mp
    _subir_disponible = True
except ImportError:
    _subir_disponible = False

try:
    from data_extractor import autorrellenar_materiales_wk
    _autofill_disponible = True
except ImportError:
    _autofill_disponible = False

try:
    from data_extractor import autorrellenar_material_vegetal_wk
    _autofill_mv_disponible = True
except ImportError:
    _autofill_mv_disponible = False

try:
    from data_extractor import autorrellenar_nomina_wk
    _autofill_nomina_disponible = True
except ImportError:
    _autofill_nomina_disponible = False

try:
    from data_extractor import autorrellenar_conteo_marlen
    _autofill_conteo_disponible = True
except ImportError:
    _autofill_conteo_disponible = False

try:
    from data_extractor import autorrellenar_siembra_wk
    _autofill_siembra_disponible = True
except ImportError:
    _autofill_siembra_disponible = False

if not st.session_state.show_auto:
    # ==== MODO DASHBOARD ====
    # Header nativo de Streamlit — mismo color #7B1C1C, botones reales sin trucos de z-index
    col_brand, col_reload, col_auto = st.columns([10, 1, 1])
    with col_brand:
        st.markdown('<span id="cfbc-brand">CFBC &#9656; CONTROL SEMANAL</span>', unsafe_allow_html=True)
    with col_reload:
        if st.button("⟳ Recargar", key="btn_reload", help="Recargar datos desde SharePoint"):
            st.cache_data.clear()
            st.rerun()
    with col_auto:
        st.button("⚙️ Auto", key="btn_auto", on_click=toggle_auto, help="Panel de Automatización")

    # iframe sin header propio (36px ya los ocupa el header nativo de arriba)
    # Asignamos una altura gigante fija porque components.html puro no tiene auto-resize
    components.html(html_final, height=4000, scrolling=False)

else:
    # ==== MODO PANEL DE AUTOMATIZACION ====
    try:
        _sp_cfg = st.secrets["sharepoint"]
        _secrets_ok = all(k in _sp_cfg for k in ["tenant_id", "client_id", "client_secret"])
    except Exception:
        _secrets_ok = False

    active_modules = sum([
        _crear_disponible,
        _subir_disponible,
        _autofill_disponible,
        _autofill_mv_disponible,
        _autofill_siembra_disponible,
        _autofill_nomina_disponible,
        _autofill_conteo_disponible,
    ])
    latest_week = f"WK{available_weeks[0]}" if available_weeks else "Sin semanas"

    top_brand, top_reload, top_back = st.columns([8.2, 1.05, 1.2], gap="small")
    with top_brand:
        st.markdown(
            '''
            <div id="auto-topbar-left"></div>
            <div class="auto-topbar-wrap">
                <div class="auto-topbar-icon">≡</div>
                <div>
                    <div class="auto-topbar-kicker">CFBC Console</div>
                    <div class="auto-topbar-title">Panel de Automatizacion</div>
                    <div class="auto-topbar-subtitle">SharePoint, cargas semanales y procesos automáticos en una sola vista</div>
                </div>
            </div>
            ''',
            unsafe_allow_html=True,
        )
    with top_reload:
        st.markdown('<div id="auto-topbar-reload"></div>', unsafe_allow_html=True)
        if st.button("Recargar", key="btn_auto_reload", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with top_back:
        st.markdown('<div id="auto-topbar-back"></div>', unsafe_allow_html=True)
        st.button(
            "Dashboard",
            key="btn_auto_dashboard",
            type="secondary",
            on_click=toggle_auto,
            use_container_width=True,
        )

    side_col, main_col = st.columns([1.15, 4.85], gap="medium")

    with side_col:
        st.markdown(
            f'''
            <div id="auto-sidebar-shell"></div>
            <div class="auto-sidebar-logo">CFBC</div>
            <div class="auto-sidebar-caption">Automation Console</div>
            <div class="auto-sidebar-note">Centro Floricultor de Baja California. Panel operativo para hojas WK, cargas SharePoint y rutinas de autorrelleno.</div>

            <div class="auto-sidebar-group">Vista actual</div>
            <div class="auto-nav-item auto-nav-item-active">Panel Auto</div>

            <div class="auto-sidebar-group">Modulos</div>
            <div class="auto-nav-item">Gestion WK</div>
            <div class="auto-nav-item">Autorrellenos</div>
            <div class="auto-nav-item">Conteo</div>
            <div class="auto-nav-item">Carga de archivos</div>

            <div class="auto-sidebar-group">Estado</div>
            <div class="auto-sidebar-metric"><span>Semanas detectadas</span><strong>{len(available_weeks)}</strong></div>
            <div class="auto-sidebar-metric"><span>Ultima semana</span><strong>{latest_week}</strong></div>
            <div class="auto-sidebar-metric"><span>Modulos activos</span><strong>{active_modules}</strong></div>
            <div class="auto-sidebar-metric"><span>Credenciales</span><strong>{'OK' if _secrets_ok else 'Pend.'}</strong></div>
            ''',
            unsafe_allow_html=True,
        )

    with main_col:
        st.markdown(
            f'''
            <div id="auto-main-shell"></div>
            <div id="auto-hero-shell"></div>
            <div class="auto-eyebrow">Centro Floricultor de Baja California</div>
            <div class="auto-title">Centro de operaciones automaticas</div>
            <div class="auto-subtitle">Administra la descarga, creacion, carga y relleno de hojas WK desde una interfaz mas limpia, modular y ejecutiva. Todo el flujo queda agrupado por tipo de operacion para trabajar mas rapido y con menos errores.</div>
            <div class="auto-hero-badges">
                <span class="auto-hero-badge">SharePoint / Graph</span>
                <span class="auto-hero-badge">Credenciales: {'OK' if _secrets_ok else 'Pendiente'}</span>
                <span class="auto-hero-badge">Semana mas reciente: {latest_week}</span>
                <span class="auto-hero-badge">Modulos listos: {active_modules}</span>
            </div>
            ''',
            unsafe_allow_html=True,
        )

        stat_cols = st.columns(4, gap="small")
        stat_values = [
            ("auto-stat-1", str(len(available_weeks)), "Semanas detectadas", "Base disponible para operaciones WK."),
            ("auto-stat-2", "Lista" if _crear_disponible else "Off", "Crear WK", "Plantilla oficial lista para alta manual."),
            ("auto-stat-3", "Lista" if _subir_disponible else "Off", "Carga PR/MP/ME", "Insercion automatica al libro principal."),
            ("auto-stat-4", "OK" if _secrets_ok else "Pend.", "Credenciales", "Estado de acceso a SharePoint / Graph."),
        ]
        for col, (marker, value, label, note) in zip(stat_cols, stat_values):
            with col:
                st.markdown(
                    f'''<div id="{marker}"></div>
                    <div class="auto-stat-label">{label}</div>
                    <div class="auto-stat-value">{value}</div>
                    <div class="auto-stat-note">{note}</div>''',
                    unsafe_allow_html=True,
                )

        tab_wk, tab_autofill, tab_conteo, tab_upload = st.tabs([
            "Gestion WK",
            "Autorrellenos",
            "Conteo",
            "Carga de archivos",
        ])

        with tab_wk:
            col_down, col_create = st.columns(2, gap="medium")

            with col_down:
                st.markdown(
                    '''
                    <div id="auto-card-download"></div>
                    <div class="auto-card-kicker">Operacion WK</div>
                    <div class="auto-card-title">Descargar hoja semanal</div>
                    <div class="auto-card-note">Prepara un archivo individual con el formato oficial para revision o descarga inmediata.</div>
                    ''',
                    unsafe_allow_html=True,
                )
                if available_weeks:
                    selected_wk = st.selectbox(
                        "Semana disponible",
                        options=available_weeks,
                        format_func=lambda c: f"WK{c}",
                        key="auto_download_wk",
                    )

                    if st.button("Preparar archivo XLSX", use_container_width=True, key="btn_prepare_wk"):
                        with st.spinner(f"Conectando con SharePoint y preparando WK{selected_wk}..."):
                            xlsx_bytes = get_sheet_xlsx(selected_wk)
                        if xlsx_bytes:
                            st.success("Archivo listo para descarga.")
                            st.download_button(
                                label=f"Descargar WK{selected_wk}.xlsx",
                                data=xlsx_bytes,
                                file_name=f"WK{selected_wk}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True,
                            )
                        else:
                            st.error(f"No se encontro WK{selected_wk} en el servidor.")
                else:
                    st.warning("No hay semanas disponibles para descargar.")

            with col_create:
                st.markdown(
                    '''
                    <div id="auto-card-create"></div>
                    <div class="auto-card-kicker">Alta manual</div>
                    <div class="auto-card-title">Crear nueva hoja WK</div>
                    <div class="auto-card-note">Genera una hoja en blanco con celdas, formulas y formato corporativo directamente en SharePoint.</div>
                    ''',
                    unsafe_allow_html=True,
                )
                if _crear_disponible:
                    nuevo_nombre = st.text_input(
                        "Codigo de hoja",
                        placeholder="WK2518",
                        key="auto_new_wk_name",
                    ).strip().upper()

                    if st.button("Crear hoja en SharePoint", type="primary", use_container_width=True, key="btn_create_wk"):
                        if not nuevo_nombre:
                            st.warning("Escribe el nombre de la hoja.")
                        elif not nuevo_nombre.startswith("WK") or len(nuevo_nombre) != 6:
                            st.warning("El formato debe ser exactamente WK#### (ej: WK2518).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id     = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Escribiendo {nuevo_nombre} via Microsoft Graph API..."):
                                    resultado = crear_hoja_wk(nuevo_nombre, tenant_id, client_id, client_secret)
                                if resultado.get("ok"):
                                    st.success(resultado["mensaje"])
                                    st.cache_data.clear()
                                else:
                                    st.error(resultado["error"])
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                else:
                    st.error("La funcion de crear hojas no esta disponible en data_extractor.py")

        with tab_autofill:
            with st.container():
                st.markdown(
                    '''
                    <div id="auto-card-fill-shell"></div>
                    <div class="auto-card-kicker">Post-creacion</div>
                    <div class="auto-section-title">Autorrellenar materiales MN</div>
                    <div class="auto-section-note">Llena por rancho las filas de FERTILIZANTES, DESINFECCION / PLAGUICIDAS, MANTENIMIENTO y MATERIAL DE EMPAQUE usando PR / MP / ME de la misma semana. Los totales y USD siguen por formula.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _autofill_disponible:
                    st.error("La funcion `autorrellenar_materiales_wk` no esta disponible en data_extractor.py")
                else:
                    fill_sel_col, fill_manual_col, fill_info_col = st.columns([1.7, 1.05, 1.15], gap="small")
                    with fill_sel_col:
                        if available_weeks:
                            fill_wk_sel = st.selectbox(
                                "WK disponible",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="autofill_wk_sel",
                            )
                            fill_week_code = fill_wk_sel
                        else:
                            fill_week_code = ""
                    with fill_manual_col:
                        fill_week_manual = st.text_input(
                            "O captura WK",
                            placeholder="2614",
                            max_chars=4,
                            key="autofill_wk_manual",
                        ).strip()
                        if fill_week_manual:
                            fill_week_code = fill_week_manual
                    with fill_info_col:
                        fill_label = f"WK{fill_week_code}" if fill_week_code else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Accion</div>
                            <div class="auto-card-title">{fill_label}</div>
                            <div class="auto-card-note">Usa este paso despues de crear y revisar la hoja.</div>''',
                            unsafe_allow_html=True,
                        )

                    if st.button(
                        f"Autorrellenar materiales {'— WK' + fill_week_code if fill_week_code else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_autofill_materials",
                        disabled=not bool(fill_week_code),
                    ):
                        if not (fill_week_code.isdigit() and len(fill_week_code) == 4):
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2614).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Autorrellenando materiales en WK{fill_week_code}..."):
                                    res_fill = autorrellenar_materiales_wk(
                                        week_code=fill_week_code,
                                        tenant_id=tenant_id,
                                        client_id=client_id_sp,
                                        client_secret=client_secret,
                                    )
                                if res_fill.get("ok"):
                                    st.success(res_fill.get("mensaje", "Materiales autorrellenados correctamente."))
                                    st.cache_data.clear()
                                else:
                                    st.error(res_fill.get("error", "No se pudo autorrellenar la WK."))
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

            with st.container():
                st.markdown(
                    '''
                    <div id="auto-card-fill-mv-shell"></div>
                    <div class="auto-card-kicker">Post-creacion · MV</div>
                    <div class="auto-section-title">Autorrellenar material vegetal</div>
                    <div class="auto-section-note">Llena la fila de MATERIAL VEGETAL (fila 14) por rancho usando la hoja MV#### del Excel fuente. Requiere que ya hayas subido el archivo MV de la semana.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _autofill_mv_disponible:
                    st.error("La funcion `autorrellenar_material_vegetal_wk` no esta disponible en data_extractor.py")
                else:
                    mv_fill_sel_col, mv_fill_manual_col, mv_fill_info_col = st.columns([1.7, 1.05, 1.15], gap="small")
                    with mv_fill_sel_col:
                        if available_weeks:
                            mv_fill_wk_sel = st.selectbox(
                                "WK disponible",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="autofill_mv_wk_sel",
                            )
                            mv_fill_week_code = mv_fill_wk_sel
                        else:
                            mv_fill_week_code = ""
                    with mv_fill_manual_col:
                        mv_fill_week_manual = st.text_input(
                            "O captura WK",
                            placeholder="2614",
                            max_chars=4,
                            key="autofill_mv_wk_manual",
                        ).strip()
                        if mv_fill_week_manual:
                            mv_fill_week_code = mv_fill_week_manual
                    with mv_fill_info_col:
                        mv_fill_label = f"WK{mv_fill_week_code}" if mv_fill_week_code else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Accion</div>
                            <div class="auto-card-title">{mv_fill_label}</div>
                            <div class="auto-card-note">Fila 14 · Columnas E→K por rancho.</div>''',
                            unsafe_allow_html=True,
                        )

                    if st.button(
                        f"Autorrellenar material vegetal {'— WK' + mv_fill_week_code if mv_fill_week_code else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_autofill_mv",
                        disabled=not bool(mv_fill_week_code),
                    ):
                        if not (mv_fill_week_code.isdigit() and len(mv_fill_week_code) == 4):
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2614).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Autorrellenando Material Vegetal en WK{mv_fill_week_code}..."):
                                    res_mv = autorrellenar_material_vegetal_wk(
                                        week_code=mv_fill_week_code,
                                        tenant_id=tenant_id,
                                        client_id=client_id_sp,
                                        client_secret=client_secret,
                                    )
                                if res_mv.get("ok"):
                                    st.success(res_mv.get("mensaje", "Material Vegetal autorrellenado correctamente."))
                                    st.cache_data.clear()
                                else:
                                    st.error(res_mv.get("error", "No se pudo autorrellenar Material Vegetal."))
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

            with st.container():
                st.markdown(
                    '''
                    <div id="auto-card-fill-siembra-shell"></div>
                    <div class="auto-card-kicker">Post-creacion · Siembra</div>
                    <div class="auto-section-title">Autorrellenar metros y charolas</div>
                    <div class="auto-section-note">Llena por rancho las filas de Metros de Siembra (fila 91) y Numero de Charolas Sembradas (fila 89) sumando los datos desde el Excel de Siembra Detalle.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _autofill_siembra_disponible:
                    st.error("La funcion `autorrellenar_siembra_wk` no esta disponible en data_extractor.py")
                else:
                    siem_sel_col, siem_manual_col, siem_info_col = st.columns([1.7, 1.05, 1.15], gap="small")
                    with siem_sel_col:
                        if available_weeks:
                            siem_wk_sel = st.selectbox(
                                "WK disponible",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="autofill_siembra_wk_sel",
                            )
                            siem_week_code = siem_wk_sel
                        else:
                            siem_week_code = ""
                    with siem_manual_col:
                        siem_week_manual = st.text_input(
                            "O captura WK",
                            placeholder="2615",
                            max_chars=4,
                            key="autofill_siembra_wk_manual",
                        ).strip()
                        if siem_week_manual:
                            siem_week_code = siem_week_manual
                    with siem_info_col:
                        siem_fill_label = f"WK{siem_week_code}" if siem_week_code else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Accion</div>
                            <div class="auto-card-title">{siem_fill_label}</div>
                            <div class="auto-card-note">Escribe en fila 89 y 91.</div>''',
                            unsafe_allow_html=True,
                        )

                    if st.button(
                        f"Autorrellenar siembra {'— WK' + siem_week_code if siem_week_code else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_autofill_siembra",
                        disabled=not bool(siem_week_code),
                    ):
                        if not (siem_week_code.isdigit() and len(siem_week_code) == 4):
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2615).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Autorrellenando siembra en WK{siem_week_code}..."):
                                    res_siem = autorrellenar_siembra_wk(
                                        week_code=siem_week_code,
                                        tenant_id=tenant_id,
                                        client_id=client_id_sp,
                                        client_secret=client_secret,
                                    )
                                if res_siem.get("ok"):
                                    st.success(res_siem.get("mensaje", "Siembra autorrellenada correctamente."))
                                    st.cache_data.clear()
                                else:
                                    st.error(res_siem.get("error", "No se pudo automatizar la siembra."))
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

            with st.container():
                st.markdown(
                    '''
                    <div id="auto-card-fill-nomina-shell"></div>
                    <div class="auto-card-kicker">Post-creacion · Nomina</div>
                    <div class="auto-section-title">Automatizar nomina MN</div>
                    <div class="auto-section-note">Llena por rancho el bloque de nomina WK usando la hoja BD del SharePoint de nomina. Toma la columna MN #### de la semana capturada y suma PLANTA + CONTRATISTA cuando caen en el mismo concepto.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _autofill_nomina_disponible:
                    st.error("La funcion `autorrellenar_nomina_wk` no esta disponible en data_extractor.py")
                else:
                    nom_sel_col, nom_manual_col, nom_info_col = st.columns([1.7, 1.05, 1.15], gap="small")
                    with nom_sel_col:
                        if available_weeks:
                            nom_wk_sel = st.selectbox(
                                "WK disponible",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="autofill_nomina_wk_sel",
                            )
                            nom_week_code = nom_wk_sel
                        else:
                            nom_week_code = ""
                    with nom_manual_col:
                        nom_week_manual = st.text_input(
                            "O captura WK",
                            placeholder="2615",
                            max_chars=4,
                            key="autofill_nomina_wk_manual",
                        ).strip()
                        if nom_week_manual:
                            nom_week_code = nom_week_manual
                    with nom_info_col:
                        nom_fill_label = f"WK{nom_week_code}" if nom_week_code else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Accion</div>
                            <div class="auto-card-title">{nom_fill_label}</div>
                            <div class="auto-card-note">Usa la columna MN {nom_week_code or '####'} de la hoja BD y escribe en E→K.</div>''',
                            unsafe_allow_html=True,
                        )

                    if st.button(
                        f"Automatizar nomina {'— WK' + nom_week_code if nom_week_code else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_autofill_nomina",
                        disabled=not bool(nom_week_code),
                    ):
                        if not (nom_week_code.isdigit() and len(nom_week_code) == 4):
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2615).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Automatizando nomina en WK{nom_week_code}..."):
                                    res_nom = autorrellenar_nomina_wk(
                                        week_code=nom_week_code,
                                        tenant_id=tenant_id,
                                        client_id=client_id_sp,
                                        client_secret=client_secret,
                                    )
                                if res_nom.get("ok"):
                                    st.success(res_nom.get("mensaje", "Nomina autorrellenada correctamente."))
                                    st.cache_data.clear()
                                else:
                                    st.error(res_nom.get("error", "No se pudo automatizar la nomina."))
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

        with tab_conteo:
            with st.container():
                st.markdown(
                    '''
                    <div id="auto-card-fill-conteo-shell"></div>
                    <div class="auto-card-kicker">Conteo de Personal · Marlen</div>
                    <div class="auto-section-title">Autorellenar Conteo</div>
                    <div class="auto-section-note">Sube el archivo <strong>TT Nomina</strong> de la semana y llena automaticamente la hoja "Conteo" en SharePoint con el conteo de trabajadores por area y ubicacion.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _autofill_conteo_disponible:
                    st.error("La funcion `autorrellenar_conteo_marlen` no esta disponible en data_extractor.py")
                else:
                    cont_sel_col, cont_manual_col, cont_info_col = st.columns([1.7, 1.05, 1.15], gap="small")
                    with cont_sel_col:
                        if available_weeks:
                            cont_wk_sel = st.selectbox(
                                "WK disponible",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="autofill_conteo_wk_sel",
                            )
                            cont_week_code = str(cont_wk_sel)
                        else:
                            cont_week_code = ""
                    with cont_manual_col:
                        cont_week_manual = st.text_input(
                            "O captura WK",
                            placeholder="2615",
                            max_chars=4,
                            key="autofill_conteo_wk_manual",
                        ).strip()
                        if cont_week_manual:
                            cont_week_code = cont_week_manual
                    with cont_info_col:
                        cont_fill_label = f"WK{cont_week_code}" if cont_week_code else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Semana</div>
                            <div class="auto-card-title">{cont_fill_label}</div>
                            <div class="auto-card-note">Se escribira en la hoja "Conteo" del SharePoint.</div>''',
                            unsafe_allow_html=True,
                        )

                    st.markdown("<div id='auto-upload-tt'></div>", unsafe_allow_html=True)
                    tt_col, posco_col, vivero_col = st.columns([1, 1, 1], gap="medium")

                    with tt_col:
                        st.markdown(
                            '''<div class="auto-mini-title">TT Nomina</div>
                            <div class="auto-mini-note">Archivo principal de tiempo y asistencia</div>''',
                            unsafe_allow_html=True,
                        )
                        tt_uploaded = st.file_uploader(
                            f"Subir TT Nomina{' · WK' + cont_week_code if cont_week_code else ''}",
                            type=["xlsx", "xls"],
                            key="upload_tt_nomina",
                            help="Archivo Excel TT Nomina exportado del sistema.",
                        )
                        if tt_uploaded:
                            st.caption(f"OK · {tt_uploaded.name} · {round(tt_uploaded.size / 1024, 1)} KB")

                    with posco_col:
                        st.markdown(
                            '''<div class="auto-mini-title">Nomina Posco</div>
                            <div class="auto-mini-note">Nomina de Poscosecha · opcional</div>''',
                            unsafe_allow_html=True,
                        )
                        posco_uploaded = st.file_uploader(
                            f"Subir Nomina Posco{' · WK' + cont_week_code if cont_week_code else ''}",
                            type=["xlsx", "xls"],
                            key="upload_nomina_posco",
                            help="Archivo Excel de nomina de Poscosecha.",
                        )
                        if posco_uploaded:
                            st.caption(f"OK · {posco_uploaded.name} · {round(posco_uploaded.size / 1024, 1)} KB")

                    with vivero_col:
                        st.markdown(
                            '''<div class="auto-mini-title">Nomina Vivero</div>
                            <div class="auto-mini-note">Nomina de Propagacion · opcional</div>''',
                            unsafe_allow_html=True,
                        )
                        vivero_uploaded = st.file_uploader(
                            f"Subir Nomina Vivero{' · WK' + cont_week_code if cont_week_code else ''}",
                            type=["xlsx", "xls"],
                            key="upload_nomina_vivero",
                            help="Archivo Excel de nomina de Vivero / Propagacion.",
                        )
                        if vivero_uploaded:
                            st.caption(f"OK · {vivero_uploaded.name} · {round(vivero_uploaded.size / 1024, 1)} KB")

                    _cont_hay_archivo = tt_uploaded is not None
                    _cont_hay_semana  = bool(cont_week_code)
                    _cont_semana_ok   = cont_week_code.isdigit() and len(cont_week_code) == 4 if cont_week_code else False

                    if st.button(
                        f"Autorellenar Conteo {'— WK' + cont_week_code if cont_week_code else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_autofill_conteo",
                        disabled=not (_cont_hay_archivo and _cont_hay_semana),
                    ):
                        if not _cont_semana_ok:
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2615).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]
                                with st.spinner(f"Leyendo TT Nomina y actualizando Conteo WK{cont_week_code}..."):
                                    res_cont = autorrellenar_conteo_marlen(
                                        week_code     = cont_week_code,
                                        tt_file       = tt_uploaded,
                                        tenant_id     = tenant_id,
                                        client_id     = client_id_sp,
                                        client_secret = client_secret,
                                        posco_file    = posco_uploaded  if posco_uploaded  else None,
                                        vivero_file   = vivero_uploaded if vivero_uploaded else None,
                                    )
                                if res_cont.get("ok"):
                                    st.success(res_cont.get("mensaje", "Conteo actualizado correctamente."))
                                    st.cache_data.clear()
                                else:
                                    st.error(res_cont.get("error", "No se pudo autorellenar el Conteo."))
                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

                    if not _cont_hay_semana:
                        st.caption("Selecciona o captura una semana para habilitar.")
                    elif not _cont_hay_archivo:
                        st.caption("Sube el archivo TT Nomina para habilitar.")

        with tab_upload:
            with st.container():
                st.markdown(
                    '''
                    <div id="auto-upload-shell"></div>
                    <div class="auto-card-kicker">Carga consolidada</div>
                    <div class="auto-section-title">Subir PR / MP / ME / MV</div>
                    <div class="auto-section-note">Carga archivos fuente para una semana de trabajo y crea las hojas correspondientes dentro del Excel principal. ME admite 2 archivos fusionados.</div>
                    ''',
                    unsafe_allow_html=True,
                )

                if not _subir_disponible:
                    st.error("La funcion `insertar_hojas_pr_me_mp` no esta disponible en data_extractor.py")
                else:
                    ctrl_sel, ctrl_manual, ctrl_state = st.columns([1.8, 1.1, 0.9], gap="small")
                    with ctrl_sel:
                        if available_weeks:
                            semana_sel = st.selectbox(
                                "Semana base",
                                options=available_weeks,
                                format_func=lambda c: f"WK{c}",
                                key="upload_wk_sel",
                            )
                            semana_code_upload = semana_sel
                        else:
                            semana_code_upload = ""
                    with ctrl_manual:
                        semana_manual = st.text_input(
                            "O captura el codigo",
                            placeholder="2518",
                            max_chars=4,
                            key="upload_wk_manual",
                        ).strip()
                        if semana_manual:
                            semana_code_upload = semana_manual
                    with ctrl_state:
                        wk_label = f"WK{semana_code_upload}" if semana_code_upload else "Sin WK"
                        st.markdown(
                            f'''<div class="auto-card-kicker">Destino</div>
                            <div class="auto-card-title">{wk_label}</div>
                            <div class="auto-card-note">Semana objetivo para la insercion.</div>''',
                            unsafe_allow_html=True,
                        )

                    upload_pr_col, upload_mp_col, upload_me_col, upload_mv_col = st.columns([1, 1, 1.15, 1], gap="medium")

                    with upload_pr_col:
                        st.markdown(
                            '''
                            <div id="auto-upload-pr"></div>
                            <div class="auto-mini-title">PR</div>
                            <div class="auto-mini-note">Plaguicidas / riego</div>
                            ''',
                            unsafe_allow_html=True,
                        )
                        pr_uploaded = st.file_uploader(
                            f"Archivo PR{' · WK' + semana_code_upload if semana_code_upload else ''}",
                            type=["xlsx", "xls"],
                            key="upload_pr",
                            help="Un archivo Excel con los datos PR de la semana seleccionada.",
                        )
                        if pr_uploaded:
                            st.caption(f"OK · {pr_uploaded.name} ({round(pr_uploaded.size/1024,1)} KB)")

                    with upload_mp_col:
                        st.markdown(
                            '''
                            <div id="auto-upload-mp"></div>
                            <div class="auto-mini-title">MP</div>
                            <div class="auto-mini-note">Mantenimiento</div>
                            ''',
                            unsafe_allow_html=True,
                        )
                        mp_uploaded = st.file_uploader(
                            f"Archivo MP{' · WK' + semana_code_upload if semana_code_upload else ''}",
                            type=["xlsx", "xls"],
                            key="upload_mp",
                            help="Un archivo Excel con los datos MP de la semana seleccionada.",
                        )
                        if mp_uploaded:
                            st.caption(f"OK · {mp_uploaded.name} ({round(mp_uploaded.size/1024,1)} KB)")

                    with upload_me_col:
                        st.markdown(
                            '''
                            <div id="auto-upload-me"></div>
                            <div class="auto-mini-title">ME</div>
                            <div class="auto-mini-note">Material de empaque · 2 archivos opcionales</div>
                            ''',
                            unsafe_allow_html=True,
                        )
                        me1_uploaded = st.file_uploader(
                            "Archivo ME 1",
                            type=["xlsx", "xls"],
                            key="upload_me1",
                            help="Primer archivo Excel ME.",
                        )
                        if me1_uploaded:
                            st.caption(f"OK · {me1_uploaded.name} ({round(me1_uploaded.size/1024,1)} KB)")
                        me2_uploaded = st.file_uploader(
                            "Archivo ME 2",
                            type=["xlsx", "xls"],
                            key="upload_me2",
                            help="Segundo archivo Excel ME; se fusiona con el primero en la misma hoja.",
                        )
                        if me2_uploaded:
                            st.caption(f"OK · {me2_uploaded.name} ({round(me2_uploaded.size/1024,1)} KB)")

                    with upload_mv_col:
                        st.markdown(
                            '''
                            <div id="auto-upload-mv"></div>
                            <div class="auto-mini-title">MV</div>
                            <div class="auto-mini-note">Material Vegetal</div>
                            ''',
                            unsafe_allow_html=True,
                        )
                        mv_uploaded = st.file_uploader(
                            f"Archivo MV{' · WK' + semana_code_upload if semana_code_upload else ''}",
                            type=["xlsx", "xls"],
                            key="upload_mv",
                            help="Un archivo Excel con los datos MV de la semana seleccionada.",
                        )
                        if mv_uploaded:
                            st.caption(f"OK · {mv_uploaded.name} ({round(mv_uploaded.size/1024,1)} KB)")

                    st.markdown("---")

                    _hay_archivos  = any([pr_uploaded, mp_uploaded, me1_uploaded, me2_uploaded, mv_uploaded])
                    _hay_semana    = bool(semana_code_upload)
                    _semana_valida = semana_code_upload.isdigit() and len(semana_code_upload) == 4

                    if st.button(
                        f"Crear hojas en SharePoint {'— WK' + semana_code_upload if semana_code_upload else ''}",
                        type="primary",
                        use_container_width=True,
                        key="btn_subir_pr_me_mp",
                        disabled=not (_hay_archivos and _hay_semana),
                    ):
                        if not _semana_valida:
                            st.warning("El codigo de semana debe ser exactamente 4 digitos (ej: 2613).")
                        else:
                            try:
                                tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                                client_id_sp  = st.secrets["sharepoint"]["client_id"]
                                client_secret = st.secrets["sharepoint"]["client_secret"]

                                tipos_subidos = []
                                if pr_uploaded:
                                    tipos_subidos.append("PR")
                                if mp_uploaded:
                                    tipos_subidos.append("MP")
                                if me1_uploaded or me2_uploaded:
                                    tipos_subidos.append("ME")
                                if mv_uploaded:
                                    tipos_subidos.append("MV")

                                with st.spinner(
                                    f"Conectando con SharePoint y creando hojas {', '.join(tipos_subidos)} para WK{semana_code_upload}..."
                                ):
                                    res = insertar_hojas_pr_me_mp(
                                        semana_code   = semana_code_upload,
                                        tenant_id     = tenant_id,
                                        client_id     = client_id_sp,
                                        client_secret = client_secret,
                                        pr_file       = pr_uploaded  if pr_uploaded  else None,
                                        mp_file       = mp_uploaded  if mp_uploaded  else None,
                                        me_file1      = me1_uploaded if me1_uploaded else None,
                                        me_file2      = me2_uploaded if me2_uploaded else None,
                                        mv_file       = mv_uploaded  if mv_uploaded  else None,
                                    )

                                hubo_error = False
                                for tipo in ["PR", "MP", "ME", "MV"]:
                                    info = res.get(tipo, {})
                                    ok   = info.get("ok")
                                    msg  = info.get("msg", "")
                                    if ok is True:
                                        st.success(msg)
                                    elif ok is False:
                                        st.error(msg)
                                        hubo_error = True

                                if not hubo_error:
                                    st.cache_data.clear()
                                    st.info("Recarga el dashboard para reflejar los nuevos datos.")
                                else:
                                    st.warning("Algunos archivos no se pudieron subir. Revisa los mensajes mostrados.")

                            except KeyError as e:
                                st.error(f"Falta configurar la credencial en secrets.toml: {e}.")
                            except Exception as e:
                                st.error(f"Error inesperado: {e}")

                    if not _hay_semana:
                        st.caption("Selecciona o captura una semana para habilitar la carga.")
                    elif not _hay_archivos:
                        st.caption("Sube al menos un archivo para habilitar la creacion de hojas.")
