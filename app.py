"""
app.py
Centro Floricultor de Baja California
Streamlit — Tablas Dinámicas Ejecutivas con AG Grid
"""

import json
import base64
import os
import time
import re
import logging
import streamlit as st
import streamlit.components.v1 as components

from data_extractor import get_datos

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="CFBC WK",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  #MainMenu, header, footer { display: none !important; }
  .stApp { background: #f0f0f0; }
  .block-container { padding: 0 !important; max-width: 100% !important; }
  section[data-testid="stSidebar"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


def validate_credentials():
    """Valida que todas las credenciales requeridas estén presentes."""
    required_keys = ["tenant_id", "client_id", "client_secret"]
    missing = [k for k in required_keys if k not in st.secrets.get("sharepoint", {})]
    
    if missing:
        st.error(f"❌ Faltan credenciales: {', '.join(missing)}")
        st.info("Configura `.streamlit/secrets.toml` con la sección [sharepoint]")
        st.stop()


@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    """Carga datos con manejo granular de errores."""
    logger.info("Iniciando carga de datos desde SharePoint")
    try:
        data = get_datos()
        logger.info(f"Datos cargados: {len(data.get('weekly_detail', []))} registros")
        return data
    except FileNotFoundError:
        return {"error": "Archivo no encontrado en SharePoint. Verifica la ruta."}
    except PermissionError:
        return {"error": "Sin permisos para acceder al archivo. Revisa credenciales."}
    except Exception as e:
        logger.error(f"Error cargando datos: {str(e)}", exc_info=True)
        return {"error": f"Error inesperado: {str(e)}"}


try:
    DATA = load_data()
except Exception as e:
    st.error(f"❌ Error cargando datos: {e}")
    st.stop()

if "error" in DATA:
    st.error(f"❌ {DATA['error']}")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Reintentar", key="retry_data"):
            st.cache_data.clear()
            st.rerun()
    with col2:
        if st.button("📋 Ver detalles técnicos", key="show_logs"):
            with st.expander("Logs"):
                st.code(DATA.get("traceback", "No disponible"))
    st.stop()

# Optimización: Renderizar HTML una sola vez
if "html_rendered" not in st.session_state:
    data_json = base64.b64encode(
        json.dumps(DATA, ensure_ascii=True, default=str).encode('utf-8')
    ).decode('ascii')

    HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CFBC — Control Operativo Ejecutivo</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-grid.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-theme-alpine.css">
<script src="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/dist/ag-grid-community.min.js"></script>
<style>
:root {
  --navy: #1e3a5f;
  --green: #16a34a;
  --red: #dc2626;
  --amber: #b45309;
  --blue: #2563eb;
  --border: #d0d0d0;
  --mono: 'Consolas','Courier New',monospace;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--mono); font-size: 12px; background: #f0f0f0; overflow-x: hidden; }

/* ── LOADER ─────────────────────────────────────── */
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

/* ── HEADER — Executive Dashboard Style ─────────────────── */
.app-hdr {
  background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8f 100%);
  border-bottom: 4px solid var(--green);
  padding: 12px 20px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
.hdr-top {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 10px;
}
.hdr-brand {
  color: #fff; font-size: 18px; font-weight: 700;
  letter-spacing: 1.5px; text-shadow: 0 1px 2px rgba(0,0,0,0.3);
}
.hdr-actions {
  display: flex; gap: 8px;
}
.hdr-btn {
  font-size: 11px; font-family: var(--mono); font-weight: 700;
  background: rgba(255,255,255,0.15);
  border: 1px solid rgba(255,255,255,0.3);
  border-radius: 6px;
  padding: 6px 14px; cursor: pointer; color: #fff;
  transition: all 0.2s;
  box-shadow: 0 2px 4px rgba(0,0,0,0.2);
}
.hdr-btn:hover { 
  background: rgba(255,255,255,0.25); 
  transform: translateY(-1px);
  box-shadow: 0 4px 8px rgba(0,0,0,0.3);
}

/* ── KPI CARDS ────────────────────────────────────── */
.kpi-strip {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
}
.kpi-card {
  background: rgba(255,255,255,0.12);
  border: 1px solid rgba(255,255,255,0.2);
  border-radius: 8px;
  padding: 10px 14px;
  backdrop-filter: blur(10px);
  transition: all 0.2s;
}
.kpi-card:hover {
  background: rgba(255,255,255,0.18);
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0,0,0,0.3);
}
.kpi-label {
  color: rgba(255,255,255,0.7);
  font-size: 9px;
  text-transform: uppercase;
  letter-spacing: 0.8px;
  margin-bottom: 4px;
}
.kpi-value {
  color: #fff;
  font-size: 20px;
  font-weight: 700;
  margin-bottom: 2px;
  text-shadow: 0 1px 2px rgba(0,0,0,0.3);
}
.kpi-delta {
  font-size: 11px;
  display: flex;
  align-items: center;
  gap: 4px;
}
.kpi-delta.pos { color: #4ade80; }
.kpi-delta.neg { color: #f87171; }
.kpi-delta-icon { font-size: 14px; }

/* ── TOOLBAR — Enhanced Controls ──────────────────────── */
.toolbar {
  background: linear-gradient(to bottom, #f8f8f8, #ebebeb);
  border-bottom: 2px solid var(--border);
  padding: 10px 16px;
  display: flex; align-items: center; gap: 12px;
  flex-wrap: wrap;
  box-shadow: inset 0 -1px 3px rgba(0,0,0,0.05);
}
.tb-group {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 0 12px;
  border-right: 1px solid #ccc;
}
.tb-group:last-child { border-right: none; }
.tb-label {
  font-size: 10px;
  color: #666;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  font-weight: 600;
}
select.tb-sel {
  font-size: 11px;
  font-family: var(--mono);
  background: #fff;
  border: 2px solid #bbb;
  border-radius: 6px;
  padding: 5px 10px;
  color: #222;
  cursor: pointer;
  transition: all 0.2s;
  min-width: 140px;
}
select.tb-sel:hover { border-color: var(--green); }
select.tb-sel:focus {
  outline: none;
  border-color: var(--green);
  box-shadow: 0 0 0 3px rgba(22,163,74,0.1);
}
.tb-btn {
  font-size: 11px;
  font-family: var(--mono);
  font-weight: 600;
  background: #fff;
  border: 2px solid #bbb;
  border-radius: 6px;
  padding: 5px 12px;
  cursor: pointer;
  color: #333;
  transition: all 0.2s;
}
.tb-btn:hover {
  background: #f0f0f0;
  border-color: var(--navy);
}
.tb-btn.active {
  background: var(--navy);
  color: #fff;
  border-color: var(--navy);
  box-shadow: 0 2px 6px rgba(30,58,95,0.3);
}
.tb-grp { display: flex; }
.tb-grp .tb-btn { border-radius: 0; margin-left: -2px; }
.tb-grp .tb-btn:first-child { border-radius: 6px 0 0 6px; margin-left: 0; }
.tb-grp .tb-btn:last-child { border-radius: 0 6px 6px 0; }
.week-ctr {
  display: flex;
  align-items: center;
  gap: 8px;
  background: #fff;
  border: 2px solid #ddd;
  border-radius: 8px;
  padding: 4px 10px;
}
.week-ctr span {
  font-size: 13px;
  font-weight: 700;
  color: var(--navy);
  min-width: 70px;
  text-align: center;
}
.tb-slider {
  width: 140px;
  accent-color: var(--green);
  cursor: pointer;
}
.yr-chip {
  font-size: 11px;
  font-family: var(--mono);
  font-weight: 700;
  padding: 5px 12px;
  border-radius: 6px;
  cursor: pointer;
  border: 2px solid transparent;
  background: #fff;
  transition: all 0.2s;
}
.yr-chip:hover {
  transform: translateY(-1px);
  box-shadow: 0 2px 6px rgba(0,0,0,0.15);
}
.yr-chip.on {
  border-color: currentColor;
  box-shadow: 0 2px 8px currentColor;
}

/* ── RANGE CONTROL BAR ────────────────────────── */
.range-bar {
  display: none;
  background: linear-gradient(to bottom, #f4f4f4, #e8e8e8);
  border-bottom: 1px solid var(--border);
  padding: 8px 16px;
  align-items: center;
  gap: 12px;
}
.range-bar.show { display: flex; }
.range-val {
  font-size: 12px;
  font-weight: 700;
  color: var(--navy);
  font-family: var(--mono);
  min-width: 40px;
  text-align: center;
}
.range-badge {
  font-size: 11px;
  font-family: var(--mono);
  font-weight: 600;
  background: linear-gradient(135deg, #e8f5e9, #c8e6c9);
  border: 2px solid #81c784;
  color: var(--green);
  padding: 4px 12px;
  border-radius: 6px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

/* ── VIEW TABS — Executive Style ───────────────────────────────── */
.view-tabs {
  background: #fff;
  border-bottom: 3px solid #d5d5d5;
  display: flex;
  padding: 0;
  box-shadow: 0 2px 4px rgba(0,0,0,0.08);
}
.vtab {
  padding: 12px 20px;
  font-size: 11px;
  font-weight: 700;
  font-family: var(--mono);
  cursor: pointer;
  border: none;
  background: transparent;
  color: #888;
  border-bottom: 3px solid transparent;
  margin-bottom: -3px;
  text-transform: uppercase;
  letter-spacing: 0.8px;
  transition: all 0.2s;
  position: relative;
}
.vtab:hover {
  color: #333;
  background: rgba(22,163,74,0.05);
}
.vtab.active {
  color: var(--green);
  border-bottom-color: var(--green);
  background: #fff;
}
.vtab.active::after {
  content: '';
  position: absolute;
  bottom: -3px;
  left: 0;
  right: 0;
  height: 3px;
  background: var(--green);
  box-shadow: 0 2px 8px rgba(22,163,74,0.4);
}

/* ── GRID CONTAINER ──────────────────────────────── */
#gridWrap {
  background: #fff;
  border: 1px solid #d5d5d5;
  border-top: none;
  box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}

/* ── STATUS BAR ──────────────────────────────────── */
.statusbar {
  background: linear-gradient(to right, #ebebeb, #f5f5f5);
  border-top: 2px solid #d5d5d5;
  padding: 8px 16px;
  font-size: 11px;
  color: #666;
  display: flex;
  gap: 20px;
  align-items: center;
  font-weight: 500;
  box-shadow: inset 0 1px 3px rgba(0,0,0,0.05);
}
.statusbar b { color: #333; font-weight: 700; }
.statusbar .st-sep { color: #bbb; margin: 0 4px; }

/* ── PRODUCTOS PANEL ─────────────────────────────── */
#prodPanel {
  display: none;
  background: #fff;
  border-top: 3px solid var(--green);
  box-shadow: 0 -4px 12px rgba(0,0,0,0.1);
}
#prodPanel.show { display: block; }
.prod-hdr {
  background: linear-gradient(135deg, #1e3a5f, #2d5a8f);
  padding: 10px 16px;
  display: flex;
  align-items: center;
  gap: 12px;
}
.prod-hdr-title {
  color: #fff;
  font-size: 13px;
  font-weight: 700;
  letter-spacing: 0.8px;
  flex: 1;
}
.prod-hdr-meta {
  color: rgba(255,255,255,0.75);
  font-size: 11px;
}
.prod-close {
  background: rgba(255,255,255,0.15);
  border: 2px solid rgba(255,255,255,0.3);
  border-radius: 6px;
  color: #fff;
  cursor: pointer;
  font-size: 11px;
  padding: 4px 12px;
  font-family: var(--mono);
  font-weight: 600;
  transition: all 0.2s;
}
.prod-close:hover {
  background: rgba(255,255,255,0.25);
  border-color: #fff;
}

/* ── AG GRID THEME — Executive Pivot Tables ─────────────────────── */
.ag-theme-alpine {
  --ag-font-family: 'Consolas', 'Courier New', monospace;
  --ag-font-size: 11px;
  --ag-row-height: 32px;
  --ag-header-height: 36px;
  --ag-cell-horizontal-padding: 10px;
  --ag-borders: solid 1px;
  --ag-border-color: #e0e0e0;
  --ag-secondary-border-color: #f0f0f0;
  --ag-header-background-color: linear-gradient(to bottom, #f8f8f8, #e8e8e8);
  --ag-header-foreground-color: #333;
  --ag-odd-row-background-color: #fafafa;
  --ag-even-row-background-color: #ffffff;
  --ag-row-hover-color: #e8f5e9;
  --ag-selected-row-background-color: #c8e6c9;
  --ag-alpine-active-color: #16a34a;
  --ag-input-focus-border-color: #16a34a;
  --ag-range-selection-border-color: #16a34a;
}
.ag-theme-alpine .ag-header-cell {
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  border-right: 1px solid #d0d0d0;
  background: linear-gradient(to bottom, #f5f5f5, #e8e8e8);
}
.ag-theme-alpine .ag-header-cell-label {
  justify-content: center;
}
.ag-theme-alpine .ag-pinned-left-cols-container {
  border-right: 3px solid #aaa !important;
  box-shadow: 2px 0 4px rgba(0,0,0,0.1);
}
.ag-theme-alpine .ag-row-group {
  background: linear-gradient(to right, #eff3fa, #e6eefa) !important;
  font-weight: 700;
  border-left: 4px solid var(--green);
}
.ag-theme-alpine .ag-row-group-indent-1 {
  padding-left: 20px;
}
.ag-theme-alpine .ag-row-group-indent-2 {
  padding-left: 40px;
}
.ag-theme-alpine .ag-row-group-leaf-indent {
  padding-left: 15px;
}
.ag-theme-alpine .ag-group-expanded .ag-icon-tree-open,
.ag-theme-alpine .ag-group-contracted .ag-icon-tree-closed {
  font-size: 14px;
  color: var(--green);
}

/* Row Styles */
.ag-theme-alpine .ag-row {
  border-bottom: 1px solid #f0f0f0;
}
.ag-theme-alpine .ag-row-level-0 {
  background: linear-gradient(to right, #f0faf4, #fafafa) !important;
  font-weight: 700;
  border-top: 2px solid #d0d0d0;
  border-left: 4px solid #16a34a;
}
.ag-theme-alpine .ag-row-level-1 {
  background: #fafafa !important;
  font-weight: 600;
  border-left: 3px solid #81c784;
}
.ag-theme-alpine .ag-row-level-2 {
  background: #ffffff !important;
  padding-left: 10px;
}

/* Total Rows */
.ag-theme-alpine .ag-row.total-row {
  background: linear-gradient(135deg, rgba(22,163,74,0.08), rgba(22,163,74,0.04)) !important;
  font-weight: 700;
  border-top: 2px solid rgba(22,163,74,0.3);
  border-bottom: 2px solid rgba(22,163,74,0.3);
}
.ag-theme-alpine .ag-row.grand-total-row {
  background: linear-gradient(135deg, rgba(30,58,95,0.12), rgba(30,58,95,0.06)) !important;
  font-weight: 700;
  font-size: 12px;
  border-top: 3px solid var(--navy);
  border-bottom: 3px solid var(--navy);
}

/* Cell Styles */
.cell-pos { color: #16a34a !important; font-weight: 700; }
.cell-neg { color: #dc2626 !important; font-weight: 700; }
.cell-muted { color: #999 !important; }
.cell-total { font-weight: 700 !important; color: #1e3a5f !important; font-size: 12px; }
.cell-highlight {
  background: linear-gradient(90deg, rgba(22,163,74,0.1), transparent);
  font-weight: 600;
}
.prod-link {
  cursor: pointer;
  text-decoration: underline dotted;
  text-underline-offset: 2px;
  transition: all 0.2s;
}
.prod-link:hover {
  color: var(--green);
  text-decoration: underline solid;
}

/* ── COMPARATIVO TABLE — Executive Pivot ───────────────────────────── */
#comparativoWrap {
  display: none;
  background: #fff;
  border: 1px solid #d5d5d5;
  border-top: none;
  overflow: hidden;
  box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}
#comparativoWrap.show { display: block; }
.cmp-stat-strip {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 12px;
  padding: 16px;
  background: linear-gradient(to bottom, #f8f8f8, #f0f0f0);
  border-bottom: 2px solid #d5d5d5;
}
.cmp-stat-box {
  background: #fff;
  border: 2px solid #e0e0e0;
  border-radius: 8px;
  padding: 12px 16px;
  transition: all 0.2s;
  box-shadow: 0 2px 4px rgba(0,0,0,0.06);
}
.cmp-stat-box:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0,0,0,0.12);
  border-color: var(--green);
}
.cmp-stat-label {
  font-size: 9px;
  text-transform: uppercase;
  letter-spacing: 0.6px;
  color: #888;
  font-weight: 600;
}
.cmp-stat-val {
  font-size: 18px;
  font-weight: 700;
  margin: 4px 0 2px;
  color: var(--navy);
}
.cmp-stat-sub {
  font-size: 10px;
  color: #aaa;
}
.cmp-tbl-wrap {
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
  scrollbar-width: thin;
  scrollbar-color: #ccc transparent;
  max-height: calc(100vh - 320px);
  overflow-y: auto;
}
.cmp-tbl-wrap::-webkit-scrollbar { height: 6px; width: 6px; }
.cmp-tbl-wrap::-webkit-scrollbar-thumb {
  background: linear-gradient(to bottom, #bbb, #999);
  border-radius: 3px;
}
.cmp-tbl-wrap::-webkit-scrollbar-track { background: #f0f0f0; }
.cmp-tbl {
  border-collapse: collapse;
  width: 100%;
  font-family: var(--mono);
  font-size: 11px;
}
.cmp-tbl th {
  padding: 10px 12px;
  background: linear-gradient(to bottom, #f5f5f5, #e8e8e8);
  color: #444;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  white-space: nowrap;
  border-bottom: 2px solid #ccc;
  border-right: 1px solid #ddd;
  position: sticky;
  top: 0;
  z-index: 2;
  text-align: right;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}
.cmp-tbl th:first-child,
.cmp-tbl th:nth-child(2) { text-align: left; }
.cmp-tbl td {
  padding: 8px 12px;
  border-bottom: 1px solid #f0f0f0;
  border-right: 1px solid #f5f5f5;
  white-space: nowrap;
  text-align: right;
}
.cmp-tbl td:first-child,
.cmp-tbl td:nth-child(2) { text-align: left; }
.cmp-grp-hdr td {
  background: linear-gradient(to right, #eff3fa, #e6eefa);
  font-weight: 700;
  border-top: 2px solid #ccc;
  font-size: 11px;
  padding: 10px 12px;
}
.cmp-grp-hdr td:first-child {
  border-left: 4px solid var(--green);
}
.cmp-row:hover td {
  background: #f0faf4;
}
.cmp-total-row td {
  background: linear-gradient(135deg, rgba(22,163,74,0.12), rgba(22,163,74,0.06));
  font-weight: 700;
  border-top: 2px solid rgba(22,163,74,0.3);
  border-bottom: 2px solid rgba(22,163,74,0.3);
  color: var(--green);
}
.cmp-total-row td:first-child {
  border-left: 4px solid rgba(22,163,74,0.6);
}
.delta-cell { font-size: 10px; white-space: nowrap; }
.delta-amt { display: block; }
.delta-pct { display: block; font-size: 9px; opacity: 0.85; }
.chg-pos { color: #16a34a; font-weight: 700; }
.chg-neg { color: #dc2626; font-weight: 700; }
.chg-0 { color: #aaa; }
.cmp-clickable {
  cursor: pointer;
  transition: all 0.2s;
}
.cmp-clickable:hover {
  background: rgba(22,163,74,0.1) !important;
  font-weight: 600;
}
</style>
</head>
<body>

<!-- LOADER -->
<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">CFBC — Cargando sistema ejecutivo...</div>
</div>

<!-- APP -->
<div id="app" style="display:none">

  <!-- HEADER EJECUTIVO -->
  <div class="app-hdr">
    <div class="hdr-top">
      <div class="hdr-brand">📊 CFBC — CONTROL EJECUTIVO</div>
      <div class="hdr-actions">
        <button class="hdr-btn" onclick="exportCSV()">⬇ EXPORTAR CSV</button>
        <button class="hdr-btn" onclick="recargar()">⟳ ACTUALIZAR</button>
      </div>
    </div>
    <div class="kpi-strip" id="hdrKpis"></div>
  </div>

  <!-- TOOLBAR -->
  <div class="toolbar">
    <div class="tb-group">
      <span class="tb-label">📂 Categoría</span>
      <select class="tb-sel" id="catSel" onchange="onCatChange(this.value)"></select>
    </div>
