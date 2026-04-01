"""
app.py
Centro Floricultor de Baja California
Streamlit — AG Grid data-dense, Excel-style, full-screen
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

st.markdown("""
<style>
  #MainMenu, header, footer { display: none !important; }
  .stApp { background: #f0f0f0; }
  .block-container { padding: 0 !important; max-width: 100% !important; }
  section[data-testid="stSidebar"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    return get_datos()


try:
    DATA = load_data()
except Exception as e:
    st.error(f"❌ Error cargando datos: {e}")
    st.stop()

if "error" in DATA:
    st.error(f"❌ {DATA['error']}")
    if st.button("🔄 Reintentar"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

data_json = base64.b64encode(
    json.dumps(DATA, ensure_ascii=True, default=str).encode('utf-8')
).decode('ascii')

HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CFBC — Control Operativo</title>
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

/* ── HEADER — single compact bar ─────────────────── */
.app-hdr {
  background: var(--navy);
  border-bottom: 3px solid var(--green);
  padding: 5px 10px;
  display: flex;
  align-items: center;
  gap: 0;
  height: 36px;
  overflow: hidden;
}
.hdr-brand {
  color: #fff; font-size: 12px; font-weight: 700;
  letter-spacing: 1px; white-space: nowrap;
  padding-right: 12px; border-right: 1px solid rgba(255,255,255,0.2);
  flex-shrink: 0;
}
.hdr-kpis { display: flex; gap: 0; flex: 1; overflow: hidden; min-width: 0; }
.hdr-kpi {
  padding: 0 12px;
  border-right: 1px solid rgba(255,255,255,0.12);
  display: flex; align-items: center; gap: 8px;
  white-space: nowrap; flex-shrink: 0;
}
.hdr-kpi-label { color: rgba(255,255,255,0.45); font-size: 9px; text-transform: uppercase; letter-spacing: 0.5px; }
.hdr-kpi-val { color: #fff; font-size: 12px; font-weight: 700; }
.hdr-kpi-delta { font-size: 10px; }
.hdr-kpi-delta.pos { color: #4ade80; }
.hdr-kpi-delta.neg { color: #f87171; }
.hdr-btn {
  margin-left: auto; flex-shrink: 0;
  font-size: 10px; font-family: var(--mono); font-weight: 700;
  background: rgba(255,255,255,0.1);
  border: 1px solid rgba(255,255,255,0.25);
  border-radius: 3px;
  padding: 3px 10px; cursor: pointer; color: rgba(255,255,255,0.85);
  height: 24px; transition: background 0.1s; white-space: nowrap;
}
.hdr-btn:hover { background: rgba(255,255,255,0.2); }

/* ── TOOLBAR — controls row ──────────────────────── */
.toolbar {
  background: #ebebeb;
  border-bottom: 1px solid var(--border);
  padding: 4px 8px;
  display: flex; align-items: center; gap: 6px;
  flex-wrap: nowrap; overflow-x: auto; scrollbar-width: none;
  height: 32px;
}
.toolbar::-webkit-scrollbar { display: none; }
.tb-label {
  font-size: 9px; color: #777;
  text-transform: uppercase; letter-spacing: 0.5px;
  white-space: nowrap; flex-shrink: 0;
}
.tb-sep { width: 1px; height: 18px; background: #ccc; flex-shrink: 0; }
select.tb-sel {
  font-size: 11px; font-family: var(--mono);
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 2px 6px; color: #222; cursor: pointer; height: 22px;
  flex-shrink: 0;
}
select.tb-sel:focus { outline: 2px solid var(--green); outline-offset: -1px; }
.tb-btn {
  font-size: 10px; font-family: var(--mono); font-weight: 700;
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 2px 8px; cursor: pointer; height: 22px;
  white-space: nowrap; color: #333; transition: background 0.1s; flex-shrink: 0;
}
.tb-btn:hover { background: #ddd; }
.tb-btn.active { background: var(--navy); color: #fff; border-color: var(--navy); }
.tb-btn.green-active { background: var(--green); color: #fff; border-color: var(--green); }
.tb-grp { display: flex; flex-shrink: 0; }
.tb-grp .tb-btn { border-radius: 0; border-right-width: 0; }
.tb-grp .tb-btn:first-child { border-radius: 3px 0 0 3px; }
.tb-grp .tb-btn:last-child { border-radius: 0 3px 3px 0; border-right-width: 1px; }
.week-ctr { display: flex; align-items: center; gap: 4px; flex-shrink: 0; }
.week-ctr span {
  font-size: 11px; font-weight: 700; color: var(--navy);
  min-width: 62px; text-align: center;
}
.tb-slider { width: 100px; accent-color: var(--green); cursor: pointer; flex-shrink: 0; }
.yr-chip {
  font-size: 10px; font-family: var(--mono); font-weight: 700;
  padding: 1px 7px; border-radius: 3px; cursor: pointer;
  border: 1px solid transparent; background: transparent;
  transition: all 0.1s; flex-shrink: 0;
}
.yr-chip.on { background: #fff; }

/* ── RANGE CONTROL BAR (Comparativo) ────────────── */
.range-bar {
  display: none;
  background: #f4f4f4;
  border-bottom: 1px solid var(--border);
  padding: 4px 10px;
  align-items: center; gap: 8px;
  height: 30px; overflow: hidden;
}
.range-bar.show { display: flex; }
.range-val {
  font-size: 11px; font-weight: 700; color: var(--navy);
  font-family: var(--mono); min-width: 36px; text-align: center;
}
.range-badge {
  font-size: 10px; font-family: var(--mono);
  background: #e8f5e9; border: 1px solid #a7d7b4;
  color: var(--green); padding: 1px 8px; border-radius: 3px;
  white-space: nowrap; flex-shrink: 0;
}

/* ── VIEW TABS ───────────────────────────────────── */
.view-tabs {
  background: #f8f8f8;
  border-bottom: 2px solid #d5d5d5;
  display: flex; padding: 0; height: 28px;
}
.vtab {
  padding: 0 14px; font-size: 10px; font-weight: 700;
  font-family: var(--mono); cursor: pointer; border: none;
  background: transparent; color: #888;
  border-bottom: 2px solid transparent; margin-bottom: -2px;
  text-transform: uppercase; letter-spacing: 0.5px;
  transition: color 0.1s; white-space: nowrap; height: 28px;
}
.vtab:hover { color: #333; background: rgba(0,0,0,0.03); }
.vtab.active { color: var(--green); border-bottom-color: var(--green); background: #fff; }

/* ── GRID CONTAINER ──────────────────────────────── */
#gridWrap {
  background: #fff;
  border: 1px solid #d5d5d5;
  border-top: none;
}

/* ── STATUS BAR ──────────────────────────────────── */
.statusbar {
  background: #ebebeb; border-top: 1px solid #ccc;
  padding: 2px 10px; font-size: 10px; color: #666;
  display: flex; gap: 14px; align-items: center;
  height: 22px; overflow: hidden;
}
.statusbar b { color: #333; }
.statusbar .st-sep { color: #bbb; }

/* ── PRODUCTOS PANEL ─────────────────────────────── */
#prodPanel {
  display: none; background: #fff;
  border-top: 2px solid var(--green);
}
#prodPanel.show { display: block; }
.prod-hdr {
  background: #1e3a5f; padding: 5px 10px;
  display: flex; align-items: center; gap: 10px; height: 28px;
}
.prod-hdr-title {
  color: #fff; font-size: 11px; font-weight: 700;
  letter-spacing: 0.5px; flex: 1;
}
.prod-hdr-meta { color: rgba(255,255,255,0.6); font-size: 10px; }
.prod-close {
  background: transparent; border: 1px solid rgba(255,255,255,0.3);
  border-radius: 3px; color: rgba(255,255,255,0.8);
  cursor: pointer; font-size: 10px; padding: 1px 8px; font-family: var(--mono);
}
.prod-close:hover { border-color: #fff; color: #fff; }

/* ── AG GRID THEME OVERRIDES ─────────────────────── */
.ag-theme-alpine {
  --ag-font-family: 'Consolas', 'Courier New', monospace;
  --ag-font-size: 11px;
  --ag-row-height: 22px;
  --ag-header-height: 25px;
  --ag-cell-horizontal-padding: 6px;
  --ag-borders: solid 1px;
  --ag-border-color: #d8d8d8;
  --ag-secondary-border-color: #e5e5e5;
  --ag-header-background-color: #e8e8e8;
  --ag-header-foreground-color: #333;
  --ag-odd-row-background-color: #fafafa;
  --ag-even-row-background-color: #ffffff;
  --ag-row-hover-color: #e8f5e9;
  --ag-selected-row-background-color: #c8e6c9;
  --ag-alpine-active-color: #16a34a;
  --ag-input-focus-border-color: #16a34a;
  --ag-range-selection-border-color: #16a34a;
  --ag-header-column-separator-display: block;
  --ag-header-column-separator-height: 60%;
  --ag-header-column-separator-color: #ccc;
}
.ag-theme-alpine .ag-header-cell {
  font-size: 10px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.3px;
}
.ag-theme-alpine .ag-pinned-left-cols-container {
  border-right: 2px solid #aaa !important;
}
.ag-theme-alpine .ag-group-row { background: #eff3fa !important; font-weight: 700; }

/* Inline cell styles injected via cellStyle */
.cell-pos { color: #16a34a !important; font-weight: 600; }
.cell-neg { color: #dc2626 !important; font-weight: 600; }
.cell-muted { color: #999 !important; }
.cell-total { font-weight: 700 !important; color: #1e3a5f !important; }
.prod-link {
  cursor: pointer;
  text-decoration: underline dotted;
  text-underline-offset: 2px;
}

/* ── COMPARATIVO TABLE ───────────────────────── */
#comparativoWrap {
  display: none;
  background: #fff;
  border: 1px solid #d5d5d5;
  border-top: none;
  overflow: hidden;
}
#comparativoWrap.show { display: block; }
.cmp-stat-strip {
  display: flex; gap: 8px; flex-wrap: wrap;
  padding: 8px 10px; background: #f4f4f4;
  border-bottom: 1px solid #d5d5d5;
}
.cmp-stat-box {
  background: #fff; border: 1px solid #ddd; border-radius: 4px;
  padding: 6px 12px; min-width: 130px;
}
.cmp-stat-label { font-size: 9px; text-transform: uppercase; letter-spacing: 0.5px; color: #888; }
.cmp-stat-val   { font-size: 14px; font-weight: 700; margin: 2px 0 1px; }
.cmp-stat-sub   { font-size: 9px; color: #aaa; }
.cmp-tbl-wrap   { overflow-x: auto; -webkit-overflow-scrolling: touch;
                  scrollbar-width: thin; scrollbar-color: #ccc transparent; }
.cmp-tbl-wrap::-webkit-scrollbar { height: 5px; width: 5px; }
.cmp-tbl-wrap::-webkit-scrollbar-thumb { background: #ccc; border-radius: 3px; }
.cmp-tbl {
  border-collapse: collapse; width: 100%;
  font-family: var(--mono); font-size: 11px;
}
.cmp-tbl th {
  padding: 5px 8px; background: #e8e8e8; color: #444;
  font-size: 10px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.3px; white-space: nowrap;
  border-bottom: 2px solid #ccc; border-right: 1px solid #ddd;
  position: sticky; top: 0; z-index: 2; text-align: right;
}
.cmp-tbl th:first-child, .cmp-tbl th:nth-child(2) { text-align: left; }
.cmp-tbl td {
  padding: 4px 8px; border-bottom: 1px solid #eee;
  border-right: 1px solid #f0f0f0; white-space: nowrap;
  text-align: right;
}
.cmp-tbl td:first-child, .cmp-tbl td:nth-child(2) { text-align: left; }
.cmp-grp-hdr td {
  background: #eff3fa; font-weight: 700;
  border-top: 2px solid #ccc; font-size: 11px;
  padding: 5px 8px;
}
.cmp-grp-hdr td:first-child { border-left: 3px solid #16a34a; }
.cmp-row:hover td { background: #f0faf4; }
.cmp-total-row td {
  background: rgba(22,163,74,.06); font-weight: 700;
  border-top: 1px solid rgba(22,163,74,.2); color: #16a34a;
}
.cmp-total-row td:first-child { border-left: 3px solid rgba(22,163,74,.4); }
.delta-cell { font-size: 10px; white-space: nowrap; }
.delta-amt  { display: block; }
.delta-pct  { display: block; font-size: 9px; opacity: 0.8; }
.chg-pos { color: #16a34a; font-weight: 600; }
.chg-neg { color: #dc2626; font-weight: 600; }
.chg-0   { color: #aaa; }
</style>
</head>
<body>

<!-- LOADER -->
<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">CFBC — Cargando datos...</div>
</div>

<!-- APP -->
<div id="app" style="display:none">

  <!-- HEADER -->
  <div class="app-hdr">
    <div class="hdr-brand">CFBC ▸ CONTROL SEMANAL</div>
    <div class="hdr-kpis" id="hdrKpis"></div>
    <button class="hdr-btn" onclick="exportCSV()">⬇ CSV</button>
    <button class="hdr-btn" onclick="recargar()" style="margin-left:4px">⟳</button>
  </div>

  <!-- TOOLBAR -->
  <div class="toolbar">
    <span class="tb-label">Cat</span>
    <select class="tb-sel" id="catSel" onchange="onCatChange(this.value)" style="max-width:200px"></select>
    <div class="tb-sep"></div>
    <div class="tb-grp">
      <button class="tb-btn active" id="btnUSD" onclick="setCurrency('usd')">USD</button>
      <button class="tb-btn" id="btnMXN" onclick="setCurrency('mxn')">MXN</button>
    </div>
    <div class="tb-sep"></div>
    <span class="tb-label">Semana</span>
    <div class="week-ctr">
      <button class="tb-btn" onclick="prevWeek()">◀</button>
      <span id="weekLabel">—</span>
      <button class="tb-btn" onclick="nextWeek()">▶</button>
    </div>
    <input type="range" class="tb-slider" id="weekSlider" min="1" max="52" value="1" oninput="onWeekSlider(this.value)">
    <div class="tb-sep"></div>
    <span class="tb-label">Años</span>
    <div id="yearChips" style="display:flex;gap:3px"></div>
  </div>

  <!-- VIEW TABS -->
  <div class="view-tabs">
    <button class="vtab active" id="vtSemana"    onclick="setView('semana')">Semana</button>
    <button class="vtab"        id="vtAnual"     onclick="setView('anual')">Anual</button>
    <button class="vtab"        id="vtComparativo" onclick="setView('comparativo')">Comparativo</button>
    <button class="vtab"        id="vtRancho"    onclick="setView('rancho')">Por Rancho</button>
    <button class="vtab"        id="vtDetalle"   onclick="setView('detalle')">Detalle Semanal</button>
    <button class="vtab"        id="vtProductos" onclick="setView('productos')">Productos</button>
    <button class="vtab"        id="vtServicios" onclick="setView('servicios')">Costo Servicios</button>
  </div>

  <!-- RANGE CONTROL BAR (solo visible en comparativo) -->
  <div class="range-bar" id="rangeBar">
    <span class="tb-label">Desde</span>
    <span class="range-val" id="fromWeekLabel">W01</span>
    <input type="range" class="tb-slider" id="fromSlider" min="1" max="52" value="1" oninput="onRangeChange()">
    <span style="color:#aaa;font-size:11px">→</span>
    <span class="tb-label">Hasta</span>
    <span class="range-val" id="toWeekLabel">W52</span>
    <input type="range" class="tb-slider" id="toSlider" min="1" max="52" value="52" oninput="onRangeChange()">
    <span class="range-badge" id="rangeBadge">W01 → W52</span>
    <div class="tb-sep"></div>
    <button class="tb-btn" onclick="resetRange()">↺ Reset</button>
  </div>

  <!-- GRID AREA -->
  <div id="gridWrap">
    <div id="myGrid" class="ag-theme-alpine" style="width:100%"></div>
  </div>

  <!-- COMPARATIVO TABLE (reemplaza gridWrap en vista comparativo) -->
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
    <div class="prod-hdr">
      <div class="prod-hdr-title" id="prodTitle">PRODUCTOS</div>
      <div class="prod-hdr-meta" id="prodMeta"></div>
      <button class="prod-close" onclick="closeProdPanel()">✕ CERRAR</button>
    </div>
    <div id="prodGrid" class="ag-theme-alpine" style="width:100%"></div>
  </div>

  <!-- STATUS BAR -->
  <div class="statusbar" id="statusbar">
    <span>Total: <b id="stTotal">—</b></span>
  </div>
</div><!-- /app -->

<script>
// ═══════════════════════════════════════════════════════════
// DATOS
// ═══════════════════════════════════════════════════════════
var _raw = atob('__DATA_JSON__');
var DATA = JSON.parse(_raw);

// ═══════════════════════════════════════════════════════════
// CONSTANTES
// ═══════════════════════════════════════════════════════════
var RANCH_ORDER = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','HOOPS','Cecilia','Cecilia 25','Christina','Albahaca-RM','Campo-VI'];
var RANCH_COLORS = {
  'Prop-RM':'#047857','PosCo-RM':'#1d4ed8','Campo-RM':'#b45309',
  'Isabela':'#7c3aed','HOOPS':'#c2410c','Cecilia':'#be185d',
  'Cecilia 25':'#047857','Christina':'#0369a1','Albahaca-RM':'#6d28d9','Campo-VI':'#64748b'
};
var YEAR_COLORS = {2021:'#0ea5e9',2022:'#f59e0b',2023:'#22c55e',2024:'#a855f7',2025:'#f97316',2026:'#ef4444'};
var CAT_MIRFE = 'FERTILIZANTES';
var CAT_MIPE  = 'DESINFECCION / PLAGUICIDAS';
// CAT_MIRFE y CAT_MIPE son categorías independientes — nunca se combinan

// ═══════════════════════════════════════════════════════════
// ESTADO
// ═══════════════════════════════════════════════════════════
var state = {
  cat: '', currency: 'usd', activeYears: {}, view: 'semana',
  weekIdx: 0, fromWeek: 1, toWeek: 52
};
var allWeeks = [];
var mainGridApi = null;
var prodGridApi = null;

function isCombined(cat) { return false; } // Cada categoría se muestra por separado

// ═══════════════════════════════════════════════════════════
// FORMATEO
// ═══════════════════════════════════════════════════════════
function fmt(n) {
  if (n === null || n === undefined || n === 0 || isNaN(n)) return '—';
  var neg = n < 0, s = Math.abs(n);
  return (neg ? '-$' : '$') + s.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
function fmtFull(n) {
  if (!n || isNaN(n)) return '—';
  var neg = n < 0, s = Math.abs(n);
  return (neg ? '-$' : '$') + s.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  var sign = n > 0 ? '+' : '';
  return sign + n.toFixed(1) + '%';
}
function wFmt(n) { return 'W' + String(n).padStart(2,'0'); }
function recargar() { window.location.reload(); }

// ═══════════════════════════════════════════════════════════
// DATA HELPERS
// ═══════════════════════════════════════════════════════════
function getActiveYears() {
  return DATA.years.filter(function(y) { return state.activeYears[y]; });
}
function getWeekDetail(cat, weekNum, yr) {
  return DATA.weekly_detail.filter(function(r) {
    return r.categoria === cat && r.week === weekNum && r.year === yr;
  });
}
function ranchFieldName(ranch) {
  return 'r_' + ranch.replace(/[^a-zA-Z0-9]/g,'_');
}
function fieldToRanch(fieldName) {
  if (!fieldName) return null;
  for (var i = 0; i < RANCH_ORDER.length; i++) {
    var rn = RANCH_ORDER[i];
    if (ranchFieldName(rn) === fieldName) return rn;
  }
  return null;
}
function monthFromRecord(rec) {
  var dr = String(rec.date_range || '').toLowerCase();
  var m = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'setiembre':9,'octubre':10,'noviembre':11,'diciembre':12,
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12
  };
  for (var k in m) {
    if (dr.indexOf(k) !== -1) return m[k];
  }
  var wk = parseInt(rec.week || 1);
  if (!wk || wk < 1) wk = 1;
  return Math.max(1, Math.min(12, Math.ceil(wk / 4.35)));
}
function monthLabel(m) {
  var names = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  return names[(m || 1) - 1] || ('M' + String(m || 1).padStart(2,'0'));
}
function sumDetail(recs, currency) {
  var out = { total: 0, ranches: {} };
  recs.forEach(function(r) {
    var v = currency === 'usd' ? r.usd_total : r.mxn_total;
    out.total += v;
    var rsrc = currency === 'usd' ? r.usd_ranches : r.mxn_ranches;
    Object.keys(rsrc || {}).forEach(function(rn) {
      out.ranches[rn] = (out.ranches[rn] || 0) + rsrc[rn];
    });
    if (r.date_range) out.date_range = r.date_range;
  });
  return out;
}

// ═══════════════════════════════════════════════════════════
// INICIALIZAR
// ═══════════════════════════════════════════════════════════
function inicializar() {
  if (!window._prodLinkBound) {
    document.addEventListener('click', function(e) {
      var el = e.target.closest('.prod-link');
      if (!el) return;
      var row = {
        _cat: decodeURIComponent(el.dataset.cat || ''),
        _year: parseInt(el.dataset.year || '0', 10),
        _week: parseInt(el.dataset.week || '0', 10),
        _fromWeek: parseInt(el.dataset.from || el.dataset.week || '0', 10),
        _toWeek: parseInt(el.dataset.to || el.dataset.week || '0', 10),
      };
      var ranch = decodeURIComponent(el.dataset.ranch || '');
      showProdPanel(row, { ranch: ranch || null });
      e.stopPropagation();
      e.preventDefault();
    });
    window._prodLinkBound = true;
  }

  // Estado inicial
  var prefCat = 'MATERIAL DE EMPAQUE';
  state.cat = DATA.categories.indexOf(prefCat) > -1 ? prefCat : DATA.categories[0];

  // Año más reciente activo
  state.activeYears = {};
  var latestYr = DATA.years[DATA.years.length - 1];
  var prevYr = DATA.years[DATA.years.length - 2];
  if (latestYr) state.activeYears[latestYr] = true;
  if (prevYr)   state.activeYears[prevYr]   = true;

  // Semanas disponibles
  var wSet = {};
  DATA.weekly_detail.forEach(function(r) { wSet[r.week] = 1; });
  allWeeks = Object.keys(wSet).map(Number).sort(function(a,b) { return a-b; });

  // Ir a la semana más reciente del año más reciente
  var wksLatest = DATA.weekly_detail
    .filter(function(r) { return r.year === latestYr; })
    .map(function(r) { return r.week; })
    .filter(function(v,i,a) { return a.indexOf(v) === i; })
    .sort(function(a,b) { return a-b; });
  var curWeek = wksLatest[wksLatest.length - 1] || allWeeks[allWeeks.length - 1];
  var idx = allWeeks.indexOf(curWeek);
  state.weekIdx = idx >= 0 ? idx : allWeeks.length - 1;

  // Rango inicial: últimas 2 semanas del año más reciente
  state.toWeek   = wksLatest[wksLatest.length - 1] || allWeeks[allWeeks.length - 1] || 52;
  state.fromWeek = wksLatest[wksLatest.length - 2] || wksLatest[0] || state.toWeek;

  buildCatSelect();
  buildYearChips();
  updateWeekControls();
  updateRangeSliders();
  buildMainGrid();
  renderView();
  updateHeader();
  document.getElementById('loader').style.display = 'none';
  document.getElementById('app').style.display   = 'block';
  setTimeout(function(){ if (mainGridApi) mainGridApi.sizeColumnsToFit(); }, 100);
}

// ═══════════════════════════════════════════════════════════
// UI BUILDERS
// ═══════════════════════════════════════════════════════════
function buildCatSelect() {
  var el = document.getElementById('catSel');
  el.innerHTML = DATA.categories.map(function(c) {
    return '<option value="' + c.replace(/"/g,'&quot;') + '"' + (c === state.cat ? ' selected' : '') + '>' + c + '</option>';
  }).join('');
}
function buildYearChips() {
  var el = document.getElementById('yearChips');
  el.innerHTML = DATA.years.map(function(y) {
    var col = YEAR_COLORS[y] || '#888';
    var on = state.activeYears[y] ? ' on' : '';
    return '<button class="yr-chip' + on + '" id="yrChip' + y + '" style="color:' + col + ';border-color:' + (state.activeYears[y] ? col : 'transparent') + ';background:' + (state.activeYears[y] ? col + '20' : 'transparent') + '" onclick="toggleYear(' + y + ')">' + y + '</button>';
  }).join('');
}
function updateWeekControls() {
  var wn = allWeeks[state.weekIdx] || 1;
  var sl = document.getElementById('weekSlider');
  sl.min = allWeeks[0] || 1; sl.max = allWeeks[allWeeks.length-1] || 52; sl.value = wn;
  var activeYrs = getActiveYears();
  var yr = activeYrs[activeYrs.length - 1] || DATA.years[DATA.years.length - 1];
  var dateText = '';
  if (DATA.week_date_ranges) dateText = DATA.week_date_ranges[yr + '-' + wn] || DATA.week_date_ranges[String(yr) + '-' + String(wn)] || '';
  if (!dateText) {
    var recs = (DATA.weekly_detail || []).filter(function(r) { return r.week === wn && r.year === yr && r.date_range; });
    if (recs.length) dateText = recs[0].date_range;
  }
  document.getElementById('weekLabel').textContent = String(yr).slice(2) + String(wn).padStart(2, '0');
}
function updateHeader() {
  var yrs = getActiveYears();
  var wn  = allWeeks[state.weekIdx] || 1;
  var curYr  = yrs[yrs.length - 1];
  var prevYr = yrs[yrs.length - 2];

  var grandTotal = 0;
  getWeekDetail(state.cat, wn, curYr).forEach(function(r) {
    grandTotal += state.currency === 'usd' ? r.usd_total : r.mxn_total;
  });

  var prevTotal = 0;
  if (prevYr) {
    getWeekDetail(state.cat, wn, prevYr).forEach(function(r) {
      prevTotal += state.currency === 'usd' ? r.usd_total : r.mxn_total;
    });
  }
  var delta = prevTotal > 0 ? (grandTotal - prevTotal) / prevTotal * 100 : null;

  var annualTotal = 0;
  var d = (DATA.summary[state.cat] || {})[curYr];
  if (d) annualTotal = state.currency === 'usd' ? d.usd : d.mxn;

  document.getElementById('hdrKpis').innerHTML = '';
}

// ═══════════════════════════════════════════════════════════
// EVENTS
// ═══════════════════════════════════════════════════════════
function onCatChange(val) {
  state.cat = val;
  renderView();
  updateHeader();
}
function setCurrency(cur) {
  state.currency = cur;
  document.getElementById('btnUSD').className = 'tb-btn' + (cur === 'usd' ? ' active' : '');
  document.getElementById('btnMXN').className = 'tb-btn' + (cur === 'mxn' ? ' active' : '');
  renderView();
  updateHeader();
}
function toggleYear(y) {
  var active = DATA.years.filter(function(yr) { return state.activeYears[yr]; });
  if (state.activeYears[y] && active.length > 1) delete state.activeYears[y];
  else state.activeYears[y] = true;
  buildYearChips();
  renderView();
  updateHeader();
}
function prevWeek() {
  if (state.weekIdx > 0) { state.weekIdx--; updateWeekControls(); renderView(); updateHeader(); }
}
function nextWeek() {
  if (state.weekIdx < allWeeks.length - 1) { state.weekIdx++; updateWeekControls(); renderView(); updateHeader(); }
}
function onWeekSlider(val) {
  var wn = parseInt(val);
  var idx = allWeeks.indexOf(wn);
  if (idx < 0) {
    idx = 0; var mn = Math.abs(allWeeks[0] - wn);
    allWeeks.forEach(function(w,i) { var d=Math.abs(w-wn); if(d<mn){mn=d;idx=i;} });
  }
  state.weekIdx = idx;
  updateWeekControls(); renderView(); updateHeader();
}
function setView(v) {
  state.view = v;
  ['semana','anual','comparativo','rancho','detalle','productos','servicios'].forEach(function(name) {
    var el = document.getElementById('vt' + name.charAt(0).toUpperCase() + name.slice(1));
    if (el) el.className = 'vtab' + (v === name ? ' active' : '');
  });
  // Mostrar/ocultar barra de rango solo en comparativo
  var rb = document.getElementById('rangeBar');
  if (rb) rb.className = 'range-bar' + (v === 'comparativo' ? ' show' : '');
  // Alternar entre AG Grid y tabla comparativo
  var gw  = document.getElementById('gridWrap');
  var cmp = document.getElementById('comparativoWrap');
  if (v === 'comparativo') {
    if (gw)  gw.style.display  = 'none';
    if (cmp) cmp.className = 'show';
  } else {
    if (gw)  gw.style.display  = '';
    if (cmp) cmp.className = '';
  }
  closeProdPanel();
  renderView();
}
function exportCSV() {
  if (mainGridApi) mainGridApi.exportDataAsCsv({ fileName: 'CFBC_' + state.view + '_' + new Date().toISOString().slice(0,10) + '.csv' });
}
function updateRangeSliders() {
  var f = state.fromWeek, t = state.toWeek;
  var fEl = document.getElementById('fromSlider');
  var tEl = document.getElementById('toSlider');
  var min = allWeeks[0] || 1, max = allWeeks[allWeeks.length-1] || 52;
  if (fEl) { fEl.min = min; fEl.max = max; fEl.value = f; }
  if (tEl) { tEl.min = min; tEl.max = max; tEl.value = t; }
  var fLbl = document.getElementById('fromWeekLabel');
  var tLbl = document.getElementById('toWeekLabel');
  var badge = document.getElementById('rangeBadge');
  var latestYr = DATA.years[DATA.years.length - 1];
  var yy = String(latestYr).slice(2);
  if (fLbl) fLbl.textContent = yy + String(f).padStart(2, '0');
  if (tLbl) tLbl.textContent = yy + String(t).padStart(2, '0');
  var count = allWeeks.filter(function(w){ return w >= f && w <= t; }).length;
  if (badge) badge.textContent = yy + String(f).padStart(2,'0') + ' → ' + yy + String(t).padStart(2,'0') + ' · ' + count + ' sem';
}
function onRangeChange() {
  var f = parseInt(document.getElementById('fromSlider').value);
  var t = parseInt(document.getElementById('toSlider').value);
  if (f > t) { var tmp = f; f = t; t = tmp; }
  state.fromWeek = f; state.toWeek = t;
  updateRangeSliders();
  if (state.view === 'comparativo') renderComparativo();
}
function resetRange() {
  var latestYr = DATA.years[DATA.years.length - 1];
  var wks = DATA.weekly_detail
    .filter(function(r){ return r.year === latestYr; })
    .map(function(r){ return r.week; })
    .filter(function(v,i,a){ return a.indexOf(v) === i; })
    .sort(function(a,b){ return a - b; });
  state.toWeek   = wks[wks.length - 1] || allWeeks[allWeeks.length - 1] || 52;
  state.fromWeek = wks[wks.length - 2] || wks[0] || state.toWeek;
  updateRangeSliders();
  if (state.view === 'comparativo') renderComparativo();
}

// ═══════════════════════════════════════════════════════════
// MAIN GRID SETUP
// ═══════════════════════════════════════════════════════════
function buildMainGrid() {
  var el = document.getElementById('myGrid');
  var opts = {
    columnDefs: [],
    rowData: [],
    rowHeight: 22,
    headerHeight: 25,
    defaultColDef: {
      sortable: true,
      filter: true,
      resizable: true,
      suppressMovable: false,
    },
    suppressCellFocus: false,
    enableCellTextSelection: true,
    animateRows: false,
    suppressColumnVirtualisation: false,
    domLayout: 'autoHeight',
    onGridReady: function(params) { mainGridApi = params.api; },
    onCellClicked: function(e) { onMainCellClick(e); },
  };
  new agGrid.Grid(el, opts);
}
function setMainGrid(colDefs, rowData, pinnedBottom, statusText) {
  if (!mainGridApi) return;
  mainGridApi.setPinnedBottomRowData([]);
  mainGridApi.setColumnDefs(colDefs);
  mainGridApi.setRowData(rowData);
  mainGridApi.sizeColumnsToFit();
  document.getElementById('stTotal').textContent = statusText || '';
  setTimeout(reportHeight, 80);
}

// ═══════════════════════════════════════════════════════════
// CELL RENDERERS
// ═══════════════════════════════════════════════════════════
function moneyRenderer(params) {
  var v = params.value;
  if (v === null || v === undefined || v === 0 || isNaN(v)) return '<span style="color:#bbb">—</span>';
  return '<span style="color:#1e3a5f;font-weight:600">' + fmt(v) + '</span>';
}
function deltaRenderer(params) {
  var v = params.value;
  if (v === null || v === undefined || isNaN(v)) return '<span style="color:#bbb">—</span>';
  if (Math.abs(v) < 0.5) return '<span style="color:#999">~0%</span>';
  var col = v > 0 ? '#16a34a' : '#dc2626';
  var arrow = v > 0 ? '▲' : '▼';
  return '<span style="color:' + col + ';font-weight:700">' + arrow + ' ' + Math.abs(v).toFixed(1) + '%</span>';
}
function deltaAmtRenderer(params) {
  var v = params.value;
  if (!v || isNaN(v)) return '<span style="color:#bbb">—</span>';
  var col = v > 0 ? '#16a34a' : '#dc2626';
  var sign = v > 0 ? '+' : '';
  return '<span style="color:' + col + '">' + sign + fmt(v) + '</span>';
}
function barRenderer(maxVal) {
  return function(params) {
    var v = params.value;
    if (!v || isNaN(v)) return '<span style="color:#bbb">—</span>';
    var pct = Math.min(v / (maxVal || 1) * 54, 54);
    var color = RANCH_COLORS[params.colDef.field] || '#16a34a';
    return '<div style="display:flex;align-items:center;gap:4px">' +
      '<div style="width:' + pct.toFixed(0) + 'px;height:7px;background:' + color + ';border-radius:1px;flex-shrink:0"></div>' +
      '<span style="color:#333">' + fmt(v) + '</span></div>';
  };
}
function catRenderer(params) {
  var v = params.value;
  if (!v) return '';
  return '<span style="font-weight:700;color:#1e3a5f;font-size:10px">' + v + '</span>';
}
function ranchRenderer(ranch) {
  var col = RANCH_COLORS[ranch] || '#888';
  return function(params) {
    var v = params.value;
    if (!v || isNaN(v) || v === 0) return '<span style="color:#ddd">—</span>';
    // mini bar proportional
    var maxV = params.colDef._maxVal || 1;
    var w = Math.min(v / maxV * 40, 40);
    return '<div style="display:flex;align-items:center;gap:3px">' +
      '<div style="width:' + w.toFixed(0) + 'px;height:6px;background:' + col + ';border-radius:1px;flex-shrink:0;min-width:2px"></div>' +
      '<span style="color:' + col + ';font-weight:600">' + fmt(v) + '</span></div>';
  };
}

// ═══════════════════════════════════════════════════════════
// VIEW ROUTER
// ═══════════════════════════════════════════════════════════
function renderView() {
  if (!mainGridApi) return;
  document.getElementById('prodPanel').className = '';
  if      (state.view === 'semana')    renderSemana();
  else if (state.view === 'anual')     renderAnual();
  else if (state.view === 'comparativo') renderComparativo();
  else if (state.view === 'rancho')    renderRancho();
  else if (state.view === 'detalle')   renderDetalle();
  else if (state.view === 'productos') renderProductosFull();
  else if (state.view === 'servicios') renderServicios();
  setTimeout(function(){ if (mainGridApi) mainGridApi.sizeColumnsToFit(); }, 50);
}

// ═══════════════════════════════════════════════════════════
// VIEW 1: SEMANA
// Rows = years (or years x MIRFE/MIPE), Cols = [Year, Week, Cat, Total, Delta, ranches]
// ═══════════════════════════════════════════════════════════
function renderSemana() {
  var yrs = getActiveYears();
  var wn  = allWeeks[state.weekIdx] || 1;
  var sym = state.currency.toUpperCase();

  var cols = [
    { field: 'year', headerName: 'AÑO', pinned: 'left', width: 70, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: function(p) { return '<span style="color:'+(YEAR_COLORS[p.value]||'#888')+';font-weight:700">'+p.value+'</span>'; } },
    { field: 'week', headerName: 'SEM', width: 60, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: function(p){ return wFmt(p.value); } },
    { field: 'cat_label', headerName: 'CATEGORÍA', width: 170, filter: 'agTextColumnFilter', cellRenderer: catRenderer },
    { field: 'total', headerName: 'TOTAL ' + sym, width: 110, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: moneyRenderer },
    { field: 'deltaAmt', headerName: 'Δ $', width: 90, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: deltaAmtRenderer },
    { field: 'deltaPct', headerName: 'Δ %', width: 72, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: deltaRenderer },
  ];
  RANCH_ORDER.forEach(function(r) {
    cols.push({
      field: 'r_' + r.replace(/[^a-zA-Z0-9]/g,'_'),
      headerName: r, width: 100, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: ranchRenderer(r)
    });
  });

  var rows = [];
  var grandTotal = 0;

  yrs.forEach(function(yr, i) {
    var prevYr = i > 0 ? yrs[i-1] : null;
    var recs = getWeekDetail(state.cat, wn, yr);
    var agg = sumDetail(recs, state.currency);
    var row = { year: yr, week: wn, cat_label: state.cat, _cat: state.cat, _year: yr, _week: wn };
    row.total = agg.total;
    if (prevYr) {
      var recsP = getWeekDetail(state.cat, wn, prevYr);
      var aggP = sumDetail(recsP, state.currency);
      row.deltaAmt = agg.total - aggP.total;
      row.deltaPct = aggP.total > 0 ? (agg.total - aggP.total) / aggP.total * 100 : null;
    }
    RANCH_ORDER.forEach(function(r) { row['r_' + r.replace(/[^a-zA-Z0-9]/g,'_')] = agg.ranches[r] || 0; });
    rows.push(row);
    if (yr === yrs[yrs.length-1]) grandTotal += agg.total;
  });

  setMainGrid(cols, rows, [], fmt(grandTotal) + ' ' + sym + ' · AÑO ' + yrs[yrs.length-1]);
}

// ═══════════════════════════════════════════════════════════
// VIEW 2: ANUAL
// Rows = years, Cols = [Year, Cat, Total, Delta, ranches...]
// ═══════════════════════════════════════════════════════════
function renderAnual() {
  var yrs = getActiveYears();
  var sym = state.currency.toUpperCase();

  var cols = [
    { field: 'year', headerName: 'AÑO', pinned: 'left', width: 70, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: function(p) { return '<span style="color:'+(YEAR_COLORS[p.value]||'#888')+';font-weight:700">'+p.value+'</span>'; } },
    { field: 'cat_label', headerName: 'CATEGORÍA', width: 170, filter: 'agTextColumnFilter', cellRenderer: catRenderer },
    { field: 'total', headerName: 'TOTAL ' + sym, width: 110, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: moneyRenderer },
    { field: 'deltaAmt', headerName: 'Δ $', width: 90, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: deltaAmtRenderer },
    { field: 'deltaPct', headerName: 'Δ %', width: 72, type: 'numericColumn', filter: 'agNumberColumnFilter', cellRenderer: deltaRenderer },
  ];
  RANCH_ORDER.forEach(function(r) {
    cols.push({
      field: 'r_' + r.replace(/[^a-zA-Z0-9]/g,'_'),
      headerName: r, width: 100, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: ranchRenderer(r)
    });
  });

  var rows = [];
  var grandTotal = 0;

  var getYrAgg = function(cat, yr) {
    var d = (DATA.summary[cat] || {})[yr] || {usd:0, mxn:0, ranches:{}, ranches_mxn:{}};
    return {
      total: state.currency === 'usd' ? d.usd : d.mxn,
      ranches: state.currency === 'usd' ? d.ranches : d.ranches_mxn
    };
  };

  yrs.forEach(function(yr, i) {
    var prevYr = i > 0 ? yrs[i-1] : null;
    var agg = getYrAgg(state.cat, yr);
    var row = { year: yr, cat_label: state.cat, _cat: state.cat, _year: yr };
    row.total = agg.total;
    if (prevYr) {
      var aggP = getYrAgg(state.cat, prevYr);
      row.deltaAmt = agg.total - aggP.total;
      row.deltaPct = aggP.total > 0 ? (agg.total - aggP.total) / aggP.total * 100 : null;
    }
    RANCH_ORDER.forEach(function(r) { row['r_' + r.replace(/[^a-zA-Z0-9]/g,'_')] = agg.ranches[r] || 0; });
    rows.push(row);
    if (yr === yrs[yrs.length-1]) grandTotal += agg.total;
  });

  setMainGrid(cols, rows, [], fmt(grandTotal) + ' ' + sym + ' · AÑO ' + yrs[yrs.length-1]);
}

// ═══════════════════════════════════════════════════════════
// VIEW 3: COMPARATIVO (tabla agrupada, equivale al Tendencia original)
// ═══════════════════════════════════════════════════════════
var rangeTableGroup = 'year'; // 'year' = Año→Semana | 'week' = Semana→Año

function setRangeTableGroup(g) {
  rangeTableGroup = g;
  document.getElementById('rtgYear').className = 'tb-btn' + (g === 'year' ? ' active' : '');
  document.getElementById('rtgWeek').className = 'tb-btn' + (g === 'week' ? ' active' : '');
  renderComparativo();
}

// Extrae solo "Mes Año" de strings como "Del 02 al 08 de Marzo 2026"
function fmtMes(dr) {
  if (!dr) return '—';
  var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var lower = dr.toLowerCase();
  for (var i = 0; i < MESES.length; i++) {
    if (lower.indexOf(MESES[i]) > -1) {
      var m = MESES[i].charAt(0).toUpperCase() + MESES[i].slice(1);
      var yrMatch = dr.match(/\b(20\d{2})\b/);
      return m + (yrMatch ? ' ' + yrMatch[1] : '');
    }
  }
  return dr;
}

// Agrega todos los registros de una lista en un objeto {usd,mxn,ranches,ranches_mxn,date_range}
function aggregateRecs(recs) {
  var out = { usd: 0, mxn: 0, ranches: {}, ranches_mxn: {}, date_range: '' };
  recs.forEach(function(r) {
    out.usd += r.usd_total; out.mxn += r.mxn_total;
    if (r.date_range) out.date_range = r.date_range;
    Object.keys(r.usd_ranches || {}).forEach(function(rn) { out.ranches[rn] = (out.ranches[rn] || 0) + r.usd_ranches[rn]; });
    Object.keys(r.mxn_ranches || {}).forEach(function(rn) { out.ranches_mxn[rn] = (out.ranches_mxn[rn] || 0) + r.mxn_ranches[rn]; });
  });
  out.usd = Math.round(out.usd * 100) / 100;
  out.mxn = Math.round(out.mxn * 100) / 100;
  return out;
}

// Retorna {yr: {usd,mxn,ranches,ranches_mxn,weekly:{wk:val}}} para el rango
function getRangeByYear(cat, fromW, toW) {
  var res = {};
  getActiveYears().forEach(function(yr) {
    var recs = DATA.weekly_detail.filter(function(r) {
      return r.categoria === cat && r.year === yr && r.week >= fromW && r.week <= toW;
    });
    if (!recs.length) return;
    var ag = aggregateRecs(recs);
    ag.weekly = {};
    recs.forEach(function(r) {
      ag.weekly[r.week] = (ag.weekly[r.week] || 0) + (state.currency === 'usd' ? r.usd_total : r.mxn_total);
    });
    res[yr] = ag;
  });
  return res;
}

// Celda de delta: valor actual vs anterior
function deltaCellHtml(val, prev) {
  if (prev === null || prev === undefined || prev === 0) return '<td class="delta-cell chg-0">—</td>';
  var diff = val - prev;
  var p = ((diff / prev) * 100).toFixed(1);
  var cls = diff > 0 ? 'chg-pos' : diff < 0 ? 'chg-neg' : 'chg-0';
  var sign = diff > 0 ? '+' : '';
  return '<td class="delta-cell ' + cls + '"><span class="delta-amt">' + sign + fmt(diff) + '</span>' +
         '<span class="delta-pct">' + sign + p + '%</span></td>';
}

function renderComparativo() {
  var f    = state.fromWeek, t = state.toWeek;
  var yrs  = getActiveYears();
  var sym  = state.currency.toUpperCase();
  var byYear = getRangeByYear(state.cat, f, t);

  var rangeWeeks = allWeeks.filter(function(w) { return w >= f && w <= t; });
  var ranchCols  = RANCH_ORDER;

  // ── Stat strip eliminado ────────────────────────────────
  document.getElementById('cmpStats').innerHTML = '';

  // ── Precargar weekData ──────────────────────────────────
  var weekData = {};
  yrs.forEach(function(yr) {
    weekData[yr] = {};
    rangeWeeks.forEach(function(w) {
      var recs = DATA.weekly_detail.filter(function(r) {
        return r.categoria === state.cat && r.year === yr && r.week === w;
      });
      if (recs.length) weekData[yr][w] = aggregateRecs(recs);
    });
  });

  var head, body;

  if (rangeTableGroup === 'year') {
    // ── MODO: Año → Semana ─────────────────────────────────
    // Cabecera
    head = '<tr><th>Semana</th><th>Fecha</th><th>Total ' + sym + '</th><th>Δ$ vs sem ant.</th>' +
      ranchCols.map(function(r) { return '<th>' + r + '</th>'; }).join('') + '</tr>';

    body = yrs.map(function(yr, yi) {
      var col = YEAR_COLORS[yr] || '#888';
      var yearTotal = byYear[yr] ? (state.currency === 'usd' ? byYear[yr].usd : byYear[yr].mxn) : 0;
      var prevYrD = yi > 0 ? byYear[yrs[yi - 1]] : null;
      var prevYrVal = prevYrD ? (state.currency === 'usd' ? prevYrD.usd : prevYrD.mxn) : null;
      var yDiff = prevYrVal !== null ? yearTotal - prevYrVal : null;
      var yPct  = (prevYrVal !== null && prevYrVal !== 0) ? ((yearTotal - prevYrVal) / prevYrVal * 100).toFixed(1) : null;
      var yCls  = yDiff === null ? 'chg-0' : yDiff > 0 ? 'chg-pos' : 'chg-neg';
      var ySign = yDiff !== null && yDiff > 0 ? '+' : '';

      // Fila de cabecera del año con totales y deltas por rancho
      var ranchHdrCells = ranchCols.map(function(r) {
        var d = byYear[yr]; if (!d) return '<td>—</td>';
        var src = state.currency === 'usd' ? d.ranches : d.ranches_mxn;
        var v = src[r] || 0;
        return '<td style="color:' + (v > 0 ? (RANCH_COLORS[r] || '#888') : '#bbb') + ';font-size:10px">' + (v > 0 ? fmt(v) : '—') + '</td>';
      }).join('');

      var hdr = '';

      // Filas de semanas dentro del año
      var prevWkVal = null;
      var wkRows = rangeWeeks.map(function(w) {
        var d   = weekData[yr][w];
        var val = d ? (state.currency === 'usd' ? d.usd : d.mxn) : 0;
        var dCell = deltaCellHtml(val, prevWkVal);
        if (val > 0) prevWkVal = val;
        var ranchCells = ranchCols.map(function(r) {
          if (!d) return '<td style="color:#ddd">—</td>';
          var src = state.currency === 'usd' ? d.ranches : d.ranches_mxn;
          var v = src[r] || 0;
          var style = 'color:' + (v > 0 ? (RANCH_COLORS[r] || '#888') : '#ddd') + (v > 0 ? ';cursor:pointer' : '');
          var attrs = v > 0 ? ' class="cmp-clickable" data-yr="' + yr + '" data-wk="' + w + '" data-ranch="' + r + '"' : '';
          return '<td style="' + style + '"' + attrs + '>' + (v > 0 ? fmt(v) : '—') + '</td>';
        }).join('');
        var totalStyle = 'color:' + (val > 0 ? col : '#bbb') + ';font-weight:' + (val > 0 ? '600' : '400') + (val > 0 ? ';cursor:pointer' : '');
        var totalAttrs = val > 0 ? ' class="cmp-clickable" data-yr="' + yr + '" data-wk="' + w + '" data-ranch=""' : '';
        return '<tr class="cmp-row">' +
          '<td style="color:' + col + ';font-weight:600">' + String(yr).slice(2) + String(w).padStart(2,'0') + '</td>' +
          '<td style="color:#999;font-size:10px">' + fmtMes(d && d.date_range) + '</td>' +
          '<td style="' + totalStyle + '"' + totalAttrs + '>' + fmt(val) + '</td>' +
          dCell + ranchCells + '</tr>';
      }).join('');

      return hdr + wkRows;
    }).join('');

  } else {
    // ── MODO: Semana → Año ─────────────────────────────────
    head = '<tr><th>Año</th><th>Total ' + sym + '</th><th>Δ$ vs año ant.</th>' +
      ranchCols.map(function(r) { return '<th>' + r + '</th>'; }).join('') + '</tr>';

    body = rangeWeeks.map(function(w) {
      // Buscar fecha de referencia para esta semana
      var dateEx = '';
      yrs.forEach(function(yr) { if (weekData[yr][w] && weekData[yr][w].date_range) dateEx = weekData[yr][w].date_range; });

      var hdr = '<tr class="cmp-grp-hdr"><td colspan="2" style="color:var(--green)">📆 ' + wFmt(w) +
        (dateEx ? ' <span style="font-size:9px;color:#999;font-weight:400">' + fmtMes(dateEx) + '</span>' : '') +
        '</td><td colspan="' + (1 + ranchCols.length) + '"></td></tr>';

      var prevYrVal = null;
      var yrRows = yrs.map(function(yr) {
        var col = YEAR_COLORS[yr] || '#888';
        var d   = weekData[yr][w];
        var val = d ? (state.currency === 'usd' ? d.usd : d.mxn) : 0;
        var dCell = deltaCellHtml(val, prevYrVal);
        if (val > 0) prevYrVal = val;
        var ranchCells = ranchCols.map(function(r) {
          if (!d) return '<td style="color:#ddd">—</td>';
          var src = state.currency === 'usd' ? d.ranches : d.ranches_mxn;
          var v = src[r] || 0;
          var style = 'color:' + (v > 0 ? (RANCH_COLORS[r] || '#888') : '#ddd') + (v > 0 ? ';cursor:pointer' : '');
          var attrs = v > 0 ? ' class="cmp-clickable" data-yr="' + yr + '" data-wk="' + w + '" data-ranch="' + r + '"' : '';
          return '<td style="' + style + '"' + attrs + '>' + (v > 0 ? fmt(v) : '—') + '</td>';
        }).join('');
        var totalStyle2 = 'color:' + (val > 0 ? col : '#bbb') + ';font-weight:' + (val > 0 ? '600' : '400') + (val > 0 ? ';cursor:pointer' : '');
        var totalAttrs2 = val > 0 ? ' class="cmp-clickable" data-yr="' + yr + '" data-wk="' + w + '" data-ranch=""' : '';
        return '<tr class="cmp-row">' +
          '<td><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:' + col + ';margin-right:5px"></span>' +
          '<strong style="color:' + col + '">' + yr + '</strong></td>' +
          '<td style="' + totalStyle2 + '"' + totalAttrs2 + '>' + fmt(val) + '</td>' +
          dCell + ranchCells + '</tr>';
      }).join('');

      // Fila de total de la semana (suma de todos los años)
      var wkTotal = yrs.reduce(function(acc, yr) {
        var d = weekData[yr][w];
        return acc + (d ? (state.currency === 'usd' ? d.usd : d.mxn) : 0);
      }, 0);
      var totalRow = '<tr class="cmp-total-row"><td>TOTAL</td><td>' + fmt(wkTotal) +
        '</td><td colspan="' + (1 + ranchCols.length) + '"></td></tr>';

      return hdr + yrRows + totalRow;
    }).join('');
  }

  document.getElementById('cmpHead').innerHTML = head;
  document.getElementById('cmpBody').innerHTML = body;

  // Status bar
  var grandTotal = yrs.reduce(function(s, yr) {
    var d = byYear[yr]; return s + (d ? (state.currency === 'usd' ? d.usd : d.mxn) : 0);
  }, 0);
  document.getElementById('stTotal').textContent = fmt(grandTotal) + ' ' + sym;
}

// ═══════════════════════════════════════════════════════════
// VIEW 4: POR RANCHO
// Rows = ranches, Cols = [rancho, prevYr, curYr, Δ$, Δ%] 
// (For the selected category)
// ═══════════════════════════════════════════════════════════
function renderRancho() {
  var yrs  = getActiveYears();
  var wn   = allWeeks[state.weekIdx] || 1;
  var cur  = yrs[yrs.length - 1];
  var prev = yrs.length > 1 ? yrs[yrs.length - 2] : null;
  var sym  = state.currency.toUpperCase();

  var cols = [
    { field: 'rancho', headerName: 'RANCHO', pinned: 'left', width: 140,
      cellRenderer: function(p) {
        var c = RANCH_COLORS[p.value] || '#888';
        return '<span style="color:' + c + ';font-weight:700">' + (p.value || '') + '</span>';
      }, filter: 'agTextColumnFilter' }
  ];
  if (prev) cols.push({ field: 'v' + prev, headerName: String(prev) + ' ' + sym, width: 120, type: 'numericColumn', cellRenderer: moneyRenderer });
  cols.push({ field: 'v' + cur, headerName: String(cur) + ' ' + sym + ' ★', width: 120, type: 'numericColumn', cellRenderer: moneyRenderer });
  if (prev) {
    cols.push({ field: 'deltaAmt', headerName: 'Δ $', width: 100, type: 'numericColumn', cellRenderer: deltaAmtRenderer });
    cols.push({ field: 'deltaPct', headerName: 'Δ %', width: 90, type: 'numericColumn', cellRenderer: deltaRenderer });
  }

  var grandCur = 0, grandPrev = 0;

  var rows = RANCH_ORDER.map(function(ranch) {
    var row = { rancho: ranch, _cat: state.cat, _week: wn, _year: cur, _fromWeek: wn, _toWeek: wn };
    var totalCur = 0, totalPrev = 0;

    var aC = sumDetail(getWeekDetail(state.cat, wn, cur), state.currency);
    totalCur = aC.ranches[ranch] || 0;
    if (prev) {
      var aP = sumDetail(getWeekDetail(state.cat, wn, prev), state.currency);
      totalPrev = aP.ranches[ranch] || 0;
    }

    row['v' + cur] = totalCur; grandCur += totalCur;
    if (prev) {
      row['v' + prev] = totalPrev; grandPrev += totalPrev;
      row.deltaAmt = totalCur - totalPrev;
      row.deltaPct = totalPrev > 0 ? (totalCur - totalPrev) / totalPrev * 100 : null;
    }
    return row;
  }).filter(function(r) { return (r['v' + cur] || 0) > 0 || (r['v' + (prev||cur)] || 0) > 0; });

  setMainGrid(cols, rows, [], fmt(grandCur) + ' ' + sym + ' · ' + state.cat);
}

// ═══════════════════════════════════════════════════════════
// VIEW 4: DETALLE SEMANAL
// Flat table of all weekly_detail rows
// ═══════════════════════════════════════════════════════════
function renderDetalle() {
  var yrs  = getActiveYears();
  var sym  = state.currency.toUpperCase();

  var cols = [
    { field: 'year',      headerName: 'AÑO',     width: 60,  filter: 'agNumberColumnFilter', type: 'numericColumn', pinned: 'left' },
    { field: 'week',      headerName: 'SEM',      width: 55,  filter: 'agNumberColumnFilter', type: 'numericColumn', pinned: 'left',
      cellRenderer: function(p) { return wFmt(p.value); } },
    { field: 'categoria', headerName: 'CATEGORÍA', width: 220, filter: 'agTextColumnFilter', pinned: 'left', cellRenderer: catRenderer },
    { field: 'usd_total', headerName: 'USD',      width: 100, filter: 'agNumberColumnFilter', type: 'numericColumn', cellRenderer: moneyRenderer },
    { field: 'mxn_total', headerName: 'MXN',      width: 110, filter: 'agNumberColumnFilter', type: 'numericColumn', cellRenderer: moneyRenderer },
    { field: 'date_range',headerName: 'PERÍODO',  width: 150, filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:#888;font-size:10px">' + (p.value||'') + '</span>'; } },
  ];
  // Ranch columns
  RANCH_ORDER.forEach(function(r) {
    var col2 = RANCH_COLORS[r] || '#888';
    cols.push({
      field: 'rn_' + r.replace(/[^a-zA-Z0-9]/g,'_'),
      headerName: r, width: 100,
      filter: 'agNumberColumnFilter', type: 'numericColumn',
      cellRenderer: function(p) {
        var v = p.value; if (!v || v < 0.01) return '<span style="color:#ddd">—</span>';
        return '<span style="color:' + col2 + '">' + fmt(v) + '</span>';
      }
    });
  });

  var rows = [];
  var grandTotal = 0;
  DATA.weekly_detail.forEach(function(r) {
    if (!state.activeYears[r.year]) return;
    if (r.categoria !== state.cat) return;

    var row = {
      year: r.year, week: r.week, categoria: r.categoria,
      usd_total: r.usd_total, mxn_total: r.mxn_total,
      date_range: r.date_range || ''
    };
    RANCH_ORDER.forEach(function(rn) {
      var src = state.currency === 'usd' ? r.usd_ranches : r.mxn_ranches;
      row['rn_' + rn.replace(/[^a-zA-Z0-9]/g,'_')] = src[rn] || 0;
    });
    grandTotal += state.currency === 'usd' ? r.usd_total : r.mxn_total;
    rows.push(row);
  });
  rows.sort(function(a,b) { return b.year !== a.year ? b.year - a.year : b.week - a.week; });
  setMainGrid(cols, rows, [], fmt(grandTotal) + ' ' + sym + ' (' + rows.length + ' registros) · ' + state.cat);
}

// ═══════════════════════════════════════════════════════════
// VIEW 5: PRODUCTOS (PR + MP + ME)
// ═══════════════════════════════════════════════════════════
function renderProductosFull() {
  var cols = [
    { field: 'tipo',     headerName: 'TIPO',     width: 60,  filter: 'agTextColumnFilter', pinned: 'left' },
    { field: 'cat',      headerName: 'CAT',      width: 55,  filter: 'agTextColumnFilter', pinned: 'left',
      cellRenderer: function(p) { var m = {'PR':'#16a34a','MP':'#7c3aed','ME':'#0369a1'}; return '<span style="color:'+(m[p.value]||'#666')+';font-weight:700">'+(p.value||'')+'</span>'; } },
    { field: 'week_code',headerName: 'WK',       width: 72,  filter: 'agNumberColumnFilter' },
    { field: 'rancho',   headerName: 'RANCHO',   width: 105, filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:'+(RANCH_COLORS[p.value]||'#666')+';font-weight:600">'+(p.value||'')+'</span>'; } },
    { field: 'producto', headerName: 'PRODUCTO', width: 240, filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:#1e3a5f">' + (p.value||'') + '</span>'; } },
    { field: 'unidades', headerName: 'UNID.',    width: 80,  filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:#555">' + (p.value||'—') + '</span>'; } },
    { field: 'gasto',    headerName: 'GASTO',    width: 100, filter: 'agNumberColumnFilter', type: 'numericColumn', cellRenderer: moneyRenderer },
  ];

  var rows = [];
  function flattenProd(dataSet, label) {
    if (!dataSet) return;
    Object.keys(dataSet).forEach(function(wkCode) {
      var byRanch = dataSet[wkCode];
      Object.keys(byRanch).forEach(function(ranch) {
        var byTipo = byRanch[ranch];
        Object.keys(byTipo).forEach(function(tipo) {
          var items = byTipo[tipo];
          if (!Array.isArray(items)) return;
          items.forEach(function(item) {
            rows.push({
              cat: label, tipo: tipo, week_code: parseInt(wkCode) || wkCode,
              rancho: ranch, producto: item[0] || '',
              unidades: item[1] || '—', gasto: parseFloat(item[2]) || 0
            });
          });
        });
      });
    });
  }
  flattenProd(DATA.productos,    'PR');
  flattenProd(DATA.productos_mp, 'MP');
  flattenProd(DATA.productos_me, 'ME');
  rows.sort(function(a,b) {
    if (b.week_code !== a.week_code) return (b.week_code||0) - (a.week_code||0);
    return (a.rancho||'').localeCompare(b.rancho||'');
  });
  var total = rows.reduce(function(s,r) { return s + (r.gasto||0); }, 0);
  setMainGrid(cols, rows, [], fmt(total) + ' · ' + rows.length + ' registros');
}

// ═══════════════════════════════════════════════════════════
// VIEW 6: COSTO SERVICIOS
// ═══════════════════════════════════════════════════════════
var SV_SUBCATS = ['Electricidad','Fletes y Acarreos','Gastos de Exportación','Certificado Fitosanitario',
  'Transporte de Personal','Compra de Flor a Terceros','Comida para el Personal','RO, TEL, RTA.Alim'];
function renderServicios() {
  var yrs  = getActiveYears();
  var sym  = state.currency.toUpperCase();

  // Build rows from servicios_data (estructura nueva del extractor)
  // Fallback: weekly_detail con categorias SV: para compatibilidad.
  var svRows = {};
  if (Array.isArray(DATA.servicios_data) && DATA.servicios_data.length) {
    DATA.servicios_data.forEach(function(r) {
      if (!state.activeYears[r.year]) return;
      var subcat = (r.subcat || '').trim();
      if (!subcat) return;
      if (!svRows[subcat]) svRows[subcat] = {};

      var src = state.currency === 'usd' ? (r.usd_ranches || {}) : (r.mxn_ranches || {});
      RANCH_ORDER.forEach(function(rn) {
        var v = src[rn] || 0;
        if (v > 0) svRows[subcat][rn] = (svRows[subcat][rn] || 0) + v;
      });

      var total = state.currency === 'usd' ? r.usd_total : r.mxn_total;
      svRows[subcat]._total = (svRows[subcat]._total || 0) + (total || 0);
    });
  } else {
    DATA.weekly_detail.forEach(function(r) {
      if (!state.activeYears[r.year]) return;
      if (!r.categoria || !r.categoria.startsWith('SV:')) return;
      var subcat = r.categoria.replace('SV:','');
      if (!svRows[subcat]) svRows[subcat] = {};
      RANCH_ORDER.forEach(function(rn) {
        var src = state.currency === 'usd' ? r.usd_ranches : r.mxn_ranches;
        var v   = src[rn] || 0;
        if (v > 0) svRows[subcat][rn] = (svRows[subcat][rn] || 0) + v;
      });
      svRows[subcat]._total = (svRows[subcat]._total || 0) + (state.currency === 'usd' ? r.usd_total : r.mxn_total);
    });
  }

  var cols = [
    { field: 'subcat', headerName: 'SUBCATEGORÍA', pinned: 'left', width: 210, filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="font-weight:700;color:#1e3a5f">'+(p.value||'')+'</span>'; } },
    { field: 'total', headerName: 'TOTAL ' + sym, width: 110, type: 'numericColumn', cellRenderer: moneyRenderer },
    { field: 'pct',   headerName: '% DEL TOTAL', width: 85,  type: 'numericColumn',
      cellRenderer: function(p) {
        var v = p.value; if (!v) return '—';
        var w = Math.min(v / 100 * 55, 55);
        return '<div style="display:flex;align-items:center;gap:4px">' +
          '<div style="width:' + w.toFixed(0) + 'px;height:6px;background:#16a34a;border-radius:1px"></div>' +
          '<span>' + v.toFixed(1) + '%</span></div>';
      }
    },
  ];
  RANCH_ORDER.forEach(function(r) {
    var col3 = RANCH_COLORS[r] || '#888';
    cols.push({
      field: 'r_' + r.replace(/[^a-zA-Z0-9]/g,'_'),
      headerName: r, width: 100, type: 'numericColumn', filter: 'agNumberColumnFilter',
      cellRenderer: function(p) {
        var v = p.value; if (!v || v < 0.01) return '<span style="color:#e0e0e0">—</span>';
        return '<span style="color:' + col3 + '">' + fmt(v) + '</span>';
      }
    });
  });

  var grandTotal = Object.values(svRows).reduce(function(s,r) { return s + (r._total||0); }, 0);
  var orderedSubcats = SV_SUBCATS.filter(function(sc) { return svRows[sc]; });
  Object.keys(svRows).forEach(function(sc) {
    if (orderedSubcats.indexOf(sc) === -1) orderedSubcats.push(sc);
  });

  var rows = orderedSubcats.map(function(sc) {
    var data = svRows[sc] || {};
    var row = { subcat: sc, total: data._total || 0 };
    row.pct = grandTotal > 0 ? (data._total || 0) / grandTotal * 100 : 0;
    RANCH_ORDER.forEach(function(r) {
      row['r_' + r.replace(/[^a-zA-Z0-9]/g,'_')] = data[r] || 0;
    });
    return row;
  });
  rows.sort(function(a,b) { return b.total - a.total; });
  setMainGrid(cols, rows, [], fmt(grandTotal) + ' ' + sym);
}

// ═══════════════════════════════════════════════════════════
// PRODUCTOS SUBPANEL (click on cell)
// ═══════════════════════════════════════════════════════════
function onMainCellClick(evt) {
  if (!evt || !evt.data || !evt.colDef) return;

  var data = evt.data;
  var clickedField = evt.colDef.field || '';
  var clickedRanch = fieldToRanch(clickedField);

  if (state.view === 'semana') {
    showProdPanel(data, { ranch: clickedRanch || null });
    return;
  }
  if (state.view === 'comparativo') {
    if (clickedRanch || clickedField === 'total' || clickedField === 'week_lbl' || clickedField === 'week') {
      showProdPanel(data, { ranch: clickedRanch || null });
    }
    return;
  }
  if (state.view === 'rancho') {
    if (clickedField === 'rancho' || clickedRanch || clickedField.indexOf('cat_') === 0) {
      showProdPanel(data, { ranch: data.rancho || null });
    }
  }
}
function showProdPanel(rowData, opts) {
  opts = opts || {};
  var cat   = rowData._cat;
  var yr    = rowData._year;
  var wn    = rowData._week;
  var fromW = rowData._fromWeek || wn;
  var toW   = rowData._toWeek || wn;
  var ranchFilter = opts.ranch || null;
  if (!cat || !yr) return;

  var isMant   = cat === 'MANTENIMIENTO';
  var isMatEmp = cat === 'MATERIAL DE EMPAQUE';
  var isMirfe  = cat === CAT_MIRFE;
  var isMipe   = cat === CAT_MIPE;

  // Regla: abrir siempre desde comparativo/semana/rancho.
  // Fuente preferente por categoría; fallback general a PR.
  var src = isMant ? 'mp' : (isMatEmp ? 'me' : 'pr');
  var tipoFilter = null;
  if (src === 'pr') {
    if (isMirfe) tipoFilter = 'MIRFE';
    else if (isMipe) tipoFilter = 'MIPE';
  }

  var dsMap  = { pr: DATA.productos, mp: DATA.productos_mp, me: DATA.productos_me };
  var ds     = dsMap[src] || {};

  var wkStart = parseInt(fromW || wn || 0);
  var wkEnd   = parseInt(toW || wn || 0);
  if (!wkStart || !wkEnd) return;
  if (wkStart > wkEnd) {
    var t = wkStart; wkStart = wkEnd; wkEnd = t;
  }

  var rows = [];
  for (var wk = wkStart; wk <= wkEnd; wk++) {
    var wkCodeLong = (yr * 100) + wk;
    var wkCodeShort = ((yr % 100) * 100) + wk;
    var weekData = ds[wkCodeShort] || ds[String(wkCodeShort)] || ds[wkCodeLong] || ds[String(wkCodeLong)];
    if (!weekData) continue;

    Object.keys(weekData).forEach(function(ranch) {
      if (ranchFilter && ranch !== ranchFilter) return;
      var byTipo = weekData[ranch];
      Object.keys(byTipo).forEach(function(tipo) {
        if (tipoFilter && tipo !== tipoFilter) return;
        (byTipo[tipo] || []).forEach(function(item) {
          rows.push({
            week_code: wkCodeShort,
            rancho: ranch,
            tipo: tipo,
            producto: item[0] || '',
            unidades: item[1] || '',
            gasto: parseFloat(item[2]) || 0
          });
        });
      });
    });
  }

  var rangeText = wkStart === wkEnd ? (wFmt(wkStart) + ' · ' + yr) : (wFmt(wkStart) + '→' + wFmt(wkEnd) + ' · ' + yr);

  // Mostrar panel siempre al abrir detalle (aunque no haya filas)
  document.getElementById('prodPanel').className = 'show';

  // Inicializar grid de productos en primer uso
  if (!prodGridApi) {
    var prodElInit = document.getElementById('prodGrid');
    var initOpts = {
      columnDefs: getProdCols(), rowData: [],
      rowHeight: 20, headerHeight: 23,
      domLayout: 'autoHeight',
      defaultColDef: { sortable: true, filter: true, resizable: true },
      onGridReady: function(p) { prodGridApi = p.api; prodGridApi.sizeColumnsToFit(); }
    };
    new agGrid.Grid(prodElInit, initOpts);
  }

  if (rows.length === 0) {
    document.getElementById('prodTitle').textContent = cat + ' — Sin datos de productos';
    document.getElementById('prodMeta').textContent = rangeText + (ranchFilter ? (' · ' + ranchFilter) : '');
    if (prodGridApi) {
      prodGridApi.setColumnDefs(getProdCols());
      prodGridApi.setRowData([]);
      prodGridApi.sizeColumnsToFit();
    }
    return;
  }

  document.getElementById('prodTitle').textContent = cat + ' ▸ ' + rangeText + (ranchFilter ? (' · ' + ranchFilter) : '');

  rows.sort(function(a,b) { return b.gasto - a.gasto; });

  var total = rows.reduce(function(s,r) { return s + r.gasto; }, 0);
  document.getElementById('prodMeta').textContent = rows.length + ' registros · ' + fmt(total);

  if (prodGridApi) {
    prodGridApi.setColumnDefs(getProdCols());
    prodGridApi.setRowData(rows);
    prodGridApi.sizeColumnsToFit();
    setTimeout(reportHeight, 80);
  }
}
function getProdCols() {
  return [
    { field: 'week_code', headerName: 'WK', width: 72, filter: 'agNumberColumnFilter' },
    { field: 'rancho', headerName: 'RANCHO', width: 110, pinned: 'left', filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:'+(RANCH_COLORS[p.value]||'#666')+';font-weight:600">'+(p.value||'')+'</span>'; } },
    { field: 'tipo',   headerName: 'TIPO',   width: 65, filter: 'agTextColumnFilter' },
    { field: 'producto', headerName: 'PRODUCTO', width: 280, filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="color:#1e3a5f">'+(p.value||'')+'</span>'; } },
    { field: 'unidades', headerName: 'UNID.', width: 90 },
    { field: 'gasto', headerName: 'GASTO USD', width: 100, type: 'numericColumn', cellRenderer: moneyRenderer },
  ];
}
function closeProdPanel() {
  document.getElementById('prodPanel').className = '';
}

// ── HELPER: abrir productos desde tabla comparativo ──
function showProdFromCmp(yr, wk, ranch) {
  var rowData = { _cat: state.cat, _year: yr, _week: wk, _fromWeek: wk, _toWeek: wk };
  showProdPanel(rowData, { ranch: ranch || null });
}

// Delegated click para celdas clickeables de comparativo
document.addEventListener('click', function(e) {
  var td = e.target.closest('td.cmp-clickable');
  if (!td) return;
  var yr    = parseInt(td.dataset.yr);
  var wk    = parseInt(td.dataset.wk);
  var ranch = td.dataset.ranch || null;
  showProdFromCmp(yr, wk, ranch || null);
});

// ═══════════════════════════════════════════════════════════
// RESIZE HELPER
// ═══════════════════════════════════════════════════════════
function resizeGrid() {
  if (mainGridApi) mainGridApi.sizeColumnsToFit();
}
window.addEventListener('resize', resizeGrid);

// ═══════════════════════════════════════════════════════════
// HEIGHT REPORTING TO STREAMLIT
// ═══════════════════════════════════════════════════════════
function reportHeight() {
  var appEl = document.getElementById('app');
  var h = appEl ? appEl.scrollHeight + 20 : document.body.scrollHeight + 20;
  window.parent.postMessage({ type: 'streamlit:setFrameHeight', height: Math.max(h, 400) }, '*');
}
var ro = new ResizeObserver(reportHeight);
ro.observe(document.body);
reportHeight();
setInterval(reportHeight, 500);

// ═══════════════════════════════════════════════════════════
// ERROR HANDLER
// ═══════════════════════════════════════════════════════════
window.onerror = function(msg, src, line) {
  document.getElementById('loader').innerHTML =
    '<div style="color:#dc2626;font-family:monospace;padding:20px;background:#fff;border-radius:8px;border:1px solid #fecaca;max-width:600px">' +
    '<b>ERROR JS:</b> ' + msg + '<br><small>línea ' + line + '</small></div>';
  return true;
};

// ═══════════════════════════════════════════════════════════
// ARRANCAR
// ═══════════════════════════════════════════════════════════
// Reconstruir weekly_series desde weekly_detail si no existe
if (!DATA.weekly_series) {
  DATA.weekly_series = {};
  DATA.categories.forEach(function(cat) { DATA.weekly_series[cat] = {}; });
  DATA.weekly_detail.forEach(function(r) {
    if (r.usd_total > 0) {
      if (!DATA.weekly_series[r.categoria]) DATA.weekly_series[r.categoria] = {};
      var key = r.year + '-W' + String(r.week).padStart(2,'0');
      DATA.weekly_series[r.categoria][key] = (DATA.weekly_series[r.categoria][key] || 0) + r.usd_total;
    }
  });
}
// Wait for AG Grid to load
if (typeof agGrid === 'undefined') {
  var checkAG = setInterval(function() {
    if (typeof agGrid !== 'undefined') { clearInterval(checkAG); inicializar(); }
  }, 100);
} else {
  inicializar();
}
</script>


</body>
</html>"""

html_final = HTML.replace('__DATA_JSON__', data_json)
components.html(html_final, height=900, scrolling=True)

# ─── Barra inferior: Descarga XLSX + Panel Crear Hoja WK ─────────────────────
st.markdown("""
<style>
  div[data-testid="stSelectbox"] > div { min-width:120px !important; }

  /* Botón ☰ estilo toolbar */
  div[data-testid="stButton"] button[kind="secondary"].menu-btn {
    font-family: monospace; font-size: 14px;
    background: #1e3a5f; color: #fff;
    border: none; border-radius: 4px;
    padding: 2px 10px; height: 38px;
  }

  /* Panel crear hoja */
  .crear-panel {
    background: #1e3a5f;
    border-top: 3px solid #16a34a;
    padding: 14px 18px 12px;
    display: flex; align-items: center; gap: 12px;
    flex-wrap: wrap;
  }
  .crear-panel-title {
    color: rgba(255,255,255,0.55); font-size: 10px;
    font-family: monospace; text-transform: uppercase;
    letter-spacing: 0.6px; white-space: nowrap;
  }
</style>
""", unsafe_allow_html=True)

# Estado del panel
if "show_crear_panel" not in st.session_state:
    st.session_state.show_crear_panel = False

# Construir lista de semanas disponibles (código YYWW)
available_weeks = sorted(
    {
        str(r["year"] % 100).zfill(2) + str(r["week"]).zfill(2)
        for r in DATA.get("weekly_detail", [])
    },
    reverse=True
)

if available_weeks:
    from data_extractor import get_sheet_xlsx
    try:
        from data_extractor import crear_hoja_wk
        _crear_disponible = True
    except ImportError:
        _crear_disponible = False

    # ── Fila de controles ─────────────────────────────────────────────────────
    col1, col2, col3, col_menu = st.columns([1.2, 1, 5, 0.18])

    with col1:
        selected_wk = st.selectbox(
            "⬇ Descargar hoja WK",
            options=available_weeks,
            format_func=lambda c: f"WK{c}",
            label_visibility="visible",
        )

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Descargar XLSX", key="dl_xlsx"):
            with st.spinner(f"Preparando WK{selected_wk}..."):
                xlsx_bytes = get_sheet_xlsx(selected_wk)
            if xlsx_bytes:
                st.download_button(
                    label=f"💾 WK{selected_wk}.xlsx",
                    data=xlsx_bytes,
                    file_name=f"WK{selected_wk}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_btn",
                )
            else:
                st.error(f"No se encontró la hoja WK{selected_wk} en el archivo.")

    with col_menu:
        st.markdown("<br>", unsafe_allow_html=True)
        if _crear_disponible:
            if st.button("☰", key="toggle_crear", help="Crear nueva hoja WK en SharePoint"):
                st.session_state.show_crear_panel = not st.session_state.show_crear_panel

    # ── Panel: Crear nueva hoja WK ────────────────────────────────────────────
    if _crear_disponible and st.session_state.show_crear_panel:
        st.markdown(
            '<div class="crear-panel">'
            '<span class="crear-panel-title">➕ Nueva hoja WK en SharePoint</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        pc1, pc2, pc3, pc4 = st.columns([1.2, 0.8, 0.8, 4])

        with pc1:
            nuevo_nombre = st.text_input(
                "Nombre de la hoja",
                placeholder="Ej: WK2518",
                key="nuevo_wk_nombre",
                label_visibility="visible",
            ).strip().upper()

        with pc2:
            st.markdown("<br>", unsafe_allow_html=True)
            crear_clicked = st.button("✚ Crear hoja", key="btn_crear_hoja", type="primary")

        with pc3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✕ Cerrar", key="btn_cerrar_panel"):
                st.session_state.show_crear_panel = False
                st.rerun()

        if crear_clicked:
            if not nuevo_nombre:
                st.warning("⚠️ Escribe el nombre de la hoja antes de crear.")
            elif not nuevo_nombre.startswith("WK") or len(nuevo_nombre) != 6:
                st.warning("⚠️ El nombre debe tener formato WK#### (ej: WK2518).")
            else:
                # Leer credenciales desde st.secrets  (sección [sharepoint])
                try:
                    tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                    client_id     = st.secrets["sharepoint"]["client_id"]
                    client_secret = st.secrets["sharepoint"]["client_secret"]
                except KeyError as e:
                    st.error(
                        f"❌ Falta la credencial **{e}** en los secrets de Streamlit. "
                        "Revisa que `.streamlit/secrets.toml` tenga la sección [sharepoint]."
                    )
                    st.stop()

                with st.spinner(f"Creando hoja {nuevo_nombre} en SharePoint…"):
                    resultado = crear_hoja_wk(
                        nuevo_nombre, tenant_id, client_id, client_secret
                    )

                if resultado.get("ok"):
                    st.success(resultado["mensaje"])
                    st.cache_data.clear()   # Refrescar datos al recargar
                else:
                    st.error(f"❌ {resultado['error']}")
