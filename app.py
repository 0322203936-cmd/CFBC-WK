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
.tb-search {
  font-size: 11px; font-family: var(--mono);
  border: 1px solid #bbb; border-radius: 3px;
  padding: 2px 8px; height: 22px; width: 130px;
  background: #fff;
}
.tb-search:focus { outline: 2px solid var(--green); outline-offset: -1px; }

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
    <div class="tb-sep"></div>
    <input type="text" class="tb-search" id="quickFilter" placeholder="🔍  filtrar tabla..." oninput="onQuickFilter(this.value)">
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
    <div id="myGrid" class="ag-theme-alpine" style="width:100%;height:500px"></div>
  </div>

  <!-- PRODUCTOS SUB-PANEL -->
  <div id="prodPanel">
    <div class="prod-hdr">
      <div class="prod-hdr-title" id="prodTitle">PRODUCTOS</div>
      <div class="prod-hdr-meta" id="prodMeta"></div>
      <button class="prod-close" onclick="closeProdPanel()">✕ CERRAR</button>
    </div>
    <div id="prodGrid" class="ag-theme-alpine" style="width:100%;height:300px"></div>
  </div>

  <!-- STATUS BAR -->
  <div class="statusbar" id="statusbar">
    <span>Filas: <b id="stRows">—</b></span>
    <span class="st-sep">|</span>
    <span>Total: <b id="stTotal">—</b></span>
    <span class="st-sep">|</span>
    <span id="stWeekDate" style="color:#888"></span>
    <span style="margin-left:auto;color:#aaa;font-size:9px" id="stInfo"></span>
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

  // Rango inicial: todas las semanas del año más reciente
  state.fromWeek = wksLatest[0] || allWeeks[0] || 1;
  state.toWeek   = wksLatest[wksLatest.length - 1] || allWeeks[allWeeks.length - 1] || 52;

  buildCatSelect();
  buildYearChips();
  updateWeekControls();
  updateRangeSliders();
  buildMainGrid();
  renderView();
  updateHeader();
  document.getElementById('loader').style.display = 'none';
  document.getElementById('app').style.display   = 'block';
  setTimeout(resizeGrid, 80);
  setTimeout(resizeGrid, 300); // segundo llamado por si AG Grid tarda en inicializar
}

// ═══════════════════════════════════════════════════════════
// UI BUILDERS
// ═══════════════════════════════════════════════════════════
function buildCatSelect() {
  var el = document.getElementById('catSel');
  el.innerHTML = DATA.categories.map(function(c) {
    var label = c === CAT_MIRFE ? c + ' (MIRFE)' : c === CAT_MIPE ? c + ' (MIPE)' : c;
    return '<option value="' + c.replace(/"/g,'&quot;') + '"' + (c === state.cat ? ' selected' : '') + '>' + label + '</option>';
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
  document.getElementById('weekLabel').textContent = wFmt(wn) + ' · ' + yr;
  document.getElementById('stWeekDate').textContent = dateText;
}
function updateHeader() {
  var yrs = getActiveYears();
  var wn  = allWeeks[state.weekIdx] || 1;
  var curYr  = yrs[yrs.length - 1];
  var prevYr = yrs[yrs.length - 2];
  // Grand total for current week, SELECTED category, current year
  var grandTotal = 0;
  getWeekDetail(state.cat, wn, curYr).forEach(function(r) { grandTotal += state.currency === 'usd' ? r.usd_total : r.mxn_total; });

  var prevTotal = 0;
  if (prevYr) {
    getWeekDetail(state.cat, wn, prevYr).forEach(function(r) { prevTotal += state.currency === 'usd' ? r.usd_total : r.mxn_total; });
  }
  var delta = prevTotal > 0 ? (grandTotal - prevTotal) / prevTotal * 100 : null;
  
  // Annual total current year for SELECTED category
  var annualTotal = 0;
  var recsAnual = DATA.weekly_detail.filter(function(r) { return r.categoria === state.cat && r.year === curYr; });
  annualTotal = sumDetail(recsAnual, state.currency).total;

  var html = '';
  html += '<div class="hdr-kpi"><span class="hdr-kpi-label">SEMANA ' + wFmt(wn) + ' · ' + curYr + '</span><span class="hdr-kpi-val">' + fmt(grandTotal) + '</span>';
  if (delta !== null) {
    var cls = delta >= 0 ? 'pos' : 'neg';
    var arrow = delta >= 0 ? '▲' : '▼';
    html += '<span class="hdr-kpi-delta ' + cls + '">' + arrow + ' ' + Math.abs(delta).toFixed(1) + '%</span>';
  }
  html += '</div>';
  html += '<div class="hdr-kpi"><span class="hdr-kpi-label">ACUM. ANUAL ' + curYr + '</span><span class="hdr-kpi-val">' + fmt(annualTotal) + '</span></div>';
  html += '<div class="hdr-kpi"><span class="hdr-kpi-label">CATEGORÍA</span><span class="hdr-kpi-val" style="font-size:11px;max-width:220px;overflow:hidden;text-overflow:ellipsis">' + state.cat + '</span></div>';
  html += '<div class="hdr-kpi"><span class="hdr-kpi-label">MONEDA</span><span class="hdr-kpi-val">' + state.currency.toUpperCase() + '</span></div>';
  document.getElementById('hdrKpis').innerHTML = html;
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
  closeProdPanel();
  renderView();
}
function onQuickFilter(val) {
  if (mainGridApi) mainGridApi.setQuickFilter(val);
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
  if (fLbl) fLbl.textContent = wFmt(f);
  if (tLbl) tLbl.textContent = wFmt(t);
  var count = allWeeks.filter(function(w){ return w >= f && w <= t; }).length;
  if (badge) badge.textContent = wFmt(f) + ' → ' + wFmt(t) + ' · ' + count + ' sem';
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
  state.fromWeek = wks[0] || 1;
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
    onGridReady: function(params) { mainGridApi = params.api; },
    onCellClicked: function(e) { onMainCellClick(e); },
  };
  new agGrid.Grid(el, opts);
}
function setMainGrid(colDefs, rowData, pinnedBottom, statusText) {
  if (!mainGridApi) return;
  mainGridApi.setPinnedBottomRowData([]);   // limpiar siempre primero
  mainGridApi.setColumnDefs(colDefs);
  mainGridApi.setRowData(rowData);
  mainGridApi.sizeColumnsToFit();
  document.getElementById('stRows').textContent  = rowData.length;
  document.getElementById('stTotal').textContent = statusText || '';
  document.getElementById('stInfo').textContent  = state.view.toUpperCase() + ' · ' + state.cat;
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
  setTimeout(resizeGrid, 30);
}

// ═══════════════════════════════════════════════════════════
// VIEW 1: SEMANA
// Rows = years, Cols = [Year, Week, Cat, Total, Delta, ranches]
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
    var recs = DATA.weekly_detail.filter(function(r) { return r.categoria === cat && r.year === yr; });
    return sumDetail(recs, state.currency);
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
// VIEW 3: COMPARATIVO (filtrado por rango W## → W##)
// Rows = una fila por SEMANA en el rango (para la categoría activa)
// ═══════════════════════════════════════════════════════════
function renderComparativo() {
  var sym  = state.currency.toUpperCase();
  var yrs  = getActiveYears();
  var fromW = state.fromWeek;
  var toW   = state.toWeek;

  var cols = [
    { field: 'week', headerName: 'SEMANA', width: 90, pinned: 'left', filter: 'agTextColumnFilter',
      cellRenderer: function(p) { return '<span style="font-weight:800;color:#1e3a5f">' + (p.value === 'TOTAL' ? 'TOTAL' : wFmt(p.value)) + '</span>'; } }
  ];

  yrs.forEach(function(yr) {
    cols.push({
      field: 'y' + yr, headerName: String(yr) + ' ' + sym, width: 110, type: 'numericColumn',
      filter: 'agNumberColumnFilter', cellRenderer: moneyRenderer
    });
  });

  var curYr = yrs[yrs.length - 1];
  var prevYr = yrs.length > 1 ? yrs[yrs.length - 2] : null;

  if (prevYr) {
    cols.push({ field: 'deltaAmt', headerName: 'Δ $ ' + prevYr + '→' + curYr, width: 120, type: 'numericColumn', cellRenderer: deltaAmtRenderer });
    cols.push({ field: 'deltaPct', headerName: 'Δ %', width: 72, type: 'numericColumn', cellRenderer: deltaRenderer });
  }

  RANCH_ORDER.forEach(function(r) {
    cols.push({
      field: 'r_' + r.replace(/[^a-zA-Z0-9]/g,'_'), headerName: r, width: 100, type: 'numericColumn',
      filter: 'agNumberColumnFilter', cellRenderer: ranchRenderer(r)
    });
  });

  var rows = [];
  var totals = { deltaAmt: 0 };
  yrs.forEach(function(yr) { totals['y' + yr] = 0; });
  RANCH_ORDER.forEach(function(rn) { totals['r_' + rn.replace(/[^a-zA-Z0-9]/g,'_')] = 0; });

  for (var w = fromW; w <= toW; w++) {
    if (allWeeks.indexOf(w) === -1) continue;
    var row = { week: w, _cat: state.cat, _fromWeek: w, _toWeek: w };
    var hasValidData = false;

    yrs.forEach(function(yr) {
      var agg = sumDetail(getWeekDetail(state.cat, w, yr), state.currency);
      row['y' + yr] = agg.total;
      totals['y' + yr] += agg.total;
      if (agg.total > 0) hasValidData = true;

      if (yr === curYr) {
        RANCH_ORDER.forEach(function(rn) {
          var rv = agg.ranches[rn] || 0;
          row['r_' + rn.replace(/[^a-zA-Z0-9]/g,'_')] = rv;
          totals['r_' + rn.replace(/[^a-zA-Z0-9]/g,'_')] += rv;
        });
      }
    });

    if (prevYr) {
      row.deltaAmt = (row['y' + curYr] || 0)  - (row['y' + prevYr] || 0);
      row.deltaPct = (row['y' + prevYr] > 0) ? row.deltaAmt / row['y' + prevYr] * 100 : null;
      totals.deltaAmt += row.deltaAmt;
    }

    if (hasValidData) rows.push(row);
  }

  var bottomRow = { week: 'TOTAL' };
  yrs.forEach(function(yr) { bottomRow['y' + yr] = totals['y' + yr]; });
  if (prevYr) {
    bottomRow.deltaAmt = totals.deltaAmt;
    bottomRow.deltaPct = (totals['y' + prevYr] > 0) ? totals.deltaAmt / totals['y' + prevYr] * 100 : null;
  }
  RANCH_ORDER.forEach(function(rn) {
    bottomRow['r_' + rn.replace(/[^a-zA-Z0-9]/g,'_')] = totals['r_' + rn.replace(/[^a-zA-Z0-9]/g,'_')];
  });

  setMainGrid(cols, rows, [bottomRow], fmt(totals['y' + curYr]) + ' ' + sym + ' · ' + state.cat + ' (' + wFmt(fromW) + '→' + wFmt(toW) + ')');
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
    totalCur += aC.ranches[ranch] || 0;
    if (prev) {
      var aP = sumDetail(getWeekDetail(state.cat, wn, prev), state.currency);
      totalPrev += aP.ranches[ranch] || 0;
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
    setTimeout(function(){ document.getElementById('prodPanel').scrollIntoView({ behavior:'smooth', block:'nearest' }); }, 50);
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
  }
  setTimeout(function(){ document.getElementById('prodPanel').scrollIntoView({ behavior:'smooth', block:'nearest' }); }, 50);
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

// ═══════════════════════════════════════════════════════════
// RESIZE HELPER
// ═══════════════════════════════════════════════════════════
function resizeGrid() {
  // Medir la altura real de todos los elementos fijos alrededor del grid
  var hdr      = document.querySelector('.app-hdr');
  var toolbar  = document.querySelector('.toolbar');
  var tabs     = document.querySelector('.view-tabs');
  var rangeBar = document.querySelector('.range-bar');
  var statusbar= document.querySelector('.statusbar');
  var prodPanel= document.getElementById('prodPanel');

  var used = 0;
  if (hdr)       used += hdr.offsetHeight;
  if (toolbar)   used += toolbar.offsetHeight;
  if (tabs)      used += tabs.offsetHeight;
  if (rangeBar && rangeBar.classList.contains('show')) used += rangeBar.offsetHeight;
  if (statusbar) used += statusbar.offsetHeight;
  if (prodPanel && prodPanel.classList.contains('show')) used += prodPanel.offsetHeight;

  // document.documentElement.clientHeight = altura real del iframe
  var available = document.documentElement.clientHeight - used - 4;
  var h = Math.max(available, 300);
  document.getElementById('myGrid').style.height = h + 'px';
  if (mainGridApi) mainGridApi.sizeColumnsToFit();
}
window.addEventListener('resize', resizeGrid);

// ═══════════════════════════════════════════════════════════
// HEIGHT REPORTING TO STREAMLIT
// ═══════════════════════════════════════════════════════════
function reportHeight() {
  var appEl = document.getElementById('app');
  var h = appEl ? appEl.scrollHeight + 60 : document.body.scrollHeight + 60;
  window.parent.postMessage({ type: 'streamlit:setFrameHeight', height: Math.max(h, 700) }, '*');
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
components.html(html_final, height=800, scrolling=False)
