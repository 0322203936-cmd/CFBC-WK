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

st.markdown("""
<style>
  #MainMenu, header, footer { display: none !important; }
  .stApp { background: #f0f0f0; }
  .block-container { padding: 0 !important; max-width: 100% !important; margin-top: -1rem !important; }
  .stMainBlockContainer { padding-top: 0 !important; }
  section[data-testid="stSidebar"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


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
  background: #f0f0f0;
}

/* ── LOADER ─────────────────────────────────── */
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

/* ── HEADER ─────────────────────────────────── */
.app-hdr {
  background: #4472C4;
  border-bottom: 3px solid var(--green);
  padding: 5px 10px;
  display: flex; align-items: center; gap: 0;
  height: 36px; overflow: hidden;
}
.hdr-brand {
  color: #ffffff; font-size: 12px; font-weight: 700;
  letter-spacing: 1px; white-space: nowrap;
  padding-right: 12px; border-right: 1px solid rgba(255,255,255,0.3);
  flex-shrink: 0;
}
.hdr-btn {
  margin-left: auto; flex-shrink: 0;
  font-size: 10px; font-weight: 700;
  background: rgba(255,255,255,0.35);
  border: 1px solid rgba(255,255,255,0.35);
  border-radius: 3px; padding: 3px 10px; cursor: pointer;
  color: #ffffff; height: 24px;
  transition: background 0.1s; white-space: nowrap;
}
.hdr-btn:hover { background: rgba(255,255,255,0.55); }

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
  overflow: hidden;
}
.pt-table-wrap {
  overflow: auto;
  width: 100%;
  /* 36px header + 28px toolbar + 26px range-bar + 28px view-tabs = 118px aprox */
  max-height: calc(100vh - 118px);
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
  display: none; background: #fdfdfd; border: 1px solid #cbd5e1; border-top: 2px solid #0f172a;
  box-shadow: 0 4px 12px rgba(0,0,0,0.06);
  margin: 5px 0 0 0; width: 100%; overflow: hidden;
}
#prodPanel.show { display: block; }
#prodTableWrap { overflow: visible; }

/* ── STATUS BAR ──────────────────────────────── */
.statusbar {
  background: #ebebeb; border-top: 1px solid #ccc;
  padding: 2px 10px; font-size: 10px; color: #666;
  display: flex; gap: 14px; align-items: center;
  height: 22px; overflow: hidden;
}
.statusbar b { color: #333; }
</style>"""

APP_HTML_BODY = """
<!-- LOADER -->
<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">CFBC &#8212; Cargando datos...</div>
</div>

<!-- APP -->
<div id="app" style="display:none">

  <!-- HEADER -->
  <div class="app-hdr">
    <div class="hdr-brand">CFBC &#9656; CONTROL SEMANAL</div>
    <button class="hdr-btn" onclick="exportCSV()" style="margin-left:auto">&#11015; CSV</button>
    <button class="hdr-btn" onclick="recargar()" style="margin-left:4px">&#8635;</button>
  </div>

  <!-- TOOLBAR -->
  <div class="toolbar">
    <span class="tb-label">Cat</span>
    <select class="tb-sel" id="catSel" onchange="onCatChange(this.value)" style="max-width:200px"></select>
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
  </div>

  <!-- VIEW TABS -->
  <div class="view-tabs">
    <button class="vtab"        id="vtAnual"        onclick="setView('anual')">Anual</button>
    <button class="vtab active" id="vtComparativo"  onclick="setView('comparativo')">Comparativo</button>
    <button class="vtab"        id="vtRancho"       onclick="setView('rancho')">Por Rancho</button>
    <button class="vtab"        id="vtServicios"    onclick="setView('servicios')">Costo Servicios</button>
  </div>

  <!-- RANGE BAR eliminada — controles movidos al toolbar -->

  <!-- MAIN TABLE AREA (todas las vistas excepto comparativo) -->
  <div id="gridWrap">
    <div class="pt-table-wrap" id="tableWrap" style="overflow:auto"></div>
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
var YEAR_COLORS = {2021:'#0ea5e9',2022:'#d97706',2023:'#16a34a',2024:'#9333ea',2025:'#f97316',2026:'#dc2626'};
var CAT_MIRFE = 'FERTILIZANTES';
var CAT_MIPE  = 'DESINFECCION / PLAGUICIDAS';

// =======================================================
// ESTADO
// =======================================================
var state = { cat:'', currency:'mxn', activeYears:{}, view:'comparativo', weekIdx:0, fromWeek:1, toWeek:52 };
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
  });

  var wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '<table class="pt-table"><thead>'+headHtml+'</thead><tbody>'+bodyHtml+'</tbody></table>';

  if (statusText !== undefined) document.getElementById('stTotal').textContent = statusText;
}

// Delegated click sobre tableWrap
document.addEventListener('click', function(e) {
  var td = e.target.closest('#tableWrap td');
  if (!td) return;
  var tr = td.closest('tr');
  var ri = parseInt(tr.dataset.ri);
  var ci = parseInt(td.dataset.ci);
  if (isNaN(ri)||isNaN(ci)) return;
  var row = _tableRows[ri];
  var col = _tableColDefs[ci];
  if (!row||!col) return;
  onMainCellClick({data:row, colDef:col});
});

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

  // Semanas del año más reciente (considerar también mano_obra_data)
  var wksLatest = DATA.weekly_detail
    .filter(function(r){return r.year===latestYr;})
    .map(function(r){return r.week;})
    .filter(function(v,i,a){return a.indexOf(v)===i;})
    .sort(function(a,b){return a-b;});
  if (!wksLatest.length && Array.isArray(DATA.mano_obra_data)) {
    wksLatest = DATA.mano_obra_data
      .filter(function(r){return r.year===latestYr;})
      .map(function(r){return r.week;})
      .filter(function(v,i,a){return a.indexOf(v)===i;})
      .sort(function(a,b){return a-b;});
  }
  var curWeek = wksLatest[wksLatest.length-1] || allWeeks[allWeeks.length-1];
  var idx = allWeeks.indexOf(curWeek);
  state.weekIdx = idx>=0 ? idx : allWeeks.length-1;

  state.toWeek   = wksLatest[wksLatest.length-1] || allWeeks[allWeeks.length-1] || 52;
  state.fromWeek = wksLatest[wksLatest.length-2] || wksLatest[0] || state.toWeek;

  buildCatSelect();
  buildYearChips();
  updateWeekControls();
  updateRangeSliders();
  // Ocultar tab Costo Servicios si la cat inicial no es de tipo servicio
  var vtSrv = document.getElementById('vtServicios');
  var _isSrvCat = (state.cat === 'COSTO SERVICIOS' || state.cat === 'COSTO MANO DE OBRA');
  if (vtSrv) vtSrv.style.display = _isSrvCat ? '' : 'none';
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
    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+'>'+c+'</option>';
  }).join('');
}
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

// =======================================================
// EVENTS
// =======================================================
function onCatChange(val) {
  state.cat = val;
  var isSrvCat = (val === 'COSTO SERVICIOS' || val === 'COSTO MANO DE OBRA');
  ['Anual','Comparativo','Rancho'].forEach(function(name) {
    var el = document.getElementById('vt' + name);
    if (el) el.style.display = isSrvCat ? 'none' : '';
  });
  var vtSrv = document.getElementById('vtServicios');
  if (vtSrv) vtSrv.style.display = isSrvCat ? '' : 'none';

  // Ajustar rango de semanas según la fuente de datos
  if (val === 'COSTO MANO DE OBRA' && Array.isArray(DATA.mano_obra_data) && DATA.mano_obra_data.length) {
    // Usar las semanas disponibles en mano_obra_data para los años activos
    var moWeeks = DATA.mano_obra_data
      .filter(function(r){ return state.activeYears[r.year]; })
      .map(function(r){ return r.week; })
      .filter(function(v,i,a){ return a.indexOf(v)===i; })
      .sort(function(a,b){ return a-b; });
    if (moWeeks.length) {
      state.fromWeek = moWeeks[0];
      state.toWeek   = moWeeks[moWeeks.length-1];
      updateRangeSliders();
    }
  }

  if (isSrvCat && state.view !== 'servicios') {
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
}
function toggleYear(y) {
  var active = DATA.years.filter(function(yr){return state.activeYears[yr];});
  if (state.activeYears[y]&&active.length>1) delete state.activeYears[y];
  else state.activeYears[y]=true;
  buildYearChips();
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
  ['anual','comparativo','rancho','servicios'].forEach(function(name){
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
  if(fLbl)fLbl.textContent=yy+String(f).padStart(2,'0');
  if(tLbl)tLbl.textContent=yy+String(t).padStart(2,'0');
  var count=allWeeks.filter(function(w){return w>=f&&w<=t;}).length;
  if(badge)badge.textContent=yy+String(f).padStart(2,'0')+' → '+yy+String(t).padStart(2,'0')+' · '+count+' sem';
}
function onRangeChange() {
  var f=parseInt(document.getElementById('fromSlider').value);
  var t=parseInt(document.getElementById('toSlider').value);
  if (f>t){var tmp=f;f=t;t=tmp;}
  state.fromWeek=f; state.toWeek=t;
  updateRangeSliders();
  renderView();
}
function resetRange() {
  var latestYr=DATA.years[DATA.years.length-1];
  var wks=DATA.weekly_detail.filter(function(r){return r.year===latestYr;}).map(function(r){return r.week;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort(function(a,b){return a-b;});
  state.toWeek   = wks[wks.length-1]||allWeeks[allWeeks.length-1]||52;
  state.fromWeek = wks[wks.length-2]||wks[0]||state.toWeek;
  updateRangeSliders();
  if (state.view==='comparativo') renderComparativo();
}

// =======================================================
// VIEW ROUTER
// =======================================================
function renderView() {
  document.getElementById('prodPanel').className='';
  if      (state.view==='anual')       renderAnual();
  else if (state.view==='comparativo') renderComparativo();
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
  var cols=[
    { field:'year', headerName:'AÑO', width:70, type:'numericColumn', pinned:'left',
      cellRenderer:function(p){ var c=YEAR_COLORS[p.value]||'#888'; return '<span style="color:'+c+';font-weight:700">'+p.value+'</span>'; }},
    { field:'cat_label', headerName:'CATEGORÍA', width:170, pinned:'left', cellRenderer:catRenderer },
    { field:'total',    headerName:'TOTAL '+sym, width:110, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'deltaAmt', headerName:'Δ $',        width:90,  type:'numericColumn', cellRenderer:deltaAmtRenderer },
    { field:'deltaPct', headerName:'Δ %',        width:72,  type:'numericColumn', cellRenderer:deltaRenderer },
  ];
  RANCH_ORDER.forEach(function(r){ cols.push({field:ranchFieldName(r),headerName:r,width:100,type:'numericColumn',cellRenderer:ranchRenderer(r)}); });

  var getYrAgg=function(cat,yr){
    var d=(DATA.summary[cat]||{})[yr]||{usd:0,mxn:0,ranches:{},ranches_mxn:{}};
    return {total:state.currency==='usd'?d.usd:d.mxn, ranches:state.currency==='usd'?d.ranches:d.ranches_mxn};
  };
  var rows=[]; var grandTotal=0;
  yrs.forEach(function(yr,i){
    var prevYr=i>0?yrs[i-1]:null;
    var agg=getYrAgg(state.cat,yr);
    var row={year:yr,cat_label:state.cat,_cat:state.cat,_year:yr};
    row.total=agg.total;
    if (prevYr){ var aggP=getYrAgg(state.cat,prevYr); row.deltaAmt=agg.total-aggP.total; row.deltaPct=aggP.total>0?(agg.total-aggP.total)/aggP.total*100:null; }
    RANCH_ORDER.forEach(function(r){ row[ranchFieldName(r)]=agg.ranches[r]||0; });
    rows.push(row);
    if (yr===yrs[yrs.length-1]) grandTotal+=agg.total;
  });
  renderPivotTable(cols, rows, fmt(grandTotal)+' '+sym+' · AÑO '+yrs[yrs.length-1]);
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
  var ranchCols=RANCH_ORDER;
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
      var val=d?(state.currency==='usd'?d.usd:d.mxn):0;
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
  var grandTotal=yrs.reduce(function(s,yr){var d=byYear[yr];return s+(d?(state.currency==='usd'?d.usd:d.mxn):0);},0);
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
  var yrs=getActiveYears(), wn=allWeeks[state.weekIdx]||1;
  var cur=yrs[yrs.length-1], prev=yrs.length>1?yrs[yrs.length-2]:null, sym=state.currency.toUpperCase();
  var cols=[
    { field:'rancho', headerName:'RANCHO', pinned:'left', width:150,
      cellRenderer:function(p){ var c=RANCH_COLORS[p.value]||'#888'; return '<span style="color:'+c+';font-weight:700">'+(p.value||'')+'</span>'; }}
  ];
  if (prev) cols.push({field:'v'+prev, headerName:String(prev)+' '+sym, width:120, type:'numericColumn', cellRenderer:moneyRenderer});
  cols.push({field:'v'+cur, headerName:String(cur)+' '+sym+' &#9733;', width:120, type:'numericColumn', cellRenderer:moneyRenderer});
  if (prev) {
    cols.push({field:'deltaAmt',headerName:'Δ $',width:100,type:'numericColumn',cellRenderer:deltaAmtRenderer});
    cols.push({field:'deltaPct',headerName:'Δ %',width:90, type:'numericColumn',cellRenderer:deltaRenderer});
  }
  var grandCur=0,grandPrev=0;
  var rows=RANCH_ORDER.map(function(ranch){
    var row={rancho:ranch,_cat:state.cat,_week:wn,_year:cur,_fromWeek:wn,_toWeek:wn};
    var aC=sumDetail(getWeekDetail(state.cat,wn,cur),state.currency);
    var totalCur=aC.ranches[ranch]||0;
    var totalPrev=0;
    if (prev){ var aP=sumDetail(getWeekDetail(state.cat,wn,prev),state.currency); totalPrev=aP.ranches[ranch]||0; }
    row['v'+cur]=totalCur; grandCur+=totalCur;
    if (prev){ row['v'+prev]=totalPrev; grandPrev+=totalPrev; row.deltaAmt=totalCur-totalPrev; row.deltaPct=totalPrev>0?(totalCur-totalPrev)/totalPrev*100:null; }
    return row;
  }).filter(function(r){return (r['v'+cur]||0)>0||(r['v'+(prev||cur)]||0)>0;});
  renderPivotTable(cols, rows, fmt(grandCur)+' '+sym+' · '+state.cat);
}

// =======================================================
// VIEW 5: DETALLE SEMANAL
// =======================================================
function renderDetalle() {
  var sym=state.currency.toUpperCase();
  var cols=[
    { field:'year',      headerName:'AÑO',      width:60,  type:'numericColumn', pinned:'left' },
    { field:'week',      headerName:'SEM',       width:55,  type:'numericColumn', pinned:'left', cellRenderer:function(p){return wFmt(p.value);} },
    { field:'categoria', headerName:'CATEGORÍA', width:220, pinned:'left', cellRenderer:catRenderer },
    { field:'usd_total', headerName:'USD',       width:100, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'mxn_total', headerName:'MXN',       width:110, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'date_range',headerName:'PERÍODO',   width:160,
      cellRenderer:function(p){return '<span style="color:#888;font-size:11px">'+(p.value||'')+'</span>';}},
  ];
  RANCH_ORDER.forEach(function(r){
    var c=RANCH_COLORS[r]||'#888';
    cols.push((function(color){return {field:'rn_'+r.replace(/[^a-zA-Z0-9]/g,'_'),headerName:r,width:100,type:'numericColumn',
      cellRenderer:function(p){var v=p.value;if(!v||v<0.01)return '<span class="cell-muted">&#8212;</span>';return '<span style="color:'+color+'">'+fmt(v)+'</span>';}};})(c));
  });
  var rows=[],grandTotal=0;
  DATA.weekly_detail.forEach(function(r){
    if (!state.activeYears[r.year]) return;
    if (r.categoria!==state.cat) return;
    var row={year:r.year,week:r.week,categoria:r.categoria,usd_total:r.usd_total,mxn_total:r.mxn_total,date_range:r.date_range||''};
    RANCH_ORDER.forEach(function(rn){var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;row['rn_'+rn.replace(/[^a-zA-Z0-9]/g,'_')]=src[rn]||0;});
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
      cellRenderer:function(p){var m={'PR':'#16a34a','MP':'#7c3aed','ME':'#0369a1'};return '<span style="color:'+(m[p.value]||'#666')+';font-weight:700">'+(p.value||'')+'</span>';}},
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
  flattenProd(DATA.productos,'PR'); flattenProd(DATA.productos_mp,'MP'); flattenProd(DATA.productos_me,'ME');
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
  var activeRanches=RANCH_ORDER.filter(function(rn){
    return weekKeys.some(function(key){
      return Object.keys(weekMap[key]||{}).some(function(k){return k.endsWith('__r__'+rn)&&weekMap[key][k]>0;});
    });
  });

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

  var thBase='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid var(--pt-hdr-border);white-space:nowrap;';
  var thPin =thBase+'position:sticky;top:0;z-index:4;';
  var thScroll=thBase+'position:sticky;top:0;z-index:3;text-align:right;';

  // ── Header nivel 1: CONCEPTO | [Ranch colspan=nColsPerRanch] ... | TOTAL ──
  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'left:0;min-width:190px">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=RANCH_COLORS[rn]||'#888';
    h1+='<th colspan="'+nColsPerRanch+'" style="'+thScroll+'border-left:2px solid #8EA9C1;text-align:center;color:'+col+'">'+rn+'</th>';
  });
  h1+='<th rowspan="2" style="'+thScroll+'border-left:2px solid #4472C4;min-width:100px;background:#9DC3E6">TOTAL</th>';
  h1+='</tr>';

  // ── Header nivel 2: por cada rancho → [wk labels... | SUB] ──
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
  h2+='</tr>';

  // ── Acumuladores totales ──────────────────────────────
  var grandByRnWk={}, grandByRn={}, grandTotal=0;
  activeRanches.forEach(function(rn){
    grandByRnWk[rn]={}; grandByRn[rn]=0;
    weekKeys.forEach(function(k){ grandByRnWk[rn][k]=0; });
  });

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
    var scByRnWk={}, scByRn={}, scTotal=0;
    activeRanches.forEach(function(rn){
      scByRnWk[rn]={}; scByRn[rn]=0;
      weekKeys.forEach(function(k){
        var v=(weekMap[k]&&weekMap[k][sc+'__r__'+rn])?weekMap[k][sc+'__r__'+rn]:0;
        scByRnWk[rn][k]=v;
        scByRn[rn]+=v;
        scTotal+=v;
        grandByRnWk[rn][k]+=v;
        grandByRn[rn]+=v;
        grandTotal+=v;
      });
    });
    if(scTotal===0) return;

    var tdPin='padding:3px 8px;position:sticky;z-index:1;background:#fff;border-bottom:1px solid #eee;border-right:1px solid #eee;white-space:nowrap;font-size:11px;';
    bodyHtml+='<tr class="pt-row">';
    bodyHtml+='<td style="'+tdPin+'left:0;color:#1e3a5f;font-weight:700">'+sc+'</td>';
    activeRanches.forEach(function(rn){
      weekKeys.forEach(function(key){
        var v=scByRnWk[rn][key];
        if(!v||v===0){bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ddd">—</td>';}
        else{bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#334155;font-weight:600">'+fmt(v)+'</td>';}
      });
      // SUB rancho
      bodyHtml+=cell(scByRn[rn],true,'#1e3a5f');
    });
    // TOTAL fila
    bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;text-align:right;font-weight:700;color:#1e3a5f;border-left:2px solid #4472C4">'+fmt(scTotal)+'</td>';
    bodyHtml+='</tr>';
  });

  // ── Fila total general ────────────────────────────────
  var totStyle='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;';
  var totPin='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;z-index:2;white-space:nowrap;';
  bodyHtml+='<tr>';
  bodyHtml+='<td style="'+totPin+'left:0;text-align:left">TOTAL GENERAL</td>';
  activeRanches.forEach(function(rn){
    weekKeys.forEach(function(key){
      var v=grandByRnWk[rn][key];
      bodyHtml+='<td style="'+totStyle+'color:#1e3a5f">'+(v?fmt(v):'—')+'</td>';
    });
    bodyHtml+='<td style="'+totStyle+'color:#1e3a5f;border-left:1px solid #aaa">'+(grandByRn[rn]?fmt(grandByRn[rn]):'—')+'</td>';
  });
  bodyHtml+='<td style="'+totStyle+'color:#1e3a5f;border-left:2px solid #4472C4">'+(grandTotal?fmt(grandTotal):'—')+'</td>';
  bodyHtml+='</tr>';

  // ── Inyectar en el DOM ────────────────────────────────
  var html='<div class="pt-table-wrap" id="tableWrap" style="overflow:auto"><table class="pt-table" style="border-collapse:collapse;width:100%"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
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
  // Ordenar: primero los del orden preferido, luego cualquier otro
  var activeRanches=MO_RANCH_ORDER.filter(function(rn){ return ranchesEnDatos[rn]; });
  Object.keys(ranchesEnDatos).forEach(function(rn){
    if(activeRanches.indexOf(rn)<0) activeRanches.push(rn);
  });

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
  // NUEVO LAYOUT: GRUPO | CONCEPTO | [Rancho: wk1 wk2 ... SUB] | TOTAL
  var nWeeks = weekKeys.length;
  var nColsPerRanch = nWeeks + 1; // semanas + SUB

  var thBase='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid var(--pt-hdr-border);white-space:nowrap;';
  var thPin=thBase+'position:sticky;top:0;z-index:4;';
  var thScroll=thBase+'position:sticky;top:0;z-index:3;text-align:right;';

  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'left:0;min-width:120px">GRUPO</th>';
  h1+='<th rowspan="2" style="'+thPin+'left:120px;min-width:140px">CONCEPTO</th>';
  activeRanches.forEach(function(rn){
    var col=MO_RANCH_COLORS[rn]||RANCH_COLORS[rn]||'#374151';
    h1+='<th colspan="'+nColsPerRanch+'" style="'+thScroll+'border-left:2px solid #8EA9C1;text-align:center;color:'+col+'">'+rn+'</th>';
  });
  h1+='<th rowspan="2" style="'+thScroll+'border-left:2px solid #4472C4;min-width:100px;background:#9DC3E6">TOTAL</th>';
  h1+='</tr>';

  // HEADER nivel 2: [por cada rancho: wk labels... | SUB]
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
  h2+='</tr>';

  // ── FILAS ─────────────────────────────────────────────
  var grandByRnWk={}, grandByRn={}, grandTotal=0;
  activeRanches.forEach(function(rn){
    grandByRnWk[rn]={}; grandByRn[rn]=0;
    weekKeys.forEach(function(k){ grandByRnWk[rn][k]=0; });
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
    bodyHtml+='<tr style="cursor:pointer;" onclick="togglePtGroup(`mo_grp_'+grpIdx+'`)" title="Hacer clic para expandir o contraer categoría">';
    bodyHtml+='<td style="'+gTdPin+'left:0">'+grp.label+'</td>';
    bodyHtml+='<td style="'+gTdPin+'left:120px"></td>';
    activeRanches.forEach(function(rn){
      weekKeys.forEach(function(key){
        bodyHtml+=cellGrp(grpByRnWk[rn][key]);
      });
      var grpDifStyle='padding:3px 6px;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;background:var(--pt-grp-bg);color:#fff;font-weight:700';
      bodyHtml+='<td style="'+grpDifStyle+'">'+fmtHcDiff(grpHcByRn[rn])+'</td>'; // DIF rancho
    });
    bodyHtml+=cellGrp(grpTotal); // TOTAL fila
    bodyHtml+='</tr>';

    // Filas subcat
    scRows.forEach(function(sc){
      var tdPin='padding:3px 8px;position:sticky;z-index:1;background:#fff;border-bottom:1px solid #eee;border-right:1px solid #eee;white-space:nowrap;';
      bodyHtml+='<tr class="pt-row mo_grp_'+grpIdx+'" style="display:none;" title="Número de personas (Headcount)">';
      bodyHtml+='<td style="'+tdPin+'left:0"></td>';
      bodyHtml+='<td style="'+tdPin+'left:120px;color:#334155;font-size:11px"><span style="color:#888;font-size:9px;margin-right:4px;">👤</span>'+sc.label+'</td>';
      activeRanches.forEach(function(rn){
        weekKeys.forEach(function(key){
          var v=sc.hcByRnWk[rn][key];
          if(!v||v===0){bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ddd">—</td>';}
          else{bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#475569;font-weight:600">'+fmtHc(v)+'</td>';}
        });
        var cellStyle = 'padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#1e3a5f;font-weight:700;';
        bodyHtml+='<td style="'+cellStyle+'">'+(sc.hcByRn[rn]?fmtHcDiff(sc.hcByRn[rn]):'—')+'</td>';
      });
      var cellStyle = 'padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#1e3a5f;font-weight:700;';
      bodyHtml+='<td style="'+cellStyle+'border-left:2px solid #4472C4">'+(sc.hcTotal?fmtHcDiff(sc.hcTotal):'—')+'</td>';
      bodyHtml+='</tr>';
    });
  });

  // Fila total general
  var totStyle='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;';
  var totPin='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;z-index:2;white-space:nowrap;';
  bodyHtml+='<tr>';
  bodyHtml+='<td colspan="2" style="'+totPin+'left:0;text-align:left">TOTAL GENERAL</td>';
  activeRanches.forEach(function(rn){
    weekKeys.forEach(function(key){
      var v=grandByRnWk[rn][key];
      bodyHtml+='<td style="'+totStyle+'color:#1e3a5f">'+( v?fmt(v):'—')+'</td>';
    });
    bodyHtml+='<td style="'+totStyle+'color:#1e3a5f;border-left:1px solid #aaa">'+(grandByRn[rn]?fmt(grandByRn[rn]):'—')+'</td>';
  });
  bodyHtml+='<td style="'+totStyle+'color:#1e3a5f;border-left:2px solid #4472C4">'+(grandTotal?fmt(grandTotal):'—')+'</td>';
  bodyHtml+='</tr>';

  // ── Inyectar en el DOM ────────────────────────────────
  var html='<div class="pt-table-wrap" id="tableWrap" style="overflow:auto"><table class="pt-table" style="border-collapse:collapse;width:100%"><thead>'+h1+h2+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
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

function showProdPanel(rowData, opts) {
  opts=opts||{};
  var cat=rowData._cat, yr=rowData._year, wn=rowData._week;
  var fromW=rowData._fromWeek||wn, toW=rowData._toWeek||wn;
  var ranchFilter=opts.ranch||null;
  if (!cat||!yr) return;

  var isMant=cat==='MANTENIMIENTO', isMatEmp=cat==='MATERIAL DE EMPAQUE';
  var isMirfe=cat===CAT_MIRFE, isMipe=cat===CAT_MIPE;
  var src=isMant?'mp':(isMatEmp?'me':'pr');
  var tipoFilter=null;
  if (src==='pr'){ if(isMirfe)tipoFilter='MIRFE'; else if(isMipe)tipoFilter='MIPE'; }
  var dsMap={pr:DATA.productos,mp:DATA.productos_mp,me:DATA.productos_me};
  var ds=dsMap[src]||{};

  var wkStart=parseInt(fromW||wn||0), wkEnd=parseInt(toW||wn||0);
  if (!wkStart||!wkEnd) return;
  if (wkStart>wkEnd){var tmp=wkStart;wkStart=wkEnd;wkEnd=tmp;}

  var rows=[];
  for (var wk=wkStart;wk<=wkEnd;wk++){
    var wkCodeShort=((yr%100)*100)+wk, wkCodeLong=(yr*100)+wk;
    var weekD=ds[wkCodeShort]||ds[String(wkCodeShort)]||ds[wkCodeLong]||ds[String(wkCodeLong)];
    if (!weekD) continue;
    Object.keys(weekD).forEach(function(ranch){
      if (ranchFilter&&ranch!==ranchFilter) return;
      var byTipo=weekD[ranch];
      Object.keys(byTipo).forEach(function(tipo){
        if (tipoFilter&&tipo!==tipoFilter) return;
        (byTipo[tipo]||[]).forEach(function(item){
          rows.push({week_code:wkCodeShort,rancho:ranch,tipo:tipo,producto:item[0]||'',unidades:item[1]||'',gasto:parseFloat(item[2])||0,ubicacion:item[3]||''});
        });
      });
    });
  }

  var rangeText=wkStart===wkEnd?(wFmt(wkStart)+' · '+yr):(wFmt(wkStart)+'→'+wFmt(wkEnd)+' · '+yr);
  var panelTitle = cat+' &#9656; '+rangeText+(ranchFilter?' · '+ranchFilter:'');
  
  var panelHtml = '';

  // ── KPI de siembra ────────────────────────────────────────────────
  var kpiSection = '';
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
    {k:'charolas',     lbl:'N\u00ba CHAROLAS'},
    {k:'esquejes',     lbl:'N\u00ba ESQUEJES'},
    {k:'metros',       lbl:'METROS SIEMBRA'},
    {k:'hectareas',    lbl:'HECT\u00c1REAS'},
  ];
  if (DATA.siembra_data) {
    var _wkKey = ((yr%100)*100) + wkStart;
    var _wkSrc = DATA.siembra_data[_wkKey] || DATA.siembra_data[String(_wkKey)] || null;
    if (_wkSrc) {
      var _sRow = ranchFilter ? (_wkSrc[ranchFilter] || _wkSrc['TOTAL'] || {}) : (_wkSrc['TOTAL'] || {});
      var _activos = _allMetas.filter(function(m){ var v=_sRow[m.k]; return v!==undefined&&v!==null&&v!==''&&v!==0; });
      if (_activos.length > 0) {
        kpiSection = '<div style="flex-shrink:0;background:#EBF3FB;border-bottom:1px solid #8EA9C1;padding:3px 8px;display:flex;align-items:center;overflow-x:auto;scrollbar-width:none;">';
        _activos.forEach(function(m,i){
          if (i>0) kpiSection += '<div style="width:1px;background:#8EA9C1;height:16px;margin:0 10px;flex-shrink:0;"></div>';
          kpiSection += '<div style="display:flex;align-items:baseline;gap:4px;flex-shrink:0;white-space:nowrap;">'+
            '<span style="font-size:9px;color:#44546A;text-transform:uppercase;">'+m.lbl+':</span>'+
            '<span style="font-size:12px;font-weight:700;color:#1e3a5f;">'+Number(_sRow[m.k]).toLocaleString('es-MX',{maximumFractionDigits:2})+'</span>'+
            '</div>';
        });
        kpiSection += '</div>';
      }
    }
  }

  // ── Zona 2: Tabla de productos ────────────────────────────────────
  var productSection = '';
  if (rows.length === 0) {
    productSection = '<div style="padding:12px 10px; color:#94a3b8; font-size:11px; text-align:center;">Sin registros de producto para este período.</div>';
  } else {
    rows.sort(function(a,b){return b.gasto-a.gasto;});
    var total=rows.reduce(function(s,r){return s+r.gasto;},0);
    var panelMeta = 'Reg: <b>'+rows.length+'</b> &nbsp;|&nbsp; Gasto: <b style="color:#16a34a">'+fmt(total)+'</b>';
    productSection =
      '<div style="flex-shrink:0; background:#f1f5f9; border-bottom:1px solid #e2e8f0; padding:4px 8px; display:flex; justify-content:space-between; align-items:center;">' +
        '<span style="font-size:9px; font-weight:700; color:#94a3b8; letter-spacing:1px; text-transform:uppercase;">DETALLE DE PRODUCTOS</span>' +
        '<span style="font-size:10px; color:#475569;">'+panelMeta+'</span>' +
      '</div>' +
      '<div style="overflow-x:auto; overflow-y:auto; flex:1; scrollbar-width:thin;">' +
        '<table style="font-size:10px; width:100%; border-collapse:collapse;">' +
          '<thead><tr style="position:sticky;top:0;z-index:1;">' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600; white-space:nowrap;">UBICACI\u00d3N</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">PRODUCTO</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">UNID.</th>' +
            '<th style="text-align:right; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">GASTO</th>' +
          '</tr></thead><tbody>';
    rows.forEach(function(r,i){
      var rowBg = (i%2===0)?'#ffffff':'#f8fafc';
      productSection += '<tr style="background:'+rowBg+'; border-bottom:1px solid #f1f5f9;">' +
        '<td style="padding:3px 6px; white-space:nowrap; font-weight:600; color:#0f172a;">'+r.ubicacion+'</td>' +
        '<td style="padding:3px 6px; color:#0f172a;">'+r.producto+'</td>' +
        '<td style="padding:3px 6px; color:#94a3b8; font-size:9px;">'+r.unidades+'</td>' +
        '<td style="padding:3px 6px; text-align:right; font-weight:700; color:#0f172a;">'+fmt(r.gasto)+'</td>' +
        '</tr>';
    });
    productSection += '</tbody></table></div>';
  }

  panelHtml =
    '<div style="flex:1; min-width:340px; border:1px solid #cbd5e1; border-top:3px solid #4472C4; display:flex; flex-direction:column; background:#fff; overflow:hidden;">' +
      '<div style="background:#4472C4; color:#fff; padding:5px 10px; flex-shrink:0; display:flex; justify-content:space-between; align-items:center;">' +
        '<div style="font-weight:700; font-size:11px; text-transform:uppercase; letter-spacing:0.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="'+panelTitle+'">'+panelTitle+'</div>' +
      '</div>' +
      kpiSection +
      productSection +
    '</div>';
  
  if (_prodViews.indexOf(panelHtml) === -1) {
    _prodViews.push(panelHtml);
    if (_prodViews.length > 2) {
      _prodViews.shift(); // keep max 2 side-by-side
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
  var h = window.innerHeight || document.documentElement.clientHeight || 700;
  window.parent.postMessage({type:'streamlit:setFrameHeight',height:Math.max(h,700)},'*');
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

# Renderizamos el iframe PRIMERO
components.html(html_final, height=900, scrolling=False)

# ─── POPUP EXCEL / SHAREPOINT ────────────────────────────────────────────────
available_weeks = sorted(
    {str(r["year"] % 100).zfill(2) + str(r["week"]).zfill(2) for r in DATA.get("weekly_detail", [])},
    reverse=True
)

if available_weeks:
    from data_extractor import get_sheet_xlsx
    try:
        from data_extractor import crear_hoja_wk
        _crear_disponible = True
    except ImportError:
        _crear_disponible = False

    with st.expander("⚙ EXCEL / SHAREPOINT"):
        st.markdown("<p style='font-size:12px; font-weight:bold; color:#1e3a5f; margin-bottom:5px;'>⬇ Descargar Archivo WK</p>", unsafe_allow_html=True)
        selected_wk = st.selectbox(
            "Semana a descargar",
            options=available_weeks,
            format_func=lambda c: f"WK{c}",
            label_visibility="collapsed"
        )
        if st.button("Preparar XLSX", key="dl_xlsx", use_container_width=True):
            with st.spinner(f"Preparando WK{selected_wk}..."):
                xlsx_bytes = get_sheet_xlsx(selected_wk)
            if xlsx_bytes:
                st.download_button(
                    label=f"💾 Confirmar WK{selected_wk}",
                    data=xlsx_bytes,
                    file_name=f"WK{selected_wk}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_btn",
                    use_container_width=True
                )
            else:
                st.error(f"No se encontró WK{selected_wk}.")

        if _crear_disponible:
            st.divider()
            st.markdown("<p style='font-size:12px; font-weight:bold; color:#1e3a5f; margin-bottom:5px;'>✚ Nueva hoja SharePoint</p>", unsafe_allow_html=True)
            nuevo_nombre = st.text_input(
                "Nombre (Ej: WK2518)",
                key="nuevo_wk_nombre",
                placeholder="Ej: WK2518",
                label_visibility="collapsed"
            ).strip().upper()

            if st.button("Crear Hoja", key="btn_crear_hoja", type="primary", use_container_width=True):
                if not nuevo_nombre:
                    st.warning("⚠️ Escribe el nombre de la hoja.")
                elif not nuevo_nombre.startswith("WK") or len(nuevo_nombre) != 6:
                    st.warning("⚠️ El formato debe ser WK####.")
                else:
                    try:
                        tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                        client_id     = st.secrets["sharepoint"]["client_id"]
                        client_secret = st.secrets["sharepoint"]["client_secret"]
                        with st.spinner(f"Creando {nuevo_nombre}…"):
                            resultado = crear_hoja_wk(nuevo_nombre, tenant_id, client_id, client_secret)
                        if resultado.get("ok"):
                            st.success(resultado["mensaje"])
                            st.cache_data.clear()
                        else:
                            st.error(f"❌ {resultado['error']}")
                    except KeyError as e:
                        st.error(f"❌ Falta credencial {e}.")
