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
  --pt-grp-bg:      #4472C4;   /* fila de grupo/año   */
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
    <span class="tb-label">Semana</span>
    <div class="week-ctr">
      <button class="tb-btn" onclick="prevWeek()">&#9664;</button>
      <span id="weekLabel">&#8212;</span>
      <button class="tb-btn" onclick="nextWeek()">&#9654;</button>
    </div>
    <input type="range" class="tb-slider" id="weekSlider" min="1" max="52" value="1" oninput="onWeekSlider(this.value)">
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

  <!-- RANGE BAR (solo comparativo) -->
  <div class="range-bar" id="rangeBar">
    <span class="tb-label">Desde</span>
    <span class="range-val" id="fromWeekLabel">W01</span>
    <input type="range" class="tb-slider" id="fromSlider" min="1" max="52" value="1" oninput="onRangeChange()">
    <span style="color:#aaa;font-size:11px">→</span>
    <span class="tb-label">Hasta</span>
    <span class="range-val" id="toWeekLabel">W52</span>
    <input type="range" class="tb-slider" id="toSlider" min="1" max="52" value="52" oninput="onRangeChange()">
    <span class="range-badge" id="rangeBadge">W01 → W52</span>

  </div>

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
  var col = RANCH_COLORS[ranch]||'#888';
  return function(p) {
    var v = p.value;
    if (!v||isNaN(v)||v===0) return '';
    return '<span style="color:'+col+';font-weight:600">'+fmt(v)+'</span>';
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

  state.activeYears = {};
  var latestYr = DATA.years[DATA.years.length-1];
  var prevYr   = DATA.years[DATA.years.length-2];
  if (latestYr) state.activeYears[latestYr] = true;
  if (prevYr)   state.activeYears[prevYr]   = true;

  var wSet = {};
  DATA.weekly_detail.forEach(function(r){ wSet[r.week]=1; });
  allWeeks = Object.keys(wSet).map(Number).sort(function(a,b){return a-b;});

  var wksLatest = DATA.weekly_detail
    .filter(function(r){return r.year===latestYr;})
    .map(function(r){return r.week;})
    .filter(function(v,i,a){return a.indexOf(v)===i;})
    .sort(function(a,b){return a-b;});
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
  var wn = allWeeks[state.weekIdx]||1;
  var sl = document.getElementById('weekSlider');
  sl.min=allWeeks[0]||1; sl.max=allWeeks[allWeeks.length-1]||52; sl.value=wn;
  var activeYrs = getActiveYears();
  var yr = activeYrs[activeYrs.length-1]||DATA.years[DATA.years.length-1];
  document.getElementById('weekLabel').textContent = String(yr).slice(2)+String(wn).padStart(2,'0');
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
  if (rb) rb.className='range-bar'+((v==='comparativo'||v==='servicios')?' show':'');
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

  var head='<tr><th>Semana</th><th>Fecha</th><th>Total '+sym+'</th><th>Δ$ vs sem ant.</th>'+ranchCols.map(function(r){return '<th>'+r+'</th>';}).join('')+'</tr>';
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
        var style='color:'+(v>0?(RANCH_COLORS[r]||'#888'):'#ddd')+(v>0?';cursor:pointer':'');
        var attrs=v>0?' class="cmp-clickable" data-yr="'+yr+'" data-wk="'+w+'" data-ranch="'+r+'"':'';
        return '<td style="'+style+'"'+attrs+'>'+(v>0?fmt(v):'')+'</td>';
      }).join('');
      var totalStyle='color:'+(val>0?col:'#bbb')+';font-weight:'+(val>0?'600':'400')+(val>0?';cursor:pointer':'');
      var totalAttrs=val>0?' class="cmp-clickable" data-yr="'+yr+'" data-wk="'+w+'" data-ranch=""':'';
      return '<tr class="cmp-row">'+
        '<td style="color:'+col+';font-weight:600">'+String(yr).slice(2)+String(w).padStart(2,'0')+'</td>'+
        '<td style="color:#777;font-size:11px">'+fmtMes(d&&d.date_range)+'</td>'+
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
// VIEW 7: COSTO SERVICIOS
// =======================================================
var SV_SUBCATS=['Electricidad','Fletes y Acarreos','Gastos de Exportación','Certificado Fitosanitario','Transporte de Personal','Compra de Flor a Terceros','Comida para el Personal','RO, TEL, RTA.Alim'];
function renderServicios() {
  var sym=state.currency.toUpperCase();
  var f=state.fromWeek, t=state.toWeek;
  var yrs=getActiveYears();
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});

  // Recopilar subcats con datos
  var subcatsSet={};
  var weekMap={};  // "yr-wk" -> { subcat -> val, _total, date_range }

  function addRecord(r) {
    if (!state.activeYears[r.year]) return;
    if (r.week<f||r.week>t) return;
    var key=r.year+'-'+r.week;
    if (!weekMap[key]) weekMap[key]={_year:r.year,_week:r.week,_total:0,date_range:r.date_range||''};
    var subcat, val;
    if (r.subcat) {
      subcat=r.subcat;
      val=state.currency==='usd'?r.usd_total:r.mxn_total;
    } else if (r.categoria&&r.categoria.startsWith('SV:')) {
      subcat=r.categoria.replace('SV:','');
      val=state.currency==='usd'?r.usd_total:r.mxn_total;
    } else return;
    if (!subcat) return;
    subcatsSet[subcat]=1;
    weekMap[key][subcat]=(weekMap[key][subcat]||0)+(val||0);
    weekMap[key]._total=(weekMap[key]._total||0)+(val||0);
  }

  var src=Array.isArray(DATA.servicios_data)&&DATA.servicios_data.length ? DATA.servicios_data : DATA.weekly_detail;
  src.forEach(addRecord);

  // Columnas: Semana | Fecha | Total | Δ$ | [subcats]
  var orderedSubcats=SV_SUBCATS.filter(function(sc){return subcatsSet[sc];});
  Object.keys(subcatsSet).forEach(function(sc){if(orderedSubcats.indexOf(sc)===-1)orderedSubcats.push(sc);});

  var cols=[
    {field:'semana',     headerName:'SEMANA',   width:80,  pinned:'left',
     cellRenderer:function(p){var c=YEAR_COLORS[p.data._year]||'#888';return '<span style="color:'+c+';font-weight:700">'+p.value+'</span>';}},
    {field:'fecha',      headerName:'FECHA',    width:120, pinned:'left',
     cellRenderer:function(p){return '<span style="color:#777;font-size:11px">'+p.value+'</span>';}},
    {field:'total',      headerName:'TOTAL '+sym, width:110, type:'numericColumn', cellRenderer:moneyRenderer},
    {field:'deltaAmt',   headerName:'Δ $',       width:90,  type:'numericColumn', cellRenderer:deltaAmtRenderer},
  ];
  orderedSubcats.forEach(function(sc){
    cols.push({field:'sc_'+sc.replace(/[^a-zA-Z0-9]/g,'_'), headerName:sc, width:130, type:'numericColumn', cellRenderer:moneyRenderer});
  });

  // Filas ordenadas por año luego semana
  var rows=[]; var prevVal=null; var grandTotal=0;
  yrs.forEach(function(yr){
    prevVal=null;
    rangeWeeks.forEach(function(w){
      var key=yr+'-'+w;
      var d=weekMap[key];
      var total=d?d._total:0;
      var row={semana:String(yr).slice(2)+String(w).padStart(2,'0'), fecha:d?fmtMes(d.date_range):'', total:total, _year:yr, _week:w};
      row.deltaAmt=(prevVal!==null&&total>0)?total-prevVal:null;
      orderedSubcats.forEach(function(sc){row['sc_'+sc.replace(/[^a-zA-Z0-9]/g,'_')]=d?d[sc]||0:0;});
      if (total>0) { rows.push(row); prevVal=total; grandTotal+=total; }
    });
  });

  renderPivotTable(cols, rows, fmt(grandTotal)+' '+sym);
}

// =======================================================
// VIEW 8: COSTO MANO DE OBRA
// =======================================================
var MO_SUBCATS = [
  'Nómina Admon','H.Extra Dom. y Festivos (Admon)','Bonos Asist./Puntualidad (Admon)',
  'Nómina Producción','H.Extra Dom. y Fest. (Prod.)','Bonos Asist./Puntualidad (Prod.)',
  'Nómina Prod. Corte','H.Extra Corte','Bonos Corte',
  'Nómina Prod. Transplante','H.Extra Transplante','Bonos Transplante',
  'Nómina Prod. Manejo Planta','H.Extra Manejo Planta','Bonos Manejo Planta',
  'Nómina HOOPS','H.Extra HOOPS','Bonos HOOPS',
  'Nómina MIPE/MIRFE','H.Extra MIPE/MIRFE','Bonos MIPE/MIRFE',
  'Nómina Op. Tractores/Cameros','H.Extra Tractores/Cameros','Bonos Tractores/Cameros',
  'Nómina Op. Chofer','H.Extra Chofer','Bonos Chofer',
  'Nómina Op. Veladores','H.Extra Veladores','Bonos Veladores',
  'Nómina Op. Soldador','H.Extra Soldador','Bonos Soldador',
  'Nómina Prod. Contratista','IMSS/INFONAVIT RCV','1.8% Estado'
];
function renderManoObra() {
  var sym=state.currency.toUpperCase();
  var f=state.fromWeek, t=state.toWeek;
  var yrs=getActiveYears();
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});

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
    weekMap[key][subcat]=(weekMap[key][subcat]||0)+(val||0);
    Object.keys(ranches||{}).forEach(function(rn){
      weekMap[key][subcat+'__r__'+rn]=(weekMap[key][subcat+'__r__'+rn]||0)+(ranches[rn]||0);
    });
  });

  var MO_GROUPS=[
    {label:'ADMINISTRACIÓN',subcats:['Nómina Admon','H.Extra Dom. y Festivos (Admon)','Bonos Asist./Puntualidad (Admon)']},
    {label:'PRODUCCIÓN',subcats:['Nómina Producción','H.Extra Dom. y Fest. (Prod.)','Bonos Asist./Puntualidad (Prod.)']},
    {label:'CORTE',subcats:['Nómina Prod. Corte','H.Extra Corte','Bonos Corte']},
    {label:'TRANSPLANTE',subcats:['Nómina Prod. Transplante','H.Extra Transplante','Bonos Transplante']},
    {label:'MANEJO PLANTA',subcats:['Nómina Prod. Manejo Planta','H.Extra Manejo Planta','Bonos Manejo Planta']},
    {label:'HOOPS',subcats:['Nómina HOOPS','H.Extra HOOPS','Bonos HOOPS']},
    {label:'MIPE / MIRFE',subcats:['Nómina MIPE/MIRFE','H.Extra MIPE/MIRFE','Bonos MIPE/MIRFE']},
    {label:'TRACTORES',subcats:['Nómina Op. Tractores/Cameros','H.Extra Tractores/Cameros','Bonos Tractores/Cameros']},
    {label:'CHOFER',subcats:['Nómina Op. Chofer','H.Extra Chofer','Bonos Chofer']},
    {label:'VELADORES',subcats:['Nómina Op. Veladores','H.Extra Veladores','Bonos Veladores']},
    {label:'SOLDADOR',subcats:['Nómina Op. Soldador','H.Extra Soldador','Bonos Soldador']},
    {label:'OTROS',subcats:['Nómina Prod. Contratista','IMSS/INFONAVIT RCV','1.8% Estado']}
  ];

  // Semanas con datos
  var weekKeys=[];
  yrs.forEach(function(yr){
    rangeWeeks.forEach(function(w){
      var key=yr+'-'+w;
      var hasSome=false;
      Object.keys(weekMap[key]||{}).forEach(function(k){if(k[0]!=='_'&&!k.includes('__r__')&&weekMap[key][k]>0)hasSome=true;});
      if(hasSome) weekKeys.push(key);
    });
  });

  // Ranchos con datos en este rango
  var activeRanches=RANCH_ORDER.filter(function(rn){
    return weekKeys.some(function(key){
      return Object.keys(weekMap[key]||{}).some(function(k){return k.endsWith('__r__'+rn)&&weekMap[key][k]>0;});
    });
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

  // ── Construir HTML ────────────────────────────────────
  // HEADER nivel 1: GRUPO | CONCEPTO | [WKxxxx colspan=R+1] ... | TOTAL
  var ncols=activeRanches.length+1; // ranchos + subtotal por semana
  var thBase='padding:5px 8px;background:var(--pt-hdr-bg);color:#1e3a5f;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid var(--pt-hdr-border);border-right:1px solid var(--pt-hdr-border);white-space:nowrap;';
  var thPin=thBase+'position:sticky;top:0;z-index:4;';
  var thScroll=thBase+'position:sticky;top:0;z-index:3;text-align:right;';

  var h1='<tr>';
  h1+='<th rowspan="2" style="'+thPin+'left:0;min-width:120px">GRUPO</th>';
  h1+='<th rowspan="2" style="'+thPin+'left:120px;min-width:140px">CONCEPTO</th>';
  weekKeys.forEach(function(key){
    var d=weekMap[key];
    var lbl=String(d._year).slice(2)+String(d._week).padStart(2,'0');
    var borderL='border-left:2px solid #8EA9C1;';
    h1+='<th colspan="'+ncols+'" style="'+thScroll+borderL+'text-align:center">WK '+lbl+'</th>';
  });
  h1+='<th rowspan="2" style="'+thScroll+'border-left:2px solid #4472C4;min-width:100px;background:#9DC3E6">TOTAL</th>';
  h1+='</tr>';

  // HEADER nivel 2: [por cada semana: ranchos... | SUB]
  var h2='<tr>';
  weekKeys.forEach(function(key){
    activeRanches.forEach(function(rn){
      var col=RANCH_COLORS[rn]||'#555';
      h2+='<th style="'+thScroll+'border-left:1px solid var(--pt-hdr-border);font-size:9px;color:'+col+';min-width:80px">'+rn+'</th>';
    });
    h2+='<th style="'+thScroll+'border-left:1px solid #aaa;min-width:88px;background:#BDD7EE">SUB</th>';
  });
  h2+='</tr>';

  // ── FILAS ─────────────────────────────────────────────
  var grandByWk={}; var grandByRn={}; var grandTotal=0;
  weekKeys.forEach(function(k){grandByWk[k]=0;});
  activeRanches.forEach(function(rn){grandByRn[rn]={}; weekKeys.forEach(function(k){grandByRn[rn][k]=0;});});

  var bodyHtml='';

  MO_GROUPS.forEach(function(grp){
    var grpByWk={}; weekKeys.forEach(function(k){grpByWk[k]=0;});
    var grpByRnWk={}; activeRanches.forEach(function(rn){grpByRnWk[rn]={}; weekKeys.forEach(function(k){grpByRnWk[rn][k]=0;});});
    var grpTotal=0;
    var scRows=[];

    grp.subcats.forEach(function(sc){
      var scByWk={}; var scByRnWk={}; var scTotal=0;
      weekKeys.forEach(function(k){scByWk[k]=0;});
      activeRanches.forEach(function(rn){scByRnWk[rn]={}; weekKeys.forEach(function(k){scByRnWk[rn][k]=0;});});
      weekKeys.forEach(function(key){
        var val=(weekMap[key]&&weekMap[key][sc])?weekMap[key][sc]:0;
        scByWk[key]=val; scTotal+=val; grpByWk[key]+=val; grandByWk[key]+=val;
        activeRanches.forEach(function(rn){
          var rv=(weekMap[key]&&weekMap[key][sc+'__r__'+rn])?weekMap[key][sc+'__r__'+rn]:0;
          scByRnWk[rn][key]=rv; grpByRnWk[rn][key]+=rv; grandByRn[rn][key]+=rv;
        });
      });
      grpTotal+=scTotal; grandTotal+=scTotal;
      if(scTotal>0) scRows.push({label:shortLabel(sc),byWk:scByWk,byRnWk:scByRnWk,total:scTotal});
    });
    if(grpTotal===0) return;

    // Fila grupo
    var gTdPin='padding:4px 8px;background:var(--pt-grp-bg);color:#fff;font-weight:700;font-size:11px;position:sticky;z-index:2;border-bottom:1px solid #ddd;border-right:1px solid rgba(255,255,255,0.2);white-space:nowrap;';
    bodyHtml+='<tr>';
    bodyHtml+='<td style="'+gTdPin+'left:0">'+grp.label+'</td>';
    bodyHtml+='<td style="'+gTdPin+'left:120px"></td>';
    weekKeys.forEach(function(key){
      activeRanches.forEach(function(rn){bodyHtml+=cellGrp(grpByRnWk[rn][key]);});
      bodyHtml+=cellGrp(grpByWk[key]);
    });
    bodyHtml+=cellGrp(grpTotal);
    bodyHtml+='</tr>';

    // Filas subcat
    scRows.forEach(function(sc){
      var tdPin='padding:3px 8px;position:sticky;z-index:1;background:#fff;border-bottom:1px solid #eee;border-right:1px solid #eee;white-space:nowrap;';
      bodyHtml+='<tr class="pt-row">';
      bodyHtml+='<td style="'+tdPin+'left:0"></td>';
      bodyHtml+='<td style="'+tdPin+'left:120px;color:#334155;font-size:11px">'+sc.label+'</td>';
      weekKeys.forEach(function(key){
        activeRanches.forEach(function(rn){
          var v=sc.byRnWk[rn][key];
          var col=RANCH_COLORS[rn]||'#555';
          if(!v||v===0){bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:#ddd">—</td>';}
          else{bodyHtml+='<td style="padding:3px 6px;border-bottom:1px solid #eee;border-right:1px solid #eee;text-align:right;color:'+col+';font-weight:600">'+fmt(v)+'</td>';}
        });
        bodyHtml+=cell(sc.byWk[key],false,'#1e3a5f');
      });
      bodyHtml+=cell(sc.total,true,'#1e3a5f');
      bodyHtml+='</tr>';
    });
  });

  // Fila total general
  var totStyle='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;text-align:right;';
  var totPin='padding:4px 8px;background:var(--pt-tot-bg);font-weight:700;border-bottom:1px solid #ddd;border-right:1px solid #ccc;position:sticky;z-index:2;white-space:nowrap;';
  bodyHtml+='<tr>';
  bodyHtml+='<td colspan="2" style="'+totPin+'left:0;text-align:left">TOTAL GENERAL</td>';
  weekKeys.forEach(function(key){
    activeRanches.forEach(function(rn){
      var v=grandByRn[rn][key]; var col=RANCH_COLORS[rn]||'#555';
      bodyHtml+='<td style="'+totStyle+'color:'+col+'">'+( v?fmt(v):'—')+'</td>';
    });
    bodyHtml+='<td style="'+totStyle+'color:#1e3a5f">'+(grandByWk[key]?fmt(grandByWk[key]):'—')+'</td>';
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

  // ── Barra de métricas de siembra (siempre visible) ───────────────
  var siembraBar = '';
  if (DATA.siembra_data) {
    var wkCodeShort = ((yr%100)*100) + wkStart;
    var wkSrc = DATA.siembra_data[wkCodeShort] || DATA.siembra_data[String(wkCodeShort)] || null;
    if (wkSrc) {
      var sRow = ranchFilter ? (wkSrc[ranchFilter] || wkSrc['TOTAL'] || {}) : (wkSrc['TOTAL'] || {});
      var sMetas = [
        {k:'charolas', lbl:'N\u00ba CHAROLAS SEMBRADAS'},
        {k:'esquejes', lbl:'N\u00ba ESQUEJES SEMBRADOS'},
        {k:'metros',   lbl:'METROS DE SIEMBRA'},
        {k:'hectareas',lbl:'HECT\u00c1REAS EN SIEMBRA'},
      ];
      siembraBar = '<div style="display:flex;gap:6px;padding:4px 6px;background:#f0fdf4;border-bottom:1px solid #bbf7d0;flex-shrink:0;flex-wrap:wrap;">';
      sMetas.forEach(function(m){
        var v = (sRow[m.k]!==undefined && sRow[m.k]!=='') ? Number(sRow[m.k]).toLocaleString('es-MX',{maximumFractionDigits:2}) : '\u2014';
        siembraBar += '<div style="flex:1;min-width:110px;text-align:center;padding:3px 6px;background:#fff;border:1px solid #bbf7d0;border-radius:3px;">' +
          '<div style="font-size:8px;color:#16a34a;text-transform:uppercase;letter-spacing:0.3px;white-space:nowrap;">' + m.lbl + '</div>' +
          '<div style="font-size:13px;font-weight:700;color:#0f172a;">' + v + '</div>' +
          '</div>';
      });
      siembraBar += '</div>';
    }
  }

  // ── Zona 1: KPIs de siembra (tarjetas prominentes) ───────────────
  var kpiSection = '';
  var sMetas = [
    {k:'charolas', lbl:'CHAROLAS SEMBRADAS', icon:'🌱'},
    {k:'esquejes', lbl:'ESQUEJES SEMBRADOS',  icon:'🌿'},
    {k:'metros',   lbl:'METROS DE SIEMBRA',   icon:'📐'},
    {k:'hectareas',lbl:'HECT\u00c1REAS EN SIEMBRA', icon:'🗺'},
  ];
  if (siembraBar !== '') {
    // siembraBar tiene wkSrc/sRow ya calculados — recalcular para diseño nuevo
    var wkCodeShort2 = ((yr%100)*100) + wkStart;
    var wkSrc2 = (DATA.siembra_data||{})[wkCodeShort2] || (DATA.siembra_data||{})[String(wkCodeShort2)] || null;
    var sRow2 = wkSrc2 ? (ranchFilter ? (wkSrc2[ranchFilter] || wkSrc2['TOTAL'] || {}) : (wkSrc2['TOTAL'] || {})) : {};
    kpiSection =
      '<div style="flex-shrink:0; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:8px 10px;">' +
        '<div style="font-size:9px; font-weight:700; color:#94a3b8; letter-spacing:1px; text-transform:uppercase; margin-bottom:6px;">INDICADORES DE SIEMBRA</div>' +
        '<div style="display:grid; grid-template-columns:repeat(4,1fr); gap:6px;">';
    sMetas.forEach(function(m){
      var raw = sRow2[m.k];
      var v = (raw !== undefined && raw !== '' && raw !== 0) ? Number(raw).toLocaleString('es-MX',{maximumFractionDigits:2}) : '\u2014';
      var hasData = (raw !== undefined && raw !== '' && raw !== 0);
      kpiSection +=
        '<div style="background:#fff; border:1px solid '+(hasData?'#a8bedf':'#e2e8f0')+'; border-top:3px solid '+(hasData?'#4472C4':'#cbd5e1')+'; border-radius:4px; padding:6px 8px; text-align:center;">' +
          '<div style="font-size:16px; line-height:1; margin-bottom:3px;">'+m.icon+'</div>' +
          '<div style="font-size:16px; font-weight:800; color:'+(hasData?'#0f172a':'#cbd5e1')+'; line-height:1.1;">'+v+'</div>' +
          '<div style="font-size:8px; color:#64748b; text-transform:uppercase; letter-spacing:0.4px; margin-top:3px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">'+m.lbl+'</div>' +
        '</div>';
    });
    kpiSection += '</div></div>';
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
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600; white-space:nowrap;">WK</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600; white-space:nowrap;">UBICACI\u00d3N</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">PRODUCTO</th>' +
            '<th style="text-align:left; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">UNID.</th>' +
            '<th style="text-align:right; background:#f8fafc; border-bottom:2px solid #e2e8f0; padding:4px 6px; color:#64748b; font-weight:600;">GASTO</th>' +
          '</tr></thead><tbody>';
    rows.forEach(function(r,i){
      var rowBg = (i%2===0)?'#ffffff':'#f8fafc';
      productSection += '<tr style="background:'+rowBg+'; border-bottom:1px solid #f1f5f9;">' +
        '<td style="padding:3px 6px; color:#94a3b8; white-space:nowrap;">'+r.week_code+'</td>' +
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

# ─── POPUP EXCEL / SHAREPOINT ────────────────────────────────────────────────
# Lo renderizamos ANTES del iframe usando float para que coexista en el scroll
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

    st.markdown("""
    <style>
    /* El contenedor principal del popover que flota sobre la app */
    div[data-testid="stPopover"] {
        position: fixed !important;
        top: 6px !important;
        right: 125px !important; /* Ajustado justo a la izquierda del CSV */
        z-index: 999999 !important;
        width: 75px !important; /* Idéntico ancho aproximado al botón CSV */
        height: 24px !important; max-height: 24px !important;
        margin: 0 !important; padding: 0 !important;
    }
    /* Estilizar SOLO el botón principal que abre el panel para que luzca idéntico al .hdr-btn */
    div[data-testid="stPopover"] button {
        width: 100% !important;
        padding: 0px 4px !important; font-size: 10px !important; font-weight: 700 !important; color: #ffffff !important;
        background: rgba(255,255,255,0.35) !important; border: 1px solid rgba(255,255,255,0.35) !important;
        border-radius: 3px !important; height: 24px !important; min-height: 24px !important; max-height: 24px !important;
        display: flex !important; align-items: center !important; justify-content: center !important; cursor: pointer !important;
        margin: 0 !important;
    }
    /* Forzar que el texto de adentro del Popover no arrastre márgenes */
    div[data-testid="stPopover"] button * {
        font-size: 10px !important; margin: 0 !important; padding: 0 !important; line-height: 1 !important; color: #ffffff !important;
        min-height: 0 !important; max-height: 24px !important;
    }
    div[data-testid="stPopover"] button p {
        font-size: 10px !important; margin: 0 !important; padding: 0 !important; line-height: 1 !important; color: #ffffff !important;
    }
    div[data-testid="stPopover"] button:hover {
        background: rgba(255,255,255,0.55) !important; border-color: rgba(255,255,255,0.55) !important;
    }
    div[data-testid="stPopoverBody"] {
        width: 250px !important;
        padding: 10px 15px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.popover("⚙ EXCEL", use_container_width=True):
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

# Renderizamos el iframe DESPUÉS
components.html(html_final, height=900, scrolling=False)
