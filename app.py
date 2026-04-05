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

from data_extractor import get_datos, SHAREPOINT_URL_WK

# ── Acción "Crear Hoja" via query param (comunicación iframe → Streamlit) ──────
_qp = st.query_params
if _qp.get("_action") == "crear_hoja":
    _nombre = _qp.get("_nombre", "").strip().upper()
    st.query_params.clear()           # limpiar antes de procesar
    if _nombre.startswith("WK") and len(_nombre) == 6:
        try:
            from data_extractor import crear_hoja_wk
            _tid = st.secrets["sharepoint"]["tenant_id"]
            _cid = st.secrets["sharepoint"]["client_id"]
            _cs  = st.secrets["sharepoint"]["client_secret"]
            with st.spinner(f"Creando {_nombre} en SharePoint…"):
                _r = crear_hoja_wk(_nombre, _tid, _cid, _cs)
            if _r.get("ok"):
                st.success(_r["mensaje"])
                st.cache_data.clear()
            else:
                st.error(f"❌ {_r['error']}")
        except KeyError as _e:
            st.error(f"❌ Falta credencial: {_e}")
        except ImportError:
            st.error("❌ crear_hoja_wk no disponible.")

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
  overflow-x: hidden;
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
  scrollbar-width: thin;
  scrollbar-color: #b0c4d8 transparent;
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
.cmp-tbl-wrap { overflow: auto; scrollbar-width: thin; scrollbar-color: #b0c4d8 transparent; }
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
.prod-hdr {
  background: #f8fafc; padding: 4px 8px; border-bottom: 1px solid #cbd5e1;
  display: flex; align-items: center; gap: 10px; height: 26px;
}
.prod-hdr-title { color: #0f172a; font-size: 11px; font-weight: 700; flex: 1; text-transform: uppercase; }
.prod-hdr-meta  { display: none; }
.prod-close {
  background: transparent; border: 1px solid #cbd5e1;
  border-radius: 2px; color: #475569; font-weight: 600;
  cursor: pointer; font-size: 9px; padding: 2px 8px; height: 18px; font-family: inherit; line-height: 1; transition: all 0.2s;
}
.prod-close:hover { border-color: #0f172a; color: #0f172a; background: #fff; }
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
  <div class="app-hdr" style="position:relative;">
    <div class="hdr-brand">CFBC &#9656; CONTROL SEMANAL</div>
    <button class="hdr-btn" onclick="exportCSV()" style="margin-left:auto">&#11015; CSV</button>
    <button class="hdr-btn" id="btnExcel" onclick="toggleExcelPanel(event)" style="margin-left:4px">&#11015; EXCEL</button>
    <button class="hdr-btn" onclick="recargar()" style="margin-left:4px">&#8635;</button>

    <!-- PANEL EXCEL FLOTANTE -->
    <div id="excelPanel" style="
        display:none; position:absolute; top:36px; right:30px; z-index:9999;
        background:#fff; border:1px solid #c0c0c0; border-radius:4px;
        box-shadow:0 4px 16px rgba(0,0,0,0.18); min-width:215px; padding:12px 14px;
        font-family:Calibri,'Segoe UI',Arial,sans-serif;">

      <!-- Sección Descargar -->
      <p style="font-size:11px;font-weight:bold;color:#1e3a5f;margin:0 0 6px 0;">&#11015; Descargar Archivo WK</p>
      <select id="excelWkSel" style="width:100%;font-size:11px;padding:3px 6px;border:1px solid #bbb;border-radius:3px;margin-bottom:7px;font-family:inherit;"></select>
      <button onclick="doDownloadExcel()" style="width:100%;font-size:10px;font-weight:700;background:#4472C4;color:#fff;border:none;border-radius:3px;padding:5px;cursor:pointer;font-family:inherit;">Preparar y Descargar XLSX</button>
      <div id="excelStatus" style="font-size:10px;color:#555;margin-top:5px;min-height:13px;"></div>

      <!-- Divisor -->
      <div style="border-top:1px solid #e0e0e0;margin:10px 0;"></div>

      <!-- Sección Crear Hoja -->
      <p style="font-size:11px;font-weight:bold;color:#1e3a5f;margin:0 0 6px 0;">&#10010; Nueva Hoja SharePoint</p>
      <input id="excelNuevoNombre" type="text" placeholder="Ej: WK2518"
        style="width:100%;font-size:11px;padding:3px 6px;border:1px solid #bbb;border-radius:3px;margin-bottom:7px;font-family:inherit;text-transform:uppercase;"
        oninput="this.value=this.value.toUpperCase()">
      <button onclick="doCrearHoja()" style="width:100%;font-size:10px;font-weight:700;background:#16a34a;color:#fff;border:none;border-radius:3px;padding:5px;cursor:pointer;font-family:inherit;">Crear Hoja</button>
      <div id="crearStatus" style="font-size:10px;color:#555;margin-top:5px;min-height:13px;"></div>
    </div>
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
    <div class="prod-hdr">
      <div class="prod-hdr-title" id="prodTitle">COMPARADOR DE PRODUCTOS</div>
      <button class="prod-close" onclick="closeProdPanel()">&#10005; CERRAR</button>
    </div>
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
function onCatChange(val) { state.cat=val; renderView(); }
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
  if (rb) rb.className='range-bar'+(v==='comparativo'?' show':'');
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
  if (state.view==='comparativo') renderComparativo();
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
  else if (state.view==='servicios')   renderServicios();
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
  var svRows={};
  if (Array.isArray(DATA.servicios_data)&&DATA.servicios_data.length){
    DATA.servicios_data.forEach(function(r){
      if (!state.activeYears[r.year]) return;
      var subcat=(r.subcat||'').trim(); if (!subcat) return;
      if (!svRows[subcat]) svRows[subcat]={};
      var src=state.currency==='usd'?(r.usd_ranches||{}):(r.mxn_ranches||{});
      RANCH_ORDER.forEach(function(rn){var v=src[rn]||0;if(v>0)svRows[subcat][rn]=(svRows[subcat][rn]||0)+v;});
      var total=state.currency==='usd'?r.usd_total:r.mxn_total;
      svRows[subcat]._total=(svRows[subcat]._total||0)+(total||0);
    });
  } else {
    DATA.weekly_detail.forEach(function(r){
      if (!state.activeYears[r.year]) return;
      if (!r.categoria||!r.categoria.startsWith('SV:')) return;
      var subcat=r.categoria.replace('SV:','');
      if (!svRows[subcat]) svRows[subcat]={};
      RANCH_ORDER.forEach(function(rn){var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;var v=src[rn]||0;if(v>0)svRows[subcat][rn]=(svRows[subcat][rn]||0)+v;});
      svRows[subcat]._total=(svRows[subcat]._total||0)+(state.currency==='usd'?r.usd_total:r.mxn_total);
    });
  }
  var cols=[
    { field:'subcat', headerName:'SUBCATEGORÍA', pinned:'left', width:210,
      cellRenderer:function(p){return '<span style="font-weight:700;color:#1e3a5f">'+(p.value||'')+'</span>';}},
    { field:'total', headerName:'TOTAL '+sym, width:110, type:'numericColumn', cellRenderer:moneyRenderer },
    { field:'pct',   headerName:'% DEL TOTAL', width:90,  type:'numericColumn',
      cellRenderer:function(p){
        var v=p.value; if(!v) return '';
        var w=Math.min(v/100*50,50);
        return '<div style="display:flex;align-items:center;gap:5px"><div style="width:'+w.toFixed(0)+'px;height:6px;background:#2E74B5;border-radius:2px;flex-shrink:0"></div><span>'+v.toFixed(1)+'%</span></div>';
      }},
  ];
  RANCH_ORDER.forEach(function(r){
    var c=RANCH_COLORS[r]||'#888';
    cols.push((function(color){return {field:'r_'+r.replace(/[^a-zA-Z0-9]/g,'_'),headerName:r,width:100,type:'numericColumn',
      cellRenderer:function(p){var v=p.value;if(!v||v<0.01)return '';return '<span style="color:'+color+'">'+fmt(v)+'</span>';}};})(c));
  });
  var grandTotal=Object.values(svRows).reduce(function(s,r){return s+(r._total||0);},0);
  var orderedSubcats=SV_SUBCATS.filter(function(sc){return svRows[sc];});
  Object.keys(svRows).forEach(function(sc){if(orderedSubcats.indexOf(sc)===-1)orderedSubcats.push(sc);});
  var rows=orderedSubcats.map(function(sc){
    var data=svRows[sc]||{};
    var row={subcat:sc,total:data._total||0,pct:grandTotal>0?(data._total||0)/grandTotal*100:0};
    RANCH_ORDER.forEach(function(r){row['r_'+r.replace(/[^a-zA-Z0-9]/g,'_')]=data[r]||0;});
    return row;
  });
  rows.sort(function(a,b){return b.total-a.total;});
  renderPivotTable(cols,rows,fmt(grandTotal)+' '+sym);
}

// =======================================================
// CELL CLICK HANDLER
// =======================================================
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
          rows.push({week_code:wkCodeShort,rancho:ranch,tipo:tipo,producto:item[0]||'',unidades:item[1]||'',gasto:parseFloat(item[2])||0});
        });
      });
    });
  }

  var rangeText=wkStart===wkEnd?(wFmt(wkStart)+' · '+yr):(wFmt(wkStart)+'→'+wFmt(wkEnd)+' · '+yr);
  var panelTitle = cat+' &#9656; '+rangeText+(ranchFilter?' · '+ranchFilter:'');
  
  var panelHtml = '';
  if (rows.length===0){
    panelHtml = '<div style="flex:1; min-width:320px; border:1px solid #cbd5e1; border-top:2px solid #0f172a; background:#fff;"><p style="padding:8px;color:#64748b;font-size:11px;margin:0;">No hay registros para este período.</p></div>';
  } else {
    rows.sort(function(a,b){return b.gasto-a.gasto;});
    var total=rows.reduce(function(s,r){return s+r.gasto;},0);
    var panelMeta = 'Reg: <b>' + rows.length + '</b> &nbsp;|&nbsp; Gasto: <b style="color:#16a34a">' + fmt(total) + '</b>';

    var html='<div style="flex:1; min-width:320px; border:1px solid #cbd5e1; border-top:2px solid #0f172a; display:flex; flex-direction:column; background:#fff; overflow:hidden;">' +
      '<div style="background:#f1f5f9; color:#0f172a; padding:4px 6px; border-bottom:1px solid #cbd5e1; flex-shrink:0; display:flex; justify-content:space-between; align-items:baseline;">' + 
      '<div style="font-weight:bold; font-size:11px; text-transform:uppercase; letter-spacing:0px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="'+panelTitle+'">' + panelTitle + '</div>' + 
      '<div style="color:#475569; font-size:10px; margin-left:8px; white-space:nowrap;">' + panelMeta + '</div></div>' +
      '<div style="overflow-x:auto; scrollbar-width:thin;"><table class="pt-table" style="font-size:10px; width:100%; border-collapse:collapse;"><thead><tr>'+
      '<th style="text-align:left; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">WK</th>'+
      '<th style="text-align:left; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">RANCHO</th>'+
      '<th style="text-align:left; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">TIPO</th>'+
      '<th style="text-align:left; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">PRODUCTO</th>'+
      '<th style="text-align:left; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">UNID.</th>'+
      '<th style="text-align:right; background:#fff; border-bottom:1px solid #cbd5e1; padding:3px 5px; color:#475569;">GASTO</th>'+
      '</tr></thead><tbody>';
    rows.forEach(function(r,i){
      var rc=RANCH_COLORS[r.rancho]||'#64748b';
      var rowBg = (i % 2 === 0) ? '#ffffff' : '#f8fafc';
      html+='<tr style="background:'+rowBg+'; border-bottom:1px solid #f1f5f9;">'+
        '<td style="padding:2px 5px; color:#64748b;">'+r.week_code+'</td>'+
        '<td style="padding:2px 5px; white-space:nowrap;"><span style="color:'+rc+';font-weight:600">'+r.rancho+'</span></td>'+
        '<td style="padding:2px 5px; color:#64748b;">'+r.tipo+'</td>'+
        '<td style="padding:2px 5px; color:#0f172a; font-weight:500;">'+r.producto+'</td>'+
        '<td style="padding:2px 5px; color:#94a3b8; font-size:9px;">'+r.unidades+'</td>'+
        '<td style="padding:2px 5px; text-align:right;"><span style="font-weight:600; color:#0f172a;">'+fmt(r.gasto)+'</span></td>'+
        '</tr>';
    });
    html+='</tbody></table></div></div>';
    panelHtml = html;
  }
  
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
  // Las tablas ya no tienen altura forzada para usar el scroll nativo.
  var tw=document.getElementById('tableWrap');
  if (tw) tw.style.height='auto';
  var cmpWrap=document.querySelector('.cmp-tbl-wrap');
  if (cmpWrap) cmpWrap.style.maxHeight='none';
}
window.addEventListener('resize', resizeTable);

// =======================================================
// HEIGHT REPORTING
// =======================================================
function reportHeight() {
  var appEl=document.getElementById('app');
  var h=appEl?appEl.scrollHeight+60:document.body.scrollHeight+60;
  window.parent.postMessage({type:'streamlit:setFrameHeight',height:Math.max(h,700)},'*');
}
var ro=new ResizeObserver(reportHeight);
ro.observe(document.body);
reportHeight();
// =======================================================
// EXCEL PANEL
// =======================================================
(function initExcelWeeks() {
  var weeks = [];
  (DATA.weekly_detail || []).forEach(function(r) {
    var code = String(r.year % 100).padStart(2,'0') + String(r.week).padStart(2,'0');
    if (weeks.indexOf(code) === -1) weeks.push(code);
  });
  weeks.sort(function(a,b){ return b.localeCompare(a); });
  var sel = document.getElementById('excelWkSel');
  if (sel) weeks.forEach(function(w){
    var o = document.createElement('option'); o.value = w; o.text = 'WK' + w; sel.appendChild(o);
  });
})();

function toggleExcelPanel(e) {
  if (e) e.stopPropagation();
  var p = document.getElementById('excelPanel');
  p.style.display = (p.style.display === 'none') ? 'block' : 'none';
}
document.addEventListener('click', function(e) {
  var p   = document.getElementById('excelPanel');
  var btn = document.getElementById('btnExcel');
  if (p && !p.contains(e.target) && e.target !== btn) p.style.display = 'none';
});

async function doDownloadExcel() {
  var wk  = document.getElementById('excelWkSel').value;
  var st2 = document.getElementById('excelStatus');
  if (!wk) { st2.textContent = '⚠ Selecciona una semana.'; return; }
  st2.textContent = '⏳ Descargando desde SharePoint…';
  var spUrl = '__SP_URL__'.replace('?e=', '?download=1&e=');
  try {
    var resp = await fetch(spUrl);
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    var buf = await resp.arrayBuffer();
    st2.textContent = '⚙ Extrayendo hoja WK' + wk + '…';
    var wb = XLSX.read(buf, {type: 'array', cellStyles: true, cellFormulas: true});
    var target = wb.SheetNames.find(function(n){
      return n.replace(/\s+/g,'').toUpperCase() === ('WK' + wk).toUpperCase();
    });
    if (!target) { st2.textContent = '❌ No se encontró WK' + wk + '.'; return; }
    var newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, wb.Sheets[target], target);
    XLSX.writeFile(newWb, 'WK' + wk + '.xlsx');
    st2.textContent = '✅ Descargado WK' + wk + '.xlsx';
  } catch(err) {
    st2.textContent = '❌ Error: ' + err.message;
  }
}

function doCrearHoja() {
  var nombre = (document.getElementById('excelNuevoNombre').value || '').trim().toUpperCase();
  var st3    = document.getElementById('crearStatus');
  if (!nombre) { st3.textContent = '⚠ Escribe el nombre.'; return; }
  if (!/^WK\d{4}$/.test(nombre)) { st3.textContent = '⚠ Formato: WK####'; return; }
  st3.textContent = '⏳ Procesando…';
  // Comunicar con Streamlit via query param y recargar
  window.parent.location.href = window.parent.location.pathname +
    '?_action=crear_hoja&_nombre=' + encodeURIComponent(nombre);
}

setInterval(reportHeight, 500);

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
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
{APP_CSS}
</head>
<body>
{APP_HTML_BODY}
{APP_JS}
</body>
</html>'''

html_final = HTML.replace('__DATA_JSON__', data_json).replace('__SP_URL__', SHAREPOINT_URL_WK)

# ─── RENDERIZAR IFRAME ────────────────────────────────────────────────────────
components.html(html_final, height=800, scrolling=False)
