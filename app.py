"""
app.py
Centro Floricultor de Baja California
Streamlit — Tablas HTML ejecutivas, sin AG Grid
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
  .stApp { background: #ffffff; }
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
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 12px; background: #fff; color: #111; }

/* ── LOADER ── */
#loader {
  position: fixed; inset: 0; background: #fff; z-index: 999;
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; gap: 12px;
}
.spin {
  width: 32px; height: 32px;
  border: 3px solid #e0e0e0; border-top-color: #0071ce;
  border-radius: 50%; animation: spin .9s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* ── HEADER ── */
.app-hdr {
  background: #0d1f3c; border-bottom: 3px solid #0071ce;
  padding: 0 12px; display: flex; align-items: center;
  height: 36px; gap: 10px;
}
.hdr-brand { color: #fff; font-size: 11px; font-weight: 700; letter-spacing: 1px; flex-shrink: 0; }
.hdr-sep   { width: 1px; height: 18px; background: rgba(255,255,255,.2); flex-shrink: 0; }
.hdr-btn {
  font-size: 10px; font-weight: 700; background: rgba(255,255,255,.1);
  border: 1px solid rgba(255,255,255,.25); border-radius: 3px;
  padding: 2px 10px; cursor: pointer; color: rgba(255,255,255,.85);
  height: 22px; margin-left: auto; transition: background .1s;
}
.hdr-btn:hover { background: rgba(255,255,255,.22); }
.hdr-btn + .hdr-btn { margin-left: 4px; }

/* ── TOOLBAR ── */
.toolbar {
  background: #f5f7fa; border-bottom: 1px solid #ddd;
  padding: 4px 10px; display: flex; align-items: center;
  gap: 6px; height: 32px; overflow-x: auto; flex-wrap: nowrap;
}
.toolbar::-webkit-scrollbar { display: none; }
.tb-lbl { font-size: .65rem; color: #777; text-transform: uppercase; letter-spacing: .4px; white-space: nowrap; font-weight: 700; }
.tb-sep { width: 1px; height: 16px; background: #ccc; flex-shrink: 0; }
select.tb-sel {
  font-size: .72rem; background: #fff; border: 1px solid #bbb;
  border-radius: 3px; padding: 1px 5px; color: #222; height: 22px; cursor: pointer;
}
.tb-btn {
  font-size: .68rem; font-weight: 700; background: #fff;
  border: 1px solid #bbb; border-radius: 3px;
  padding: 1px 8px; cursor: pointer; height: 22px;
  color: #333; white-space: nowrap; transition: background .1s;
}
.tb-btn:hover  { background: #eef4fb; border-color: #0071ce; color: #0071ce; }
.tb-btn.active { background: #0d1f3c; color: #fff; border-color: #0d1f3c; }
.tb-grp { display: flex; }
.tb-grp .tb-btn { border-radius: 0; border-right-width: 0; }
.tb-grp .tb-btn:first-child { border-radius: 3px 0 0 3px; }
.tb-grp .tb-btn:last-child  { border-radius: 0 3px 3px 0; border-right-width: 1px; }
.week-ctr { display: flex; align-items: center; gap: 4px; }
.week-ctr span { font-size: .72rem; font-weight: 700; color: #0d1f3c; min-width: 56px; text-align: center; }
.tb-slider { width: 90px; accent-color: #0071ce; cursor: pointer; }
.yr-chip {
  font-size: .68rem; font-weight: 700; padding: 1px 7px;
  border-radius: 3px; cursor: pointer; border: 1px solid transparent;
  background: transparent; transition: all .1s;
}
.yr-chip.on { background: #fff; }

/* ── RANGE BAR (comparativo) ── */
.range-bar {
  display: none; background: #f5f7fa; border-bottom: 1px solid #ddd;
  padding: 3px 10px; align-items: center; gap: 8px; height: 28px;
}
.range-bar.show { display: flex; }
.range-val { font-size: .70rem; font-weight: 700; color: #0d1f3c; min-width: 34px; text-align: center; }
.range-badge {
  font-size: .66rem; background: #eef3fb; border: 1px solid #c5d6f0;
  color: #0071ce; padding: 1px 8px; border-radius: 3px;
}

/* ── VIEW TABS ── */
.view-tabs {
  background: #f5f7fa; border-bottom: 2px solid #ddd;
  display: flex; height: 28px;
}
.vtab {
  padding: 0 14px; font-size: .68rem; font-weight: 700;
  cursor: pointer; border: none; background: transparent; color: #888;
  border-bottom: 2px solid transparent; margin-bottom: -2px;
  text-transform: uppercase; letter-spacing: .3px;
  transition: color .1s; white-space: nowrap; height: 28px;
}
.vtab:hover  { color: #0071ce; background: rgba(0,113,206,.04); }
.vtab.active { color: #0071ce; border-bottom-color: #0071ce; background: #fff; }

/* ── STATUS BAR ── */
.statusbar {
  background: #f5f7fa; border-top: 1px solid #ddd;
  padding: 2px 12px; font-size: .66rem; color: #666;
  display: flex; align-items: center; gap: 12px; height: 20px;
}
.statusbar b { color: #111; }

/* ── TABLA EJECUTIVA — estilo limpio ── */
.tbl-wrap {
  overflow-x: auto; overflow-y: auto;
  background: #fff;
}
.tbl-wrap::-webkit-scrollbar { height: 4px; width: 4px; }
.tbl-wrap::-webkit-scrollbar-thumb { background: #ddd; border-radius: 2px; }

table.t {
  border-collapse: collapse; width: 100%;
  font-family: Arial, sans-serif; font-size: .71rem;
}
table.t thead th {
  padding: 4px 12px;
  background: #fff; color: #555;
  font-size: .64rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: .2px;
  white-space: nowrap;
  border-bottom: 2px solid #e0e0e0;
  text-align: right;
  position: sticky; top: 0; z-index: 2;
  cursor: pointer; user-select: none;
}
table.t thead th:first-child { text-align: left; }
table.t thead th:hover { color: #0071ce; }
table.t thead th.asc::after  { content: ' ▲'; font-size: .5rem; }
table.t thead th.desc::after { content: ' ▼'; font-size: .5rem; }

table.t tbody td {
  padding: 3px 12px;
  border-bottom: 1px solid #f2f2f2;
  text-align: right; color: #222;
  white-space: nowrap;
}
table.t tbody td:first-child { text-align: left; color: #111; }
table.t tbody tr:hover td { background: #f7f9fc; }

/* Fila de grupo (año/semana agrupador) */
table.t tr.grp td {
  background: #f5f7fa; font-weight: 700; color: #111;
  border-top: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0;
  padding: 5px 12px; font-size: .72rem;
}
table.t tr.grp td:first-child { border-left: 3px solid #0071ce; padding-left: 10px; }

/* Fila subtotal */
table.t tr.sub td {
  background: #eef3fb; font-weight: 700;
  border-top: 1px solid #d0dcea; color: #111;
}
/* Fila total general */
table.t tr.total td {
  background: #fff; font-weight: 700;
  border-top: 2px solid #bbb; color: #111;
  font-size: .72rem;
}

/* Colores positivo/negativo */
.pos { color: #16a34a; font-weight: 600; }
.neg { color: #dc2626; font-weight: 600; }
.mut { color: #bbb; }

/* Clickable cells */
.clickable { cursor: pointer; text-decoration: underline dotted; text-underline-offset: 2px; }

/* ── PRODUCTOS PANEL ── */
#prodPanel {
  display: none; background: #fff; border-top: 2px solid #0071ce;
}
#prodPanel.show { display: block; }
.prod-hdr {
  background: #0d1f3c; padding: 4px 12px;
  display: flex; align-items: center; gap: 8px; height: 26px;
}
.prod-hdr-title { color: #fff; font-size: .70rem; font-weight: 700; flex: 1; }
.prod-hdr-meta  { color: rgba(255,255,255,.6); font-size: .66rem; }
.prod-close {
  background: transparent; border: 1px solid rgba(255,255,255,.3);
  border-radius: 3px; color: rgba(255,255,255,.8);
  cursor: pointer; font-size: .66rem; padding: 1px 8px;
}
.prod-close:hover { border-color: #fff; color: #fff; }
#prodTblWrap { max-height: 280px; overflow-y: auto; overflow-x: auto; }
#prodTblWrap::-webkit-scrollbar { height: 4px; width: 4px; }
#prodTblWrap::-webkit-scrollbar-thumb { background: #ddd; border-radius: 2px; }
</style>
</head>
<body>

<div id="loader"><div class="spin"></div><div style="font-size:11px;color:#888">Cargando datos...</div></div>

<div id="app" style="display:none">

  <!-- HEADER -->
  <div class="app-hdr">
    <div class="hdr-brand">CFBC ▸ CONTROL SEMANAL</div>
    <div class="hdr-sep"></div>
    <button class="hdr-btn" onclick="exportCSV()">⬇ CSV</button>
    <button class="hdr-btn" onclick="recargar()">⟳</button>
  </div>

  <!-- TOOLBAR -->
  <div class="toolbar">
    <span class="tb-lbl">Cat</span>
    <select class="tb-sel" id="catSel" onchange="onCatChange(this.value)" style="max-width:200px"></select>
    <div class="tb-sep"></div>
    <div class="tb-grp">
      <button class="tb-btn active" id="btnUSD" onclick="setCurrency('usd')">USD</button>
      <button class="tb-btn"        id="btnMXN" onclick="setCurrency('mxn')">MXN</button>
    </div>
    <div class="tb-sep"></div>
    <span class="tb-lbl">Semana</span>
    <div class="week-ctr">
      <button class="tb-btn" onclick="prevWeek()">◀</button>
      <span id="weekLabel">—</span>
      <button class="tb-btn" onclick="nextWeek()">▶</button>
    </div>
    <input type="range" class="tb-slider" id="weekSlider" min="1" max="52" value="1" oninput="onWeekSlider(this.value)">
    <div class="tb-sep"></div>
    <span class="tb-lbl">Años</span>
    <div id="yearChips" style="display:flex;gap:3px"></div>
  </div>

  <!-- VIEW TABS -->
  <div class="view-tabs">
    <button class="vtab active" id="vtSemana"      onclick="setView('semana')">Semana</button>
    <button class="vtab"        id="vtAnual"        onclick="setView('anual')">Anual</button>
    <button class="vtab"        id="vtComparativo"  onclick="setView('comparativo')">Comparativo</button>
    <button class="vtab"        id="vtServicios"    onclick="setView('servicios')">Costo Servicios</button>
  </div>

  <!-- RANGE BAR — solo comparativo -->
  <div class="range-bar" id="rangeBar">
    <span class="tb-lbl">Desde</span>
    <span class="range-val" id="fromWeekLabel">W01</span>
    <input type="range" class="tb-slider" id="fromSlider" min="1" max="52" value="1" oninput="onRangeChange()">
    <span style="color:#aaa">→</span>
    <span class="tb-lbl">Hasta</span>
    <span class="range-val" id="toWeekLabel">W52</span>
    <input type="range" class="tb-slider" id="toSlider" min="1" max="52" value="52" oninput="onRangeChange()">
    <span class="range-badge" id="rangeBadge">W01 → W52</span>
    <div class="tb-sep"></div>
    <button class="tb-btn" onclick="resetRange()">↺ Reset</button>
  </div>

  <!-- CONTENIDO PRINCIPAL -->
  <div class="tbl-wrap" id="mainWrap" style="height:calc(100vh - 120px)">
    <table class="t" id="mainTbl">
      <thead id="mainHead"></thead>
      <tbody id="mainBody"></tbody>
    </table>
  </div>

  <!-- PRODUCTOS SUB-PANEL -->
  <div id="prodPanel">
    <div class="prod-hdr">
      <div class="prod-hdr-title" id="prodTitle">PRODUCTOS</div>
      <div class="prod-hdr-meta"  id="prodMeta"></div>
      <button class="prod-close"  onclick="closeProdPanel()">✕ CERRAR</button>
    </div>
    <div id="prodTblWrap">
      <table class="t" id="prodTbl">
        <thead id="prodHead"></thead>
        <tbody id="prodBody"></tbody>
      </table>
    </div>
  </div>

  <!-- STATUS BAR -->
  <div class="statusbar"><span>Total: <b id="stTotal">—</b></span></div>

</div><!-- /app -->

<script>
var DATA = JSON.parse(atob('__DATA_JSON__'));

// ── Reconstruir weekly_series si falta ──
if (!DATA.weekly_series) {
  DATA.weekly_series = {};
  DATA.categories.forEach(function(c){ DATA.weekly_series[c] = {}; });
  DATA.weekly_detail.forEach(function(r){
    if (r.usd_total > 0){
      if (!DATA.weekly_series[r.categoria]) DATA.weekly_series[r.categoria] = {};
      var k = r.year+'-W'+String(r.week).padStart(2,'0');
      DATA.weekly_series[r.categoria][k] = (DATA.weekly_series[r.categoria][k]||0) + r.usd_total;
    }
  });
}

var RANCH_ORDER  = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','HOOPS','Cecilia','Cecilia 25','Christina','Albahaca-RM','Campo-VI'];
var RANCH_COLORS = {'Prop-RM':'#047857','PosCo-RM':'#1d4ed8','Campo-RM':'#b45309','Isabela':'#7c3aed','HOOPS':'#c2410c','Cecilia':'#be185d','Cecilia 25':'#047857','Christina':'#0369a1','Albahaca-RM':'#6d28d9','Campo-VI':'#64748b'};
var YEAR_COLORS  = {2021:'#0ea5e9',2022:'#f59e0b',2023:'#22c55e',2024:'#a855f7',2025:'#f97316',2026:'#ef4444'};
var CAT_MIRFE = 'FERTILIZANTES', CAT_MIPE = 'DESINFECCION / PLAGUICIDAS';
var SV_SUBCATS = ['Electricidad','Fletes y Acarreos','Gastos de Exportación','Certificado Fitosanitario','Transporte de Personal','Compra de Flor a Terceros','Comida para el Personal','RO, TEL, RTA.Alim'];

var state = { cat:'', currency:'usd', activeYears:{}, view:'semana', weekIdx:0, fromWeek:1, toWeek:52 };
var allWeeks = [];
var _sortState = {}; // {col, dir} por vista

// ── FORMATO ──
function fmt(n){
  if (!n || isNaN(n) || n===0) return '—';
  var neg=n<0, s=Math.abs(n);
  return (neg?'-$':'$')+s.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:0});
}
function fmtPct(n){
  if (n===null||n===undefined||isNaN(n)) return '—';
  return (n>0?'+':'')+n.toFixed(1)+'%';
}
function wFmt(n){ return 'W'+String(n).padStart(2,'0'); }
function recargar(){ window.location.reload(); }

// ── DATOS HELPERS ──
function getActiveYears(){ return DATA.years.filter(function(y){ return state.activeYears[y]; }); }
function getWeekDetail(cat,wk,yr){
  return DATA.weekly_detail.filter(function(r){ return r.categoria===cat&&r.week===wk&&r.year===yr; });
}
function sumDetail(recs){
  var out={total:0,ranches:{}};
  recs.forEach(function(r){
    var v = state.currency==='usd'?r.usd_total:r.mxn_total;
    out.total+=v;
    var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
    Object.keys(src||{}).forEach(function(rn){ out.ranches[rn]=(out.ranches[rn]||0)+src[rn]; });
  });
  return out;
}
function aggregateRecs(recs){
  var out={usd:0,mxn:0,ranches:{},ranches_mxn:{},date_range:''};
  recs.forEach(function(r){
    out.usd+=r.usd_total; out.mxn+=r.mxn_total;
    if(r.date_range) out.date_range=r.date_range;
    Object.keys(r.usd_ranches||{}).forEach(function(rn){ out.ranches[rn]=(out.ranches[rn]||0)+r.usd_ranches[rn]; });
    Object.keys(r.mxn_ranches||{}).forEach(function(rn){ out.ranches_mxn[rn]=(out.ranches_mxn[rn]||0)+r.mxn_ranches[rn]; });
  });
  return out;
}
function fmtMes(dr){
  if(!dr) return '—';
  var MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var lower=dr.toLowerCase();
  for(var i=0;i<MESES.length;i++){
    if(lower.indexOf(MESES[i])>-1){
      var m=MESES[i].charAt(0).toUpperCase()+MESES[i].slice(1);
      var ym=dr.match(/\b(20\d{2})\b/);
      return m+(ym?' '+ym[1]:'');
    }
  }
  return dr;
}

// ── DELTA CELL HTML ──
function deltaHtml(val, prev){
  if(prev===null||prev===undefined||prev===0) return '<td class="mut">—</td>';
  var d=val-prev, p=((d/prev)*100).toFixed(1);
  var cls=d>0?'pos':d<0?'neg':'mut';
  var sign=d>0?'+':'';
  return '<td class="'+cls+'" style="font-size:.67rem">'+sign+fmt(d)+'<br><span style="font-size:.60rem;opacity:.8">'+sign+p+'%</span></td>';
}

// ── TABLE RENDER ──
function setTable(headHtml, bodyHtml, statusText){
  document.getElementById('mainHead').innerHTML = headHtml;
  document.getElementById('mainBody').innerHTML = bodyHtml;
  document.getElementById('stTotal').textContent = statusText||'';
  // bind sort
  document.querySelectorAll('#mainHead th[data-col]').forEach(function(th){
    th.onclick = function(){ sortTable(th.dataset.col, th.dataset.num==='1'); };
  });
}

function sortTable(col, numeric){
  var tbody = document.getElementById('mainBody');
  var rows = Array.from(tbody.querySelectorAll('tr:not(.grp):not(.sub):not(.total)'));
  if (!rows.length) return;
  var dir = (_sortState.col===col && _sortState.dir==='asc') ? 'desc' : 'asc';
  _sortState = {col:col, dir:dir};
  rows.sort(function(a,b){
    var ca=a.querySelector('td[data-col="'+col+'"]'), cb=b.querySelector('td[data-col="'+col+'"]');
    if(!ca||!cb) return 0;
    var va=ca.dataset.val||ca.textContent.replace(/[$,]/g,'').trim();
    var vb=cb.dataset.val||cb.textContent.replace(/[$,]/g,'').trim();
    if(numeric){ va=parseFloat(va)||0; vb=parseFloat(vb)||0; }
    if(va<vb) return dir==='asc'?-1:1;
    if(va>vb) return dir==='asc'?1:-1;
    return 0;
  });
  rows.forEach(function(r){ tbody.appendChild(r); });
  document.querySelectorAll('#mainHead th').forEach(function(th){
    th.classList.remove('asc','desc');
    if(th.dataset.col===col) th.classList.add(dir);
  });
}

// ── EXPORT CSV ──
function exportCSV(){
  var rows=[], head=document.querySelectorAll('#mainHead th');
  var hRow=[]; head.forEach(function(th){ hRow.push('"'+th.textContent.trim()+'"'); });
  rows.push(hRow.join(','));
  document.querySelectorAll('#mainBody tr').forEach(function(tr){
    var cells=[]; tr.querySelectorAll('td').forEach(function(td){ cells.push('"'+(td.dataset.val||td.textContent.replace(/"/g,"'")).trim()+'"'); });
    if(cells.length) rows.push(cells.join(','));
  });
  var blob=new Blob([rows.join('\n')],{type:'text/csv'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='CFBC_'+state.view+'_'+new Date().toISOString().slice(0,10)+'.csv';
  a.click();
}

// ── INIT ──
function inicializar(){
  var prefCat='MATERIAL DE EMPAQUE';
  state.cat = DATA.categories.indexOf(prefCat)>-1 ? prefCat : DATA.categories[0];
  state.activeYears={};
  var latestYr=DATA.years[DATA.years.length-1];
  var prevYr=DATA.years[DATA.years.length-2];
  if(latestYr) state.activeYears[latestYr]=true;
  if(prevYr)   state.activeYears[prevYr]=true;

  var wSet={};
  DATA.weekly_detail.forEach(function(r){ wSet[r.week]=1; });
  allWeeks=Object.keys(wSet).map(Number).sort(function(a,b){return a-b;});

  var wksLatest=DATA.weekly_detail.filter(function(r){return r.year===latestYr;}).map(function(r){return r.week;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort(function(a,b){return a-b;});
  var curWeek=wksLatest[wksLatest.length-1]||allWeeks[allWeeks.length-1];
  var idx=allWeeks.indexOf(curWeek);
  state.weekIdx=idx>=0?idx:allWeeks.length-1;
  state.toWeek=wksLatest[wksLatest.length-1]||allWeeks[allWeeks.length-1]||52;
  state.fromWeek=wksLatest[wksLatest.length-2]||wksLatest[0]||state.toWeek;

  buildCatSelect(); buildYearChips(); updateWeekControls(); updateRangeSliders();
  renderView();
  document.getElementById('loader').style.display='none';
  document.getElementById('app').style.display='block';
  resizeWrap();
}

function buildCatSelect(){
  var el=document.getElementById('catSel');
  el.innerHTML=DATA.categories.map(function(c){
    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+'>'+c+'</option>';
  }).join('');
}
function buildYearChips(){
  var el=document.getElementById('yearChips');
  el.innerHTML=DATA.years.map(function(y){
    var col=YEAR_COLORS[y]||'#888';
    var on=state.activeYears[y]?' on':'';
    return '<button class="yr-chip'+on+'" id="yrChip'+y+'" style="color:'+col+';border-color:'+(state.activeYears[y]?col:'transparent')+';background:'+(state.activeYears[y]?col+'20':'transparent')+'" onclick="toggleYear('+y+')">'+y+'</button>';
  }).join('');
}
function updateWeekControls(){
  var wn=allWeeks[state.weekIdx]||1;
  var sl=document.getElementById('weekSlider');
  sl.min=allWeeks[0]||1; sl.max=allWeeks[allWeeks.length-1]||52; sl.value=wn;
  var yrs=getActiveYears();
  var yr=yrs[yrs.length-1]||DATA.years[DATA.years.length-1];
  document.getElementById('weekLabel').textContent=String(yr).slice(2)+String(wn).padStart(2,'0');
}
function updateRangeSliders(){
  var f=state.fromWeek, t=state.toWeek;
  var min=allWeeks[0]||1, max=allWeeks[allWeeks.length-1]||52;
  var fEl=document.getElementById('fromSlider'), tEl=document.getElementById('toSlider');
  if(fEl){fEl.min=min;fEl.max=max;fEl.value=f;}
  if(tEl){tEl.min=min;tEl.max=max;tEl.value=t;}
  var yr=DATA.years[DATA.years.length-1]; var yy=String(yr).slice(2);
  var fLbl=document.getElementById('fromWeekLabel'), tLbl=document.getElementById('toWeekLabel'), badge=document.getElementById('rangeBadge');
  if(fLbl) fLbl.textContent=yy+String(f).padStart(2,'0');
  if(tLbl) tLbl.textContent=yy+String(t).padStart(2,'0');
  var cnt=allWeeks.filter(function(w){return w>=f&&w<=t;}).length;
  if(badge) badge.textContent=yy+String(f).padStart(2,'0')+' → '+yy+String(t).padStart(2,'0')+' · '+cnt+' sem';
}

// ── EVENTS ──
function onCatChange(v){ state.cat=v; renderView(); }
function setCurrency(c){
  state.currency=c;
  document.getElementById('btnUSD').className='tb-btn'+(c==='usd'?' active':'');
  document.getElementById('btnMXN').className='tb-btn'+(c==='mxn'?' active':'');
  renderView();
}
function toggleYear(y){
  var active=DATA.years.filter(function(yr){return state.activeYears[yr];});
  if(state.activeYears[y]&&active.length>1) delete state.activeYears[y];
  else state.activeYears[y]=true;
  buildYearChips(); renderView();
}
function prevWeek(){ if(state.weekIdx>0){state.weekIdx--;updateWeekControls();renderView();} }
function nextWeek(){ if(state.weekIdx<allWeeks.length-1){state.weekIdx++;updateWeekControls();renderView();} }
function onWeekSlider(v){
  var wn=parseInt(v), idx=allWeeks.indexOf(wn);
  if(idx<0){idx=0;var mn=Math.abs(allWeeks[0]-wn);allWeeks.forEach(function(w,i){var d=Math.abs(w-wn);if(d<mn){mn=d;idx=i;}});}
  state.weekIdx=idx; updateWeekControls(); renderView();
}
function setView(v){
  state.view=v;
  ['semana','anual','comparativo','servicios'].forEach(function(n){
    var el=document.getElementById('vt'+n.charAt(0).toUpperCase()+n.slice(1));
    if(el) el.className='vtab'+(v===n?' active':'');
  });
  var rb=document.getElementById('rangeBar');
  if(rb) rb.className='range-bar'+(v==='comparativo'?' show':'');
  closeProdPanel();
  renderView();
}
function onRangeChange(){
  var f=parseInt(document.getElementById('fromSlider').value);
  var t=parseInt(document.getElementById('toSlider').value);
  if(f>t){var tmp=f;f=t;t=tmp;}
  state.fromWeek=f; state.toWeek=t;
  updateRangeSliders();
  if(state.view==='comparativo') renderComparativo();
}
function resetRange(){
  var yr=DATA.years[DATA.years.length-1];
  var wks=DATA.weekly_detail.filter(function(r){return r.year===yr;}).map(function(r){return r.week;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort(function(a,b){return a-b;});
  state.toWeek=wks[wks.length-1]||allWeeks[allWeeks.length-1]||52;
  state.fromWeek=wks[wks.length-2]||wks[0]||state.toWeek;
  updateRangeSliders(); renderComparativo();
}

// ── VIEW ROUTER ──
function renderView(){
  _sortState={};
  if     (state.view==='semana')      renderSemana();
  else if(state.view==='anual')       renderAnual();
  else if(state.view==='comparativo') renderComparativo();
  else if(state.view==='servicios')   renderServicios();
  resizeWrap();
}

// ══════════════════════════════════
// VISTA: SEMANA
// ══════════════════════════════════
function renderSemana(){
  var yrs=getActiveYears(), wn=allWeeks[state.weekIdx]||1, sym=state.currency.toUpperCase();
  var head='<tr>';
  head+='<th data-col="yr" style="text-align:left">AÑO</th>';
  head+='<th data-col="sem" data-num="1">SEM</th>';
  head+='<th data-col="cat" style="text-align:left">CATEGORÍA</th>';
  head+='<th data-col="total" data-num="1">TOTAL '+sym+'</th>';
  head+='<th data-col="damt" data-num="1">Δ $</th>';
  head+='<th data-col="dpct" data-num="1">Δ %</th>';
  RANCH_ORDER.forEach(function(r){ head+='<th data-col="r_'+r+'" data-num="1">'+r+'</th>'; });
  head+='</tr>';

  var grandTotal=0, body='';
  yrs.forEach(function(yr,i){
    var prevYr=i>0?yrs[i-1]:null;
    var recs=getWeekDetail(state.cat,wn,yr);
    var agg=sumDetail(recs);
    var col=YEAR_COLORS[yr]||'#888';
    var dAmt='', dPct='';
    if(prevYr){
      var aggP=sumDetail(getWeekDetail(state.cat,wn,prevYr));
      var diff=agg.total-aggP.total;
      var pct=aggP.total>0?(diff/aggP.total*100):null;
      dAmt='<span class="'+(diff>0?'pos':diff<0?'neg':'mut')+'">'+(diff>0?'+':'')+fmt(diff)+'</span>';
      dPct=pct!==null?'<span class="'+(pct>0?'pos':pct<0?'neg':'mut')+'">'+(pct>0?'+':'')+pct.toFixed(1)+'%</span>':'—';
    }
    if(yr===yrs[yrs.length-1]) grandTotal+=agg.total;
    body+='<tr>';
    body+='<td data-col="yr" data-val="'+yr+'"><span style="color:'+col+';font-weight:700">'+yr+'</span></td>';
    body+='<td data-col="sem" data-val="'+wn+'" style="text-align:center">'+wFmt(wn)+'</td>';
    body+='<td data-col="cat" data-val="'+state.cat+'" style="font-weight:700;color:#0d1f3c">'+state.cat+'</td>';
    body+='<td data-col="total" data-val="'+(agg.total||0)+'" style="color:#0d1f3c;font-weight:600">'+fmt(agg.total)+'</td>';
    body+='<td data-col="damt">'+dAmt+'</td>';
    body+='<td data-col="dpct">'+dPct+'</td>';
    RANCH_ORDER.forEach(function(r){
      var v=agg.ranches[r]||0;
      var rc=RANCH_COLORS[r]||'#888';
      body+='<td data-col="r_'+r+'" data-val="'+v+'" style="color:'+(v>0?rc:'#ddd')+'" class="'+(v>0?'clickable':'')+'"'+(v>0?' onclick="onCellClick(\''+yr+'\',\''+wn+'\',\''+wn+'\',\''+r+'\')"':'')+'>'+fmt(v)+'</td>';
    });
    body+='</tr>';
  });
  body+='<tr class="total"><td colspan="3">TOTAL</td><td>'+fmt(grandTotal)+'</td><td colspan="'+(2+RANCH_ORDER.length)+'"></td></tr>';
  setTable(head, body, fmt(grandTotal)+' '+sym);
}

// ══════════════════════════════════
// VISTA: ANUAL
// ══════════════════════════════════
function renderAnual(){
  var yrs=getActiveYears(), sym=state.currency.toUpperCase();
  var head='<tr>';
  head+='<th data-col="yr" style="text-align:left">AÑO</th>';
  head+='<th data-col="cat" style="text-align:left">CATEGORÍA</th>';
  head+='<th data-col="total" data-num="1">TOTAL '+sym+'</th>';
  head+='<th data-col="damt" data-num="1">Δ $</th>';
  head+='<th data-col="dpct" data-num="1">Δ %</th>';
  RANCH_ORDER.forEach(function(r){ head+='<th data-col="r_'+r+'" data-num="1">'+r+'</th>'; });
  head+='</tr>';

  function getYrAgg(yr){
    var d=(DATA.summary[state.cat]||{})[yr]||{usd:0,mxn:0,ranches:{},ranches_mxn:{}};
    return {total:state.currency==='usd'?d.usd:d.mxn, ranches:state.currency==='usd'?d.ranches:d.ranches_mxn};
  }

  var grandTotal=0, body='';
  yrs.forEach(function(yr,i){
    var prevYr=i>0?yrs[i-1]:null;
    var agg=getYrAgg(yr);
    var col=YEAR_COLORS[yr]||'#888';
    var dAmt='', dPct='';
    if(prevYr){
      var aggP=getYrAgg(prevYr);
      var diff=agg.total-aggP.total;
      var pct=aggP.total>0?(diff/aggP.total*100):null;
      dAmt='<span class="'+(diff>0?'pos':diff<0?'neg':'mut')+'">'+(diff>0?'+':'')+fmt(diff)+'</span>';
      dPct=pct!==null?'<span class="'+(pct>0?'pos':pct<0?'neg':'mut')+'">'+(pct>0?'+':'')+pct.toFixed(1)+'%</span>':'—';
    }
    if(yr===yrs[yrs.length-1]) grandTotal+=agg.total;
    body+='<tr>';
    body+='<td data-col="yr" data-val="'+yr+'"><span style="color:'+col+';font-weight:700">'+yr+'</span></td>';
    body+='<td data-col="cat" style="font-weight:700;color:#0d1f3c">'+state.cat+'</td>';
    body+='<td data-col="total" data-val="'+(agg.total||0)+'" style="color:#0d1f3c;font-weight:600">'+fmt(agg.total)+'</td>';
    body+='<td data-col="damt">'+dAmt+'</td>';
    body+='<td data-col="dpct">'+dPct+'</td>';
    RANCH_ORDER.forEach(function(r){
      var v=(agg.ranches||{})[r]||0;
      var rc=RANCH_COLORS[r]||'#888';
      body+='<td data-col="r_'+r+'" data-val="'+v+'" style="color:'+(v>0?rc:'#ddd')+'">'+fmt(v)+'</td>';
    });
    body+='</tr>';
  });
  body+='<tr class="total"><td colspan="2">TOTAL</td><td>'+fmt(grandTotal)+'</td><td colspan="'+(2+RANCH_ORDER.length)+'"></td></tr>';
  setTable(head, body, fmt(grandTotal)+' '+sym);
}

// ══════════════════════════════════
// VISTA: COMPARATIVO
// ══════════════════════════════════
function renderComparativo(){
  var f=state.fromWeek, t=state.toWeek;
  var yrs=getActiveYears(), sym=state.currency.toUpperCase();
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});

  // Precargar datos por año/semana
  var weekData={};
  yrs.forEach(function(yr){
    weekData[yr]={};
    rangeWeeks.forEach(function(w){
      var recs=DATA.weekly_detail.filter(function(r){return r.categoria===state.cat&&r.year===yr&&r.week===w;});
      if(recs.length) weekData[yr][w]=aggregateRecs(recs);
    });
  });

  // Totales por año
  var byYear={};
  yrs.forEach(function(yr){
    var recs=DATA.weekly_detail.filter(function(r){return r.categoria===state.cat&&r.year===yr&&r.week>=f&&r.week<=t;});
    if(recs.length) byYear[yr]=aggregateRecs(recs);
  });

  // Cabecera
  var head='<tr>';
  head+='<th style="text-align:left">Semana</th>';
  head+='<th style="text-align:left">Fecha</th>';
  head+='<th>Total '+sym+'</th>';
  head+='<th>Δ vs ant.</th>';
  RANCH_ORDER.forEach(function(r){ head+='<th>'+r+'</th>'; });
  head+='</tr>';

  var body='';
  var grandTotal=0;

  // Modo: Año → filas de semana
  yrs.forEach(function(yr, yi){
    var col=YEAR_COLORS[yr]||'#888';
    var yearTotal=byYear[yr]?(state.currency==='usd'?byYear[yr].usd:byYear[yr].mxn):0;
    grandTotal+=yearTotal;

    // Fila de encabezado del año
    var prevYrD=yi>0?byYear[yrs[yi-1]]:null;
    var prevYrVal=prevYrD?(state.currency==='usd'?prevYrD.usd:prevYrD.mxn):null;
    var yDiff=prevYrVal!==null?yearTotal-prevYrVal:null;
    var yPct=(yDiff!==null&&prevYrVal!==0)?(yDiff/prevYrVal*100):null;
    var yCls=yDiff===null?'mut':yDiff>0?'pos':'neg';
    var ySign=yDiff!==null&&yDiff>0?'+':'';

    var ranchHdrCells=RANCH_ORDER.map(function(r){
      var d=byYear[yr]; if(!d) return '<td class="mut">—</td>';
      var src=state.currency==='usd'?d.ranches:d.ranches_mxn;
      var v=src[r]||0;
      return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888'):'#ddd')+';font-size:.68rem">'+fmt(v)+'</td>';
    }).join('');

    body+='<tr class="grp">';
    body+='<td colspan="2" style="color:'+col+'">'+yr+'</td>';
    body+='<td style="color:'+col+';font-weight:700">'+fmt(yearTotal)+'</td>';
    body+='<td class="'+yCls+'" style="font-size:.67rem">'+(yDiff!==null?ySign+fmt(yDiff):'—')+(yPct!==null?'<br><span style="font-size:.60rem">'+ySign+yPct.toFixed(1)+'%</span>':'')+'</td>';
    body+=ranchHdrCells;
    body+='</tr>';

    // Filas de semanas dentro del año
    var prevWkVal=null;
    rangeWeeks.forEach(function(w){
      var d=weekData[yr][w];
      var val=d?(state.currency==='usd'?d.usd:d.mxn):0;
      var dCell=deltaHtml(val,prevWkVal);
      if(val>0) prevWkVal=val;

      var ranchCells=RANCH_ORDER.map(function(r){
        if(!d) return '<td class="mut">—</td>';
        var src=state.currency==='usd'?d.ranches:d.ranches_mxn;
        var v=src[r]||0;
        var rc=RANCH_COLORS[r]||'#888';
        var click=v>0?' class="clickable" onclick="onCellClick(\''+yr+'\',\''+w+'\',\''+w+'\',\''+r+'\')"':'';
        return '<td style="color:'+(v>0?rc:'#ddd')+'"'+click+'>'+fmt(v)+'</td>';
      }).join('');

      var totalClick=val>0?' class="clickable" onclick="onCellClick(\''+yr+'\',\''+w+'\',\''+w+'\',\'\')"':'';
      body+='<tr>';
      body+='<td style="color:'+col+';font-weight:600;padding-left:20px">'+String(yr).slice(2)+String(w).padStart(2,'0')+'</td>';
      body+='<td style="color:#999;font-size:.67rem">'+fmtMes(d&&d.date_range)+'</td>';
      body+='<td style="color:'+col+';font-weight:'+(val>0?'600':'400')+'"'+totalClick+'>'+fmt(val)+'</td>';
      body+=dCell+ranchCells;
      body+='</tr>';
    });

    // Subtotal año
    body+='<tr class="sub">';
    body+='<td colspan="2">Subtotal '+yr+'</td>';
    body+='<td>'+fmt(yearTotal)+'</td>';
    body+='<td colspan="'+(1+RANCH_ORDER.length)+'"></td>';
    body+='</tr>';
  });

  // Total general
  body+='<tr class="total">';
  body+='<td colspan="2">TOTAL GENERAL</td>';
  body+='<td>'+fmt(grandTotal)+' '+sym+'</td>';
  body+='<td colspan="'+(1+RANCH_ORDER.length)+'"></td>';
  body+='</tr>';

  setTable(head, body, fmt(grandTotal)+' '+sym);
}

// ══════════════════════════════════
// VISTA: SERVICIOS
// ══════════════════════════════════
function renderServicios(){
  var yrs=getActiveYears(), sym=state.currency.toUpperCase();

  var svRows={};
  if(Array.isArray(DATA.servicios_data)&&DATA.servicios_data.length){
    DATA.servicios_data.forEach(function(r){
      if(!state.activeYears[r.year]) return;
      var subcat=(r.subcat||'').trim(); if(!subcat) return;
      if(!svRows[subcat]) svRows[subcat]={_total:0};
      var src=state.currency==='usd'?(r.usd_ranches||{}):(r.mxn_ranches||{});
      RANCH_ORDER.forEach(function(rn){ var v=src[rn]||0; if(v>0) svRows[subcat][rn]=(svRows[subcat][rn]||0)+v; });
      svRows[subcat]._total+=(state.currency==='usd'?r.usd_total:r.mxn_total)||0;
    });
  } else {
    DATA.weekly_detail.forEach(function(r){
      if(!state.activeYears[r.year]) return;
      if(!r.categoria||!r.categoria.startsWith('SV:')) return;
      var subcat=r.categoria.replace('SV:','');
      if(!svRows[subcat]) svRows[subcat]={_total:0};
      var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
      RANCH_ORDER.forEach(function(rn){ var v=(src||{})[rn]||0; if(v>0) svRows[subcat][rn]=(svRows[subcat][rn]||0)+v; });
      svRows[subcat]._total+=(state.currency==='usd'?r.usd_total:r.mxn_total)||0;
    });
  }

  var grandTotal=Object.keys(svRows).reduce(function(s,k){return s+(svRows[k]._total||0);},0);
  var ordered=SV_SUBCATS.filter(function(sc){return svRows[sc];});
  Object.keys(svRows).forEach(function(sc){if(ordered.indexOf(sc)===-1) ordered.push(sc);});
  ordered.sort(function(a,b){return (svRows[b]._total||0)-(svRows[a]._total||0);});

  var head='<tr>';
  head+='<th data-col="sub" style="text-align:left">SUBCATEGORÍA</th>';
  head+='<th data-col="total" data-num="1">TOTAL '+sym+'</th>';
  head+='<th data-col="pct" data-num="1">% DEL TOTAL</th>';
  RANCH_ORDER.forEach(function(r){ head+='<th data-col="r_'+r+'" data-num="1">'+r+'</th>'; });
  head+='</tr>';

  var body='';
  ordered.forEach(function(sc){
    var d=svRows[sc]||{};
    var tot=d._total||0;
    var pct=grandTotal>0?(tot/grandTotal*100):0;
    var barW=Math.min(pct/100*50,50).toFixed(0);
    body+='<tr>';
    body+='<td data-col="sub" style="font-weight:700;color:#0d1f3c">'+sc+'</td>';
    body+='<td data-col="total" data-val="'+tot+'">'+fmt(tot)+'</td>';
    body+='<td data-col="pct" data-val="'+pct.toFixed(2)+'" style="white-space:nowrap">';
    body+='<div style="display:flex;align-items:center;gap:5px"><div style="width:'+barW+'px;height:6px;background:#0071ce;border-radius:2px;flex-shrink:0"></div><span>'+pct.toFixed(1)+'%</span></div></td>';
    RANCH_ORDER.forEach(function(r){
      var v=d[r]||0;
      var rc=RANCH_COLORS[r]||'#888';
      body+='<td data-col="r_'+r+'" data-val="'+v+'" style="color:'+(v>0?rc:'#ddd')+'">'+fmt(v)+'</td>';
    });
    body+='</tr>';
  });
  body+='<tr class="total"><td>TOTAL</td><td>'+fmt(grandTotal)+'</td><td>100%</td><td colspan="'+RANCH_ORDER.length+'"></td></tr>';
  setTable(head, body, fmt(grandTotal)+' '+sym);
}

// ══════════════════════════════════
// PANEL DE PRODUCTOS (sub-panel click)
// ══════════════════════════════════
function onCellClick(yr, fromWk, toWk, ranch){
  yr=parseInt(yr); fromWk=parseInt(fromWk); toWk=parseInt(toWk);
  var cat=state.cat;
  var isMant=cat==='MANTENIMIENTO', isMatEmp=cat==='MATERIAL DE EMPAQUE';
  var src=isMant?'mp':isMatEmp?'me':'pr';
  var dsMap={pr:DATA.productos,mp:DATA.productos_mp,me:DATA.productos_me};
  var ds=dsMap[src]||{};

  var rows=[];
  for(var wk=fromWk;wk<=toWk;wk++){
    var wkS=((yr%100)*100)+wk, wkL=(yr*100)+wk;
    var wd=ds[wkS]||ds[String(wkS)]||ds[wkL]||ds[String(wkL)];
    if(!wd) continue;
    Object.keys(wd).forEach(function(r){
      if(ranch&&r!==ranch) return;
      var byTipo=wd[r];
      Object.keys(byTipo).forEach(function(tipo){
        (byTipo[tipo]||[]).forEach(function(item){
          rows.push({wk:wkS,rancho:r,tipo:tipo,producto:item[0]||'',unidades:item[1]||'—',gasto:parseFloat(item[2])||0});
        });
      });
    });
  }

  rows.sort(function(a,b){return b.gasto-a.gasto;});
  var total=rows.reduce(function(s,r){return s+r.gasto;},0);

  document.getElementById('prodTitle').textContent = cat+' ▸ '+String(yr).slice(2)+String(fromWk).padStart(2,'0')+(ranch?' · '+ranch:'');
  document.getElementById('prodMeta').textContent  = rows.length+' registros · '+fmt(total);
  document.getElementById('prodPanel').className   = 'show';

  var head='<tr><th style="text-align:left">WK</th><th style="text-align:left">Rancho</th><th style="text-align:left">Tipo</th><th style="text-align:left">Producto</th><th style="text-align:left">Unid.</th><th>Gasto</th></tr>';
  var body=rows.map(function(r){
    var rc=RANCH_COLORS[r.rancho]||'#666';
    return '<tr><td>'+r.wk+'</td><td style="color:'+rc+';font-weight:600">'+r.rancho+'</td><td style="color:#888">'+r.tipo+'</td><td style="color:#0d1f3c">'+r.producto+'</td><td>'+r.unidades+'</td><td>'+fmt(r.gasto)+'</td></tr>';
  }).join('');
  document.getElementById('prodHead').innerHTML=head;
  document.getElementById('prodBody').innerHTML=body||'<tr><td colspan="6" style="text-align:center;color:#aaa;padding:12px">Sin datos de productos</td></tr>';
  resizeWrap();
}
function closeProdPanel(){
  document.getElementById('prodPanel').className='';
  resizeWrap();
}

// ── RESIZE ──
function resizeWrap(){
  var hdr=document.querySelector('.app-hdr');
  var tb=document.querySelector('.toolbar');
  var tabs=document.querySelector('.view-tabs');
  var rb=document.querySelector('.range-bar.show');
  var sb=document.querySelector('.statusbar');
  var pp=document.getElementById('prodPanel');
  var used=0;
  if(hdr)  used+=hdr.offsetHeight;
  if(tb)   used+=tb.offsetHeight;
  if(tabs) used+=tabs.offsetHeight;
  if(rb)   used+=rb.offsetHeight;
  if(sb)   used+=sb.offsetHeight;
  if(pp&&pp.classList.contains('show')) used+=pp.offsetHeight;
  var h=Math.max(document.documentElement.clientHeight-used-2,200);
  var mw=document.getElementById('mainWrap');
  if(mw) mw.style.height=h+'px';
  reportHeight();
}
window.addEventListener('resize', resizeWrap);

function reportHeight(){
  var h=document.getElementById('app');
  window.parent.postMessage({type:'streamlit:setFrameHeight',height:Math.max(h?h.scrollHeight+40:700,700)},'*');
}
var ro=new ResizeObserver(reportHeight);
ro.observe(document.body);
setInterval(reportHeight,500);

window.onerror=function(msg,src,line){
  document.getElementById('loader').innerHTML='<div style="color:#dc2626;padding:20px;font-size:12px"><b>ERROR:</b> '+msg+' (línea '+line+')</div>';
  return true;
};

inicializar();
</script>
</body>
</html>"""

html_final = HTML.replace('__DATA_JSON__', data_json)
components.html(html_final, height=800, scrolling=False)

# ─── Barra inferior: Descarga XLSX + Panel Crear Hoja WK ─────────────────────
st.markdown("""
<style>
  div[data-testid="stSelectbox"] > div { min-width:120px !important; }
  .crear-panel {
    background: #0d1f3c; border-top: 3px solid #0071ce;
    padding: 14px 18px 12px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
  }
  .crear-panel-title { color: rgba(255,255,255,0.55); font-size: 10px; font-family: Arial; text-transform: uppercase; letter-spacing: 0.6px; }
</style>
""", unsafe_allow_html=True)

if "show_crear_panel" not in st.session_state:
    st.session_state.show_crear_panel = False

available_weeks = sorted(
    { str(r["year"] % 100).zfill(2) + str(r["week"]).zfill(2) for r in DATA.get("weekly_detail", []) },
    reverse=True
)

if available_weeks:
    from data_extractor import get_sheet_xlsx
    try:
        from data_extractor import crear_hoja_wk
        _crear_disponible = True
    except ImportError:
        _crear_disponible = False

    col1, col2, col3, col_menu = st.columns([1.2, 1, 5, 0.18])

    with col1:
        selected_wk = st.selectbox("⬇ Descargar hoja WK", options=available_weeks, format_func=lambda c: f"WK{c}", label_visibility="visible")

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Descargar XLSX", key="dl_xlsx"):
            with st.spinner(f"Preparando WK{selected_wk}..."):
                xlsx_bytes = get_sheet_xlsx(selected_wk)
            if xlsx_bytes:
                st.download_button(label=f"💾 WK{selected_wk}.xlsx", data=xlsx_bytes, file_name=f"WK{selected_wk}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx_btn")
            else:
                st.error(f"No se encontró la hoja WK{selected_wk} en el archivo.")

    with col_menu:
        st.markdown("<br>", unsafe_allow_html=True)
        if _crear_disponible:
            if st.button("☰", key="toggle_crear", help="Crear nueva hoja WK en SharePoint"):
                st.session_state.show_crear_panel = not st.session_state.show_crear_panel

    if _crear_disponible and st.session_state.show_crear_panel:
        st.markdown('<div class="crear-panel"><span class="crear-panel-title">➕ Nueva hoja WK en SharePoint</span></div>', unsafe_allow_html=True)
        pc1, pc2, pc3, pc4 = st.columns([1.2, 0.8, 0.8, 4])
        with pc1:
            nuevo_nombre = st.text_input("Nombre de la hoja", placeholder="Ej: WK2518", key="nuevo_wk_nombre", label_visibility="visible").strip().upper()
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
                try:
                    tenant_id     = st.secrets["sharepoint"]["tenant_id"]
                    client_id     = st.secrets["sharepoint"]["client_id"]
                    client_secret = st.secrets["sharepoint"]["client_secret"]
                except KeyError as e:
                    st.error(f"❌ Falta la credencial **{e}** en los secrets de Streamlit.")
                    st.stop()
                with st.spinner(f"Creando hoja {nuevo_nombre} en SharePoint…"):
                    resultado = crear_hoja_wk(nuevo_nombre, tenant_id, client_id, client_secret)
                if resultado.get("ok"):
                    st.success(resultado["mensaje"])
                    st.cache_data.clear()
                else:
                    st.error(f"❌ {resultado['error']}")
