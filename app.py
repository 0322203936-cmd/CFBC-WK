"""
app_mejorado.py
Centro Floricultor de Baja California
Tablas Dinámicas Profesionales — AG Grid Enterprise Features
"""

import json
import base64
import os
import streamlit as st
import streamlit.components.v1 as components

from data_extractor import get_datos

st.set_page_config(
    page_title="CFBC WK · Análisis Dinámico",
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
<title>CFBC — Análisis Dinámico</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-grid.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-theme-alpine.css">
<script src="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/dist/ag-grid-community.min.js"></script>
<style>
:root {
  --navy: #1e3a5f;
  --green: #16a34a;
  --red: #dc2626;
  --amber: #d97706;
  --blue: #2563eb;
  --border: #d0d0d0;
  --mono: 'Consolas','Courier New',monospace;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--mono); font-size: 12px; background: #f0f0f0; overflow-x: hidden; }

/* ── LOADER ───────────────────────────────────────── */
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

/* ── HEADER ───────────────────────────────────────── */
.app-hdr {
  background: var(--navy);
  border-bottom: 3px solid var(--green);
  padding: 8px 12px;
  display: flex;
  align-items: center;
  gap: 12px;
  height: 44px;
}
.hdr-brand {
  color: #fff; font-size: 13px; font-weight: 700;
  letter-spacing: 1.2px; white-space: nowrap;
  padding-right: 12px; border-right: 1px solid rgba(255,255,255,0.2);
}
.hdr-kpis { display: flex; gap: 0; flex: 1; overflow: hidden; }
.hdr-kpi {
  padding: 0 14px;
  border-right: 1px solid rgba(255,255,255,0.12);
  display: flex; align-items: center; gap: 10px;
  white-space: nowrap;
}
.hdr-kpi-label { 
  color: rgba(255,255,255,0.5); 
  font-size: 9px; 
  text-transform: uppercase; 
  letter-spacing: 0.6px; 
}
.hdr-kpi-val { 
  color: #fff; 
  font-size: 13px; 
  font-weight: 700; 
}
.hdr-kpi-delta { font-size: 10px; }
.hdr-kpi-delta.pos { color: #4ade80; }
.hdr-kpi-delta.neg { color: #f87171; }

/* ── TOOLBAR ──────────────────────────────────────── */
.toolbar {
  background: #ebebeb;
  border-bottom: 1px solid var(--border);
  padding: 6px 10px;
  display: flex; align-items: center; gap: 8px;
  flex-wrap: wrap; min-height: 36px;
}
.tb-label {
  font-size: 9px; color: #666;
  text-transform: uppercase; letter-spacing: 0.5px;
  white-space: nowrap; font-weight: 600;
}
.tb-sep { width: 1px; height: 20px; background: #ccc; }
select.tb-sel {
  font-size: 11px; font-family: var(--mono);
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 3px 8px; color: #222; cursor: pointer; height: 24px;
}
select.tb-sel:focus { outline: 2px solid var(--green); outline-offset: -1px; }
.tb-btn {
  font-size: 10px; font-family: var(--mono); font-weight: 700;
  background: #fff; border: 1px solid #bbb; border-radius: 3px;
  padding: 3px 10px; cursor: pointer; height: 24px;
  white-space: nowrap; color: #333; transition: all 0.15s;
}
.tb-btn:hover { background: #ddd; }
.tb-btn.active { background: var(--navy); color: #fff; border-color: var(--navy); }
.tb-btn.green-active { background: var(--green); color: #fff; border-color: var(--green); }
.tb-grp { display: flex; }
.tb-grp .tb-btn { border-radius: 0; border-right-width: 0; }
.tb-grp .tb-btn:first-child { border-radius: 3px 0 0 3px; }
.tb-grp .tb-btn:last-child { border-radius: 0 3px 3px 0; border-right-width: 1px; }

/* ── PIVOT CONTROLS ───────────────────────────────── */
.pivot-bar {
  background: #f8f8f8;
  border-bottom: 1px solid var(--border);
  padding: 6px 10px;
  display: flex; align-items: center; gap: 10px;
  flex-wrap: wrap;
}
.pivot-zone {
  display: flex; align-items: center; gap: 6px;
  padding: 4px 8px;
  background: #fff;
  border: 1px dashed #bbb;
  border-radius: 4px;
  min-height: 28px;
}
.pivot-zone-label {
  font-size: 9px; 
  color: #888;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  font-weight: 600;
}
.pivot-chip {
  font-size: 10px;
  font-family: var(--mono);
  background: var(--navy);
  color: #fff;
  padding: 2px 8px;
  border-radius: 3px;
  cursor: move;
  user-select: none;
}
.pivot-chip.agg {
  background: var(--green);
}

/* ── GRID CONTAINER ───────────────────────────────── */
#gridWrap {
  background: #fff;
  border: 1px solid #d5d5d5;
  margin: 0;
}
#myGrid {
  width: 100%;
  height: 600px;
}

/* ── AG GRID CUSTOM STYLES ────────────────────────── */
.ag-theme-alpine {
  --ag-header-height: 32px;
  --ag-header-foreground-color: #333;
  --ag-header-background-color: #f5f5f5;
  --ag-odd-row-background-color: #fafafa;
  --ag-row-hover-color: #f0f7ff;
  --ag-selected-row-background-color: #e3f2fd;
  --ag-font-family: var(--mono);
  --ag-font-size: 11px;
}

/* Estilos para grupos */
.ag-row-group {
  font-weight: 700 !important;
  background: #f9f9f9 !important;
}
.ag-row-group-indent-1 {
  background: #f5f5f5 !important;
}
.ag-row-group-indent-2 {
  background: #f0f0f0 !important;
}

/* Celdas numéricas */
.ag-cell.num-cell {
  text-align: right !important;
  font-variant-numeric: tabular-nums;
}

/* Totales */
.ag-row-total {
  background: #e8f5e9 !important;
  font-weight: 700 !important;
  border-top: 2px solid var(--green) !important;
}

/* Filtros */
.ag-header-cell-filtered {
  background: #fff3cd !important;
}

/* ── STATUSBAR ────────────────────────────────────── */
.statusbar {
  background: #f8f8f8;
  border-top: 1px solid var(--border);
  padding: 4px 10px;
  display: flex; align-items: center; gap: 12px;
  font-size: 10px; color: #666;
  height: 26px;
}
.status-item {
  display: flex; align-items: center; gap: 6px;
}
.status-label {
  color: #888;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  font-size: 9px;
}
.status-val {
  color: var(--navy);
  font-weight: 700;
}
</style>
</head>
<body>

<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">Cargando análisis dinámico...</div>
</div>

<div id="app" style="display:none;">
  <!-- HEADER -->
  <div class="app-hdr">
    <div class="hdr-brand">CFBC · ANÁLISIS DINÁMICO</div>
    <div class="hdr-kpis" id="kpiBar"></div>
  </div>

  <!-- TOOLBAR -->
  <div class="toolbar">
    <span class="tb-label">Vista:</span>
    <div class="tb-grp">
      <button class="tb-btn active" id="btnVistaNormal">📊 Normal</button>
      <button class="tb-btn" id="btnVistaPivot">🔄 Pivot Table</button>
      <button class="tb-btn" id="btnVistaAgrupada">📁 Agrupada</button>
    </div>
    
    <div class="tb-sep"></div>
    
    <span class="tb-label">Semana:</span>
    <select class="tb-sel" id="selWeek"></select>
    
    <div class="tb-sep"></div>
    
    <span class="tb-label">Rancho:</span>
    <select class="tb-sel" id="selRanch">
      <option value="">Todos</option>
    </select>
    
    <span class="tb-label">Categoría:</span>
    <select class="tb-sel" id="selCategory">
      <option value="">Todas</option>
    </select>
    
    <div class="tb-sep"></div>
    
    <button class="tb-btn" id="btnExpandAll">⊞ Expandir Todo</button>
    <button class="tb-btn" id="btnCollapseAll">⊟ Colapsar Todo</button>
    <button class="tb-btn" id="btnResetFilters">⟲ Reset Filtros</button>
    <button class="tb-btn green-active" id="btnExport">⬇ Exportar Excel</button>
  </div>

  <!-- PIVOT CONTROLS (hidden by default) -->
  <div class="pivot-bar" id="pivotControls" style="display:none;">
    <div class="pivot-zone">
      <span class="pivot-zone-label">Filas:</span>
      <span class="pivot-chip" draggable="true">Rancho</span>
      <span class="pivot-chip" draggable="true">Categoría</span>
    </div>
    <div class="pivot-zone">
      <span class="pivot-zone-label">Columnas:</span>
      <span class="pivot-chip" draggable="true">Semana</span>
    </div>
    <div class="pivot-zone">
      <span class="pivot-zone-label">Valores:</span>
      <span class="pivot-chip agg" draggable="true">Σ USD Total</span>
    </div>
  </div>

  <!-- GRID -->
  <div id="gridWrap">
    <div id="myGrid" class="ag-theme-alpine"></div>
  </div>

  <!-- STATUSBAR -->
  <div class="statusbar">
    <div class="status-item">
      <span class="status-label">Registros:</span>
      <span class="status-val" id="statusRows">0</span>
    </div>
    <div class="status-item">
      <span class="status-label">Filtrados:</span>
      <span class="status-val" id="statusFiltered">0</span>
    </div>
    <div class="status-item">
      <span class="status-label">Total USD:</span>
      <span class="status-val" id="statusTotal">$0</span>
    </div>
    <div class="status-item">
      <span class="status-label">Promedio:</span>
      <span class="status-val" id="statusAvg">$0</span>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════════
// DATOS
// ═══════════════════════════════════════════════════════════
const DATA_B64 = '__DATA_JSON__';
const DATA = JSON.parse(atob(DATA_B64));

let mainGridApi = null;
let currentView = 'normal';

// ═══════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════
function fmt(v) {
  if (typeof v !== 'number') return v;
  return '$' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function pct(v) {
  if (typeof v !== 'number') return '';
  return (v > 0 ? '+' : '') + v.toFixed(1) + '%';
}

// ═══════════════════════════════════════════════════════════
// VALUE FORMATTERS
// ═══════════════════════════════════════════════════════════
const usdFormatter = (params) => {
  if (params.value == null) return '';
  return fmt(params.value);
};

const pctFormatter = (params) => {
  if (params.value == null) return '';
  return pct(params.value);
};

// ═══════════════════════════════════════════════════════════
// COLUMNAS PARA VISTA NORMAL
// ═══════════════════════════════════════════════════════════
const columnDefsNormal = [
  {
    field: 'rancho',
    headerName: 'Rancho',
    width: 140,
    pinned: 'left',
    filter: 'agSetColumnFilter',
    rowGroup: false,
    enableRowGroup: true,
  },
  {
    field: 'categoria',
    headerName: 'Categoría',
    width: 200,
    pinned: 'left',
    filter: 'agSetColumnFilter',
    rowGroup: false,
    enableRowGroup: true,
  },
  {
    field: 'year',
    headerName: 'Año',
    width: 80,
    filter: 'agNumberColumnFilter',
    enableRowGroup: true,
  },
  {
    field: 'week',
    headerName: 'Sem',
    width: 70,
    filter: 'agNumberColumnFilter',
    enableRowGroup: true,
  },
  {
    field: 'usd_total',
    headerName: 'USD Total',
    width: 130,
    type: 'numericColumn',
    valueFormatter: usdFormatter,
    filter: 'agNumberColumnFilter',
    aggFunc: 'sum',
    cellClass: 'num-cell',
  },
  {
    field: 'pct_week',
    headerName: '% Sem',
    width: 100,
    type: 'numericColumn',
    valueFormatter: pctFormatter,
    filter: 'agNumberColumnFilter',
    aggFunc: 'avg',
    cellClass: 'num-cell',
  },
  {
    field: 'pct_ranch',
    headerName: '% Rancho',
    width: 110,
    type: 'numericColumn',
    valueFormatter: pctFormatter,
    filter: 'agNumberColumnFilter',
    aggFunc: 'avg',
    cellClass: 'num-cell',
  },
  {
    field: 'pct_cat',
    headerName: '% Cat',
    width: 100,
    type: 'numericColumn',
    valueFormatter: pctFormatter,
    filter: 'agNumberColumnFilter',
    aggFunc: 'avg',
    cellClass: 'num-cell',
  },
];

// ═══════════════════════════════════════════════════════════
// COLUMNAS PARA VISTA AGRUPADA
// ═══════════════════════════════════════════════════════════
const columnDefsGrouped = [
  {
    field: 'rancho',
    headerName: 'Rancho',
    rowGroup: true,
    hide: true,
  },
  {
    field: 'categoria',
    headerName: 'Categoría',
    rowGroup: true,
    hide: true,
  },
  {
    field: 'year',
    headerName: 'Año',
    width: 80,
  },
  {
    field: 'week',
    headerName: 'Sem',
    width: 70,
  },
  {
    field: 'usd_total',
    headerName: 'USD Total',
    width: 140,
    type: 'numericColumn',
    valueFormatter: usdFormatter,
    aggFunc: 'sum',
    cellClass: 'num-cell',
  },
  {
    field: 'pct_week',
    headerName: '% Sem',
    width: 100,
    type: 'numericColumn',
    valueFormatter: pctFormatter,
    aggFunc: 'avg',
    cellClass: 'num-cell',
  },
];

// ═══════════════════════════════════════════════════════════
// COLUMNAS PARA PIVOT TABLE
// ═══════════════════════════════════════════════════════════
const columnDefsPivot = [
  {
    field: 'rancho',
    headerName: 'Rancho',
    rowGroup: true,
    hide: true,
    enablePivot: true,
  },
  {
    field: 'categoria',
    headerName: 'Categoría',
    rowGroup: true,
    hide: true,
    enablePivot: true,
  },
  {
    field: 'week',
    headerName: 'Semana',
    pivot: true,
    hide: true,
    enablePivot: true,
  },
  {
    field: 'usd_total',
    headerName: 'USD Total',
    aggFunc: 'sum',
    valueFormatter: usdFormatter,
    cellClass: 'num-cell',
  },
];

// ═══════════════════════════════════════════════════════════
// CONFIGURACIÓN BASE DEL GRID
// ═══════════════════════════════════════════════════════════
const gridOptions = {
  columnDefs: columnDefsNormal,
  rowData: DATA.weekly_detail || [],
  
  // Agrupación
  groupDefaultExpanded: 0,
  groupDisplayType: 'multipleColumns',
  groupIncludeTotalFooter: true,
  grandTotalRow: 'bottom',
  
  // Agregaciones
  suppressAggFuncInHeader: false,
  
  // Filtros
  floatingFilter: true,
  
  // Configuración general
  animateRows: true,
  enableRangeSelection: true,
  enableCharts: true,
  
  // Selección
  rowSelection: 'multiple',
  
  // Sidebar
  sideBar: {
    toolPanels: [
      {
        id: 'columns',
        labelDefault: 'Columnas',
        labelKey: 'columns',
        iconKey: 'columns',
        toolPanel: 'agColumnsToolPanel',
        toolPanelParams: {
          suppressRowGroups: false,
          suppressValues: false,
          suppressPivots: false,
          suppressPivotMode: false,
        },
      },
      {
        id: 'filters',
        labelDefault: 'Filtros',
        labelKey: 'filters',
        iconKey: 'filter',
        toolPanel: 'agFiltersToolPanel',
      },
    ],
    defaultToolPanel: '',
  },
  
  // Status bar
  statusBar: {
    statusPanels: [
      {
        statusPanel: 'agTotalAndFilteredRowCountComponent',
        align: 'left',
      },
      {
        statusPanel: 'agAggregationComponent',
        statusPanelParams: {
          aggFuncs: ['sum', 'avg', 'min', 'max', 'count'],
        },
      },
    ],
  },
  
  // Callbacks
  onGridReady: (params) => {
    mainGridApi = params.api;
    params.api.sizeColumnsToFit();
    updateStatusBar();
  },
  
  onFilterChanged: () => {
    updateStatusBar();
  },
  
  onRowDataUpdated: () => {
    updateStatusBar();
  },
};

// ═══════════════════════════════════════════════════════════
// INICIALIZAR
// ═══════════════════════════════════════════════════════════
function inicializar() {
  // Crear grid
  const gridDiv = document.querySelector('#myGrid');
  new agGrid.Grid(gridDiv, gridOptions);
  
  // Poblar selectores
  populateSelectors();
  
  // Event listeners
  setupEventListeners();
  
  // KPIs header
  updateKPIs();
  
  // Ocultar loader
  document.getElementById('loader').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  
  // Resize
  setTimeout(resizeGrid, 100);
}

// ═══════════════════════════════════════════════════════════
// POBLAR SELECTORES
// ═══════════════════════════════════════════════════════════
function populateSelectors() {
  const weeks = [...new Set(DATA.weekly_detail.map(r => r.year + '-W' + String(r.week).padStart(2, '0')))].sort().reverse();
  const ranches = [...new Set(DATA.weekly_detail.map(r => r.rancho))].sort();
  const categories = [...new Set(DATA.weekly_detail.map(r => r.categoria))].sort();
  
  const selWeek = document.getElementById('selWeek');
  const selRanch = document.getElementById('selRanch');
  const selCategory = document.getElementById('selCategory');
  
  weeks.forEach(w => {
    const opt = document.createElement('option');
    opt.value = w;
    opt.textContent = w;
    selWeek.appendChild(opt);
  });
  
  ranches.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r;
    opt.textContent = r;
    selRanch.appendChild(opt);
  });
  
  categories.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    selCategory.appendChild(opt);
  });
}

// ═══════════════════════════════════════════════════════════
// EVENT LISTENERS
// ═══════════════════════════════════════════════════════════
function setupEventListeners() {
  // Vistas
  document.getElementById('btnVistaNormal').addEventListener('click', () => cambiarVista('normal'));
  document.getElementById('btnVistaPivot').addEventListener('click', () => cambiarVista('pivot'));
  document.getElementById('btnVistaAgrupada').addEventListener('click', () => cambiarVista('grouped'));
  
  // Controles
  document.getElementById('btnExpandAll').addEventListener('click', () => {
    if (mainGridApi) mainGridApi.expandAll();
  });
  
  document.getElementById('btnCollapseAll').addEventListener('click', () => {
    if (mainGridApi) mainGridApi.collapseAll();
  });
  
  document.getElementById('btnResetFilters').addEventListener('click', () => {
    if (mainGridApi) {
      mainGridApi.setFilterModel(null);
      document.getElementById('selRanch').value = '';
      document.getElementById('selCategory').value = '';
    }
  });
  
  document.getElementById('btnExport').addEventListener('click', exportToExcel);
  
  // Filtros
  document.getElementById('selRanch').addEventListener('change', applyFilters);
  document.getElementById('selCategory').addEventListener('change', applyFilters);
}

// ═══════════════════════════════════════════════════════════
// CAMBIAR VISTA
// ═══════════════════════════════════════════════════════════
function cambiarVista(vista) {
  currentView = vista;
  
  // Actualizar botones
  document.querySelectorAll('.toolbar .tb-grp .tb-btn').forEach(btn => btn.classList.remove('active'));
  
  if (vista === 'normal') {
    document.getElementById('btnVistaNormal').classList.add('active');
    document.getElementById('pivotControls').style.display = 'none';
    mainGridApi.setColumnDefs(columnDefsNormal);
    mainGridApi.setPivotMode(false);
    mainGridApi.setRowGroupColumns([]);
  } else if (vista === 'grouped') {
    document.getElementById('btnVistaAgrupada').classList.add('active');
    document.getElementById('pivotControls').style.display = 'none';
    mainGridApi.setColumnDefs(columnDefsGrouped);
    mainGridApi.setPivotMode(false);
    mainGridApi.setRowGroupColumns(['rancho', 'categoria']);
    mainGridApi.setGroupDefaultExpanded(1);
  } else if (vista === 'pivot') {
    document.getElementById('btnVistaPivot').classList.add('active');
    document.getElementById('pivotControls').style.display = 'flex';
    mainGridApi.setColumnDefs(columnDefsPivot);
    mainGridApi.setPivotMode(true);
  }
  
  mainGridApi.sizeColumnsToFit();
  updateStatusBar();
}

// ═══════════════════════════════════════════════════════════
// APLICAR FILTROS
// ═══════════════════════════════════════════════════════════
function applyFilters() {
  const ranch = document.getElementById('selRanch').value;
  const category = document.getElementById('selCategory').value;
  
  const filterModel = {};
  
  if (ranch) {
    filterModel.rancho = {
      filterType: 'text',
      type: 'equals',
      filter: ranch,
    };
  }
  
  if (category) {
    filterModel.categoria = {
      filterType: 'text',
      type: 'equals',
      filter: category,
    };
  }
  
  mainGridApi.setFilterModel(filterModel);
}

// ═══════════════════════════════════════════════════════════
// ACTUALIZAR STATUS BAR
// ═══════════════════════════════════════════════════════════
function updateStatusBar() {
  if (!mainGridApi) return;
  
  let totalRows = 0;
  let filteredRows = 0;
  let totalUSD = 0;
  
  mainGridApi.forEachNode((node) => {
    totalRows++;
  });
  
  mainGridApi.forEachNodeAfterFilter((node) => {
    if (node.data) {
      filteredRows++;
      totalUSD += node.data.usd_total || 0;
    }
  });
  
  document.getElementById('statusRows').textContent = totalRows.toLocaleString();
  document.getElementById('statusFiltered').textContent = filteredRows.toLocaleString();
  document.getElementById('statusTotal').textContent = fmt(totalUSD);
  document.getElementById('statusAvg').textContent = filteredRows > 0 ? fmt(totalUSD / filteredRows) : '$0';
}

// ═══════════════════════════════════════════════════════════
// ACTUALIZAR KPIs
// ═══════════════════════════════════════════════════════════
function updateKPIs() {
  const detail = DATA.weekly_detail || [];
  const total = detail.reduce((sum, r) => sum + (r.usd_total || 0), 0);
  const weeks = new Set(detail.map(r => r.year + '-W' + r.week)).size;
  const ranches = new Set(detail.map(r => r.rancho)).size;
  
  const kpiBar = document.getElementById('kpiBar');
  kpiBar.innerHTML = `
    <div class="hdr-kpi">
      <span class="hdr-kpi-label">Total USD</span>
      <span class="hdr-kpi-val">${fmt(total)}</span>
    </div>
    <div class="hdr-kpi">
      <span class="hdr-kpi-label">Semanas</span>
      <span class="hdr-kpi-val">${weeks}</span>
    </div>
    <div class="hdr-kpi">
      <span class="hdr-kpi-label">Ranchos</span>
      <span class="hdr-kpi-val">${ranches}</span>
    </div>
    <div class="hdr-kpi">
      <span class="hdr-kpi-label">Registros</span>
      <span class="hdr-kpi-val">${detail.length.toLocaleString()}</span>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// EXPORTAR A EXCEL
// ═══════════════════════════════════════════════════════════
function exportToExcel() {
  if (!mainGridApi) return;
  
  const date = new Date().toISOString().split('T')[0];
  const filename = `CFBC_Analisis_${currentView}_${date}.xlsx`;
  
  mainGridApi.exportDataAsExcel({
    fileName: filename,
    sheetName: 'Datos',
    exportMode: 'xlsx',
    allColumns: true,
  });
}

// ═══════════════════════════════════════════════════════════
// RESIZE
// ═══════════════════════════════════════════════════════════
function resizeGrid() {
  const hdr = document.querySelector('.app-hdr');
  const toolbar = document.querySelector('.toolbar');
  const pivot = document.querySelector('.pivot-bar');
  const status = document.querySelector('.statusbar');
  
  let used = 0;
  if (hdr) used += hdr.offsetHeight;
  if (toolbar) used += toolbar.offsetHeight;
  if (pivot && pivot.style.display !== 'none') used += pivot.offsetHeight;
  if (status) used += status.offsetHeight;
  
  const available = document.documentElement.clientHeight - used - 4;
  const h = Math.max(available, 400);
  document.getElementById('myGrid').style.height = h + 'px';
  if (mainGridApi) mainGridApi.sizeColumnsToFit();
}

window.addEventListener('resize', resizeGrid);

// ═══════════════════════════════════════════════════════════
// HEIGHT REPORTING
// ═══════════════════════════════════════════════════════════
function reportHeight() {
  const appEl = document.getElementById('app');
  const h = appEl ? appEl.scrollHeight + 60 : document.body.scrollHeight + 60;
  window.parent.postMessage({ type: 'streamlit:setFrameHeight', height: Math.max(h, 750) }, '*');
}

const ro = new ResizeObserver(reportHeight);
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
if (typeof agGrid === 'undefined') {
  const checkAG = setInterval(function() {
    if (typeof agGrid !== 'undefined') {
      clearInterval(checkAG);
      inicializar();
    }
  }, 100);
} else {
  inicializar();
}
</script>

</body>
</html>"""

html_final = HTML.replace('__DATA_JSON__', data_json)
components.html(html_final, height=850, scrolling=False)

# ─── Barra inferior: igual que antes ───────────────────────────────────────────
st.markdown("""
<style>
  div[data-testid="stSelectbox"] > div { min-width:120px !important; }
  div[data-testid="stButton"] button[kind="secondary"].menu-btn {
    font-family: monospace; font-size: 14px;
    background: #1e3a5f; color: #fff;
    border: none; border-radius: 4px;
    padding: 2px 10px; height: 38px;
  }
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

if "show_crear_panel" not in st.session_state:
    st.session_state.show_crear_panel = False

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
                    st.cache_data.clear()
                else:
                    st.error(f"❌ {resultado['error']}")
