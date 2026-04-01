# ==============================
# app.py (VERSIÓN COMPLETA CON PIVOT)
# ==============================

import json
import base64
import streamlit as st
import streamlit.components.v1 as components
from data_extractor import get_datos

st.set_page_config(
    page_title="CFBC WK",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    return get_datos()

DATA = load_data()

data_json = base64.b64encode(
    json.dumps(DATA, ensure_ascii=True, default=str).encode('utf-8')
).decode('ascii')

HTML = f"""
<!DOCTYPE html>
<html>
<head>
<script src="https://cdn.jsdelivr.net/npm/ag-grid-community/dist/ag-grid-community.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community/styles/ag-grid.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community/styles/ag-theme-alpine.css">
</head>
<body>

<div style="padding:10px">
  <button onclick="setView('normal')">Normal</button>
  <button onclick="setView('pivot')">Pivot</button>
</div>

<div id="grid" class="ag-theme-alpine" style="height:80vh;width:100%"></div>

<script>
var DATA = JSON.parse(atob("{data_json}"));
var gridApi;

var gridOptions = {{
  columnDefs: [],
  rowData: [],

  defaultColDef: {{
    sortable: true,
    filter: true,
    resizable: true,
    enableRowGroup: true,
    enablePivot: true,
    enableValue: true
  }},

  sideBar: {{
    toolPanels: ['columns','filters']
  }},

  onGridReady: function(params) {{
    gridApi = params.api;
    renderNormal();
  }}
}};

new agGrid.Grid(document.getElementById('grid'), gridOptions);

function setView(v) {{
  if (v === 'pivot') renderPivot();
  else renderNormal();
}}

function renderNormal() {{
  gridApi.setPivotMode(false);

  gridApi.setColumnDefs([
    {{field:'year'}},
    {{field:'week'}},
    {{field:'categoria'}},
    {{field:'usd_total'}}
  ]);

  gridApi.setRowData(DATA.weekly_detail);
}}

function renderPivot() {{
  gridApi.setPivotMode(true);

  gridApi.setColumnDefs([
    {{field:'year', rowGroup:true, hide:true}},
    {{field:'categoria', rowGroup:true, hide:true}},
    {{field:'week', pivot:true}},
    {{field:'usd_total', aggFunc:'sum'}}
  ]);

  gridApi.setRowData(DATA.weekly_detail);
}}

</script>

</body>
</html>
"""

components.html(HTML, height=900)
