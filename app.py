"""
app.py
Centro Floricultor de Baja California
Streamlit — carga datos con Python, renderiza el HTML original tal cual
"""

import json
import base64
import os
import streamlit as st
import streamlit.components.v1 as components

from data_extractor import get_datos

# Cargar imagen de flores desde el repositorio
_flores_b64 = ""
for _ext in ["jpg","jpeg","png","webp"]:
    _path = os.path.join(os.path.dirname(__file__), f"flores.{_ext}")
    if os.path.exists(_path):
        with open(_path, "rb") as _f:
            _flores_b64 = f"data:image/{_ext};base64," + base64.b64encode(_f.read()).decode()
        break

st.set_page_config(
    page_title="CFBC WK",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Quitar padding de Streamlit para que el HTML ocupe toda la pantalla
st.markdown("""
<style>
  #MainMenu, header, footer { display: none !important; }
  .stApp { background: #f0f4f8; }
  .block-container { padding: 0 !important; max-width: 100% !important; }
  section[data-testid="stSidebar"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
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

# Serializar DATA a JSON para inyectarlo en el HTML
import base64
data_json = base64.b64encode(
    json.dumps(DATA, ensure_ascii=True, default=str).encode('utf-8')
).decode('ascii')

# ─────────────────────────────────────────────
# HTML ORIGINAL — reemplazar google.script.run
# por datos ya cargados desde Python
# ─────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Ejecución Semanal — Comparativo</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/plotly.js/2.27.0/plotly.min.js"></script>
<style>
:root{
  --bg:#f0f4f8;--surface:#ffffff;--surface2:#f8fafc;--border:#e2e8f0;
  --green:#0a7c52;--text:#1e293b;--muted:#64748b;--dim:#94a3b8;
  --gold:#b45309;--red:#dc2626;--blue:#2563eb;
  --navy:#0f2044;--accent-soft:rgba(10,124,82,.08);
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Syne',sans-serif;min-height:100vh;overflow-x:hidden;overflow-y:visible}

.scroll-x{
  position:relative;overflow-x:auto;overflow-y:visible;
  -webkit-overflow-scrolling:touch;scrollbar-width:thin;scrollbar-color:var(--border) transparent;
}
.scroll-x::-webkit-scrollbar{height:5px}
.scroll-x::-webkit-scrollbar-track{background:transparent}
.scroll-x::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:4px}
.scroll-x::-webkit-scrollbar-thumb:hover{background:#94a3b8}
.scroll-fade{position:relative;}
.scroll-fade::after{
  content:'';position:absolute;top:0;right:0;bottom:0;width:40px;
  background:linear-gradient(to right,transparent,var(--surface));
  pointer-events:none;border-radius:0 12px 12px 0;opacity:0;transition:opacity .3s;
}
.scroll-fade.has-overflow::after{opacity:1}
.scroll-hint{
  display:none;align-items:center;gap:4px;font-size:.6rem;
  font-family:'IBM Plex Mono',monospace;color:var(--dim);margin-top:4px;
  animation:pulse-hint 2s ease-in-out infinite;
}
.scroll-hint.show{display:flex}

.prod-cell{cursor:pointer;text-decoration:underline dotted;text-underline-offset:3px}
@keyframes pulse-hint{0%,100%{opacity:.4}50%{opacity:1}}

#loader{position:fixed;inset:0;background:#f0f4f8;z-index:999;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px}
.spin{width:44px;height:44px;border:3px solid #e2e8f0;border-top-color:#0a7c52;border-radius:50%;animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.load-txt{font-family:'IBM Plex Mono',monospace;font-size:.85rem;color:#64748b;font-weight:600}
.load-sub{font-family:'IBM Plex Mono',monospace;font-size:.72rem;color:#94a3b8}

.hdr{background:var(--navy);border-bottom:4px solid var(--green);padding:22px 32px 18px;position:relative;overflow:hidden}
.hdr::before{content:'';position:absolute;top:-80px;right:-60px;width:300px;height:300px;background:radial-gradient(circle,rgba(10,124,82,.18),transparent 70%);pointer-events:none}
.hdr-logo{display:flex;align-items:center;gap:14px;margin-bottom:10px}
.hdr-logo-mark{width:38px;height:38px;background:var(--green);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1rem;font-weight:800;color:#fff;font-family:'Syne',sans-serif;flex-shrink:0}
.hdr-title{font-size:1rem;font-weight:800;letter-spacing:-.3px;color:#ffffff;text-transform:uppercase}
.hdr-sub{font-size:.7rem;color:rgba(255,255,255,.5);margin-top:1px;font-family:'IBM Plex Mono',monospace}
.badge-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.badge{font-family:'IBM Plex Mono',monospace;font-size:.67rem;font-weight:600;padding:3px 10px;border-radius:20px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);color:rgba(255,255,255,.7)}
.badge.muted{background:rgba(255,255,255,.04);border-color:rgba(255,255,255,.08);color:rgba(255,255,255,.4)}

.cat-select-wrap{padding:10px 24px;background:transparent;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;flex-wrap:nowrap;overflow-x:auto;scrollbar-width:none}
.cat-select-wrap::-webkit-scrollbar{display:none}
.cat-select-label{font-size:.65rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace;white-space:nowrap}
.cat-select-outer{position:relative;display:flex;align-items:center;flex:1;min-width:160px;max-width:400px}
.cat-select{width:100%;background:var(--surface2);border:1.5px solid var(--border);border-radius:10px;color:var(--text);font-family:'IBM Plex Mono',monospace;font-size:.78rem;font-weight:600;padding:8px 36px 8px 14px;cursor:pointer;appearance:none;-webkit-appearance:none;outline:none;transition:border-color .2s}
.cat-select:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(10,124,82,.1)}
.cat-select option{background:var(--surface);color:var(--text)}
.cat-arrow{position:absolute;right:12px;color:var(--green);font-size:.75rem;pointer-events:none}
.cat-count{font-size:.65rem;font-family:'IBM Plex Mono',monospace;color:var(--muted);white-space:nowrap;background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:4px 8px}
.main{padding:16px 24px;display:grid;gap:16px;min-width:0}

.view-tabs{display:flex;gap:0;background:var(--surface);border-bottom:2px solid var(--border)}
.view-tab{flex:1;max-width:220px;padding:11px 20px;font-size:.78rem;font-weight:700;cursor:pointer;border:none;border-bottom:2px solid transparent;margin-bottom:-2px;color:var(--muted);background:transparent;transition:all .2s;font-family:'Syne',sans-serif;letter-spacing:.3px}
.view-tab:hover{color:var(--text);background:var(--accent-soft)}
.view-tab.active{color:var(--green);border-bottom-color:var(--green);background:var(--accent-soft)}

.ctrl-bar{display:flex;align-items:center;gap:14px;flex-wrap:nowrap;padding:12px 32px;background:var(--surface);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:50;overflow-x:auto;scrollbar-width:none;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.ctrl-bar::-webkit-scrollbar{display:none}
.ctrl-label{font-size:.68rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace;white-space:nowrap}
.toggle-group{display:flex;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;overflow:hidden;flex-shrink:0}
.toggle-btn{padding:6px 16px;font-size:.78rem;font-weight:700;cursor:pointer;transition:all .2s;background:transparent;border:none;color:var(--muted);font-family:'Syne',sans-serif;white-space:nowrap}
.toggle-btn.active{background:var(--green);color:#fff}
.year-chips{display:flex;gap:5px;flex-wrap:nowrap;flex-shrink:0}
.yr-chip{padding:4px 10px;font-size:.72rem;font-weight:700;border-radius:6px;cursor:pointer;border:1.5px solid rgba(255,255,255,.25);transition:all .2s;font-family:'IBM Plex Mono',monospace;white-space:nowrap;background:rgba(0,0,0,.35);color:#fff;opacity:.6}
.yr-chip.on{opacity:1;border-color:currentColor;background:rgba(0,0,0,.55)}

.week-nav{display:flex;align-items:center;gap:12px;padding:10px 24px;background:var(--surface);border-bottom:1px solid var(--border);overflow-x:auto;scrollbar-width:none;flex-wrap:nowrap}
.week-nav::-webkit-scrollbar{display:none}
.week-nav-btn{width:32px;height:32px;border:1.5px solid var(--border);border-radius:8px;background:var(--surface2);color:var(--muted);cursor:pointer;font-size:.9rem;display:flex;align-items:center;justify-content:center;transition:all .2s;flex-shrink:0}
.week-nav-btn:hover{border-color:var(--green);color:var(--green);background:var(--accent-soft)}
.week-info{display:flex;flex-direction:column;gap:1px;min-width:160px}
.week-num{font-size:.9rem;font-weight:800;font-family:'IBM Plex Mono',monospace;color:var(--navy)}
.week-date{font-size:.65rem;font-family:'IBM Plex Mono',monospace;color:var(--muted)}
.week-slider{flex:1;min-width:120px;max-width:260px;accent-color:var(--green);cursor:pointer;height:3px}
.week-avail{font-size:.62rem;font-family:'IBM Plex Mono',monospace;color:var(--dim)}

.range-ctrl{display:flex;align-items:center;gap:16px;padding:14px 32px;background:var(--surface);border-bottom:1px solid var(--border);overflow-x:auto;scrollbar-width:none;flex-wrap:nowrap}
.range-ctrl::-webkit-scrollbar{display:none}
.range-group{display:flex;flex-direction:column;gap:4px}
.range-label{font-size:.62rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace}
.range-val{font-size:.9rem;font-weight:700;font-family:'IBM Plex Mono',monospace;color:var(--navy)}
.range-slider{width:160px;accent-color:var(--green);cursor:pointer}
.range-sep{font-size:1.2rem;color:var(--dim);padding-top:12px}
.range-badge{font-size:.72rem;font-family:'IBM Plex Mono',monospace;background:var(--accent-soft);border:1px solid rgba(10,124,82,.2);color:var(--green);padding:5px 12px;border-radius:8px;white-space:nowrap;align-self:flex-end;margin-bottom:2px}

.main{padding:24px 32px;display:grid;gap:20px;min-width:0}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:20px;min-width:0}
@media(max-width:900px){
  .row2{grid-template-columns:1fr}
  .ctrl-bar,.cat-select-wrap,.view-tabs,.week-nav,.range-ctrl,.main{padding-left:12px;padding-right:12px}
}

.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px;min-width:0;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.06)}
.card-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;gap:8px}
.card-title{font-size:.75rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace;white-space:nowrap;font-weight:600}
.card-note{font-size:.65rem;color:var(--dim);font-family:'IBM Plex Mono',monospace;white-space:nowrap;background:var(--surface2);border:1px solid var(--border);border-radius:5px;padding:2px 7px}
.chart-wrap{position:relative}
.chart-wrap.tall{height:300px}
.chart-wrap.medium{height:240px}
.chart-wrap.short{height:190px}

.kpi-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.kpi{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:14px 16px;position:relative;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--accent),transparent)}
.kpi-yr{font-size:.63rem;font-family:'IBM Plex Mono',monospace;color:var(--muted);font-weight:600;letter-spacing:1px}
.kpi-val{font-size:1.3rem;font-weight:800;font-family:'IBM Plex Mono',monospace;margin:4px 0 3px;letter-spacing:-.5px}
.kpi-delta{font-size:.67rem;font-family:'IBM Plex Mono',monospace}
.up{color:#0a7c52}.down{color:var(--red)}.flat{color:var(--muted)}

.ranch-grid{display:flex;flex-direction:column;gap:5px;margin-top:4px}
.ranch-row{display:flex;align-items:center;gap:8px}
.ranch-lbl{width:96px;font-size:.7rem;font-family:'IBM Plex Mono',monospace;font-weight:600;flex-shrink:0}
.ranch-bar-outer{flex:1;height:8px;background:var(--bg);border-radius:4px;overflow:hidden}
.ranch-bar-inner{height:100%;border-radius:4px;transition:width .5s ease}
.ranch-usd{width:76px;text-align:right;font-size:.68rem;font-family:'IBM Plex Mono',monospace;color:var(--text);font-weight:600}
.ranch-pct{width:36px;text-align:right;font-size:.65rem;font-family:'IBM Plex Mono',monospace;color:var(--muted)}

.heatmap-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:thin;scrollbar-color:var(--border) transparent;padding-bottom:6px}
.heatmap-wrap::-webkit-scrollbar{height:5px}
.heatmap-wrap::-webkit-scrollbar-track{background:transparent}
.heatmap-wrap::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:4px}
.heatmap{border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:.62rem;width:100%}
.heatmap th{padding:4px 6px;color:var(--muted);font-weight:600;text-align:center;white-space:nowrap;border-bottom:1px solid var(--border)}
.heatmap th.yr-th{text-align:left;color:var(--text);min-width:52px}
.heatmap td{padding:3px 4px;text-align:center;border-radius:4px;cursor:pointer;transition:opacity .15s;min-width:28px}
.heatmap td:hover{opacity:.7;outline:1px solid var(--green)}
.hm-yr-lbl{font-weight:700;text-align:left;color:var(--muted);padding-right:8px}

.data-table{width:100%;border-collapse:collapse;font-size:.74rem}
.data-table th{padding:9px 10px;text-align:left;color:var(--muted);font-family:'IBM Plex Mono',monospace;font-size:.62rem;letter-spacing:.8px;border-bottom:2px solid var(--border);text-transform:uppercase;font-weight:700;white-space:nowrap;background:var(--surface2)}
.data-table td{padding:8px 10px;border-bottom:1px solid var(--border);font-family:'IBM Plex Mono',monospace;white-space:nowrap;color:var(--text)}
.data-table tr:hover td{background:var(--accent-soft)}
.yr-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px}
.chg-pos{color:#0a7c52;font-weight:600}.chg-neg{color:var(--red);font-weight:600}.chg-0{color:var(--muted)}
.no-data{display:flex;align-items:center;justify-content:center;height:160px;color:var(--dim);font-family:'IBM Plex Mono',monospace;font-size:.78rem;flex-direction:column;gap:8px}

.tr-group-hdr td{background:#f1f5f9;font-weight:700;border-top:2px solid var(--border);font-size:.72rem}
.tr-group-hdr td:first-child{border-left:3px solid var(--accent)}
.tr-week td{font-size:.71rem}
.tr-week:hover td{background:var(--accent-soft)}
.tr-total td{background:rgba(10,124,82,.05);font-weight:700;border-top:1px solid rgba(10,124,82,.2)}
.tr-total td:first-child{border-left:3px solid rgba(10,124,82,.4)}
.delta-cell{font-size:.66rem;font-family:'IBM Plex Mono',monospace;white-space:nowrap}
.delta-amt{display:block}.delta-pct{display:block;font-size:.6rem;opacity:.8}

.btn-reload{padding:6px 14px;font-size:.72rem;font-weight:700;border:1.5px solid var(--border);border-radius:8px;background:var(--surface);color:var(--muted);cursor:pointer;font-family:'IBM Plex Mono',monospace;transition:all .2s}
.btn-reload:hover{border-color:var(--green);color:var(--green);background:var(--accent-soft)}
.stat-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px}
.stat-box{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.stat-label{font-size:.6rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace}
.stat-val{font-size:1rem;font-weight:700;font-family:'IBM Plex Mono',monospace;margin-top:2px;color:var(--navy)}

.table-scroll-wrap{position:relative}
.scroll-hint{display:none;align-items:center;justify-content:flex-end;gap:4px;font-size:.6rem;font-family:'IBM Plex Mono',monospace;color:var(--green);margin-bottom:4px;animation:pulse-hint 2s ease-in-out infinite;}
.scroll-hint.show{display:flex}

.prod-cell{cursor:pointer;text-decoration:underline dotted;text-underline-offset:3px}

/* ── PANEL PRODUCTOS — Power BI compact ─────────────────────── */
.pnl{display:none;margin:0 24px 16px;background:var(--surface);border:1px solid var(--border);border-radius:4px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.pnl.show{display:block;animation:pnl-in .15s ease}
@keyframes pnl-in{from{opacity:0}to{opacity:1}}

/* Header — barra oscura estilo Power BI */
.pnl-hdr{display:flex;align-items:center;justify-content:space-between;padding:6px 10px 6px 12px;background:rgba(110,81,115,0.88);gap:10px;border-bottom:2px solid #0a7c52}
.pnl-hdr-left{display:flex;align-items:center;gap:8px;min-width:0;flex:1}
.pnl-meta{display:flex;align-items:center;gap:10px;flex-wrap:wrap;flex:1}
.pnl-title{font-size:.73rem;font-weight:700;color:#fff;font-family:'IBM Plex Mono',monospace;white-space:nowrap;letter-spacing:.2px}
.pnl-title span{font-weight:400;color:rgba(255,255,255,.7);font-size:.7rem}
.pnl-sep{color:rgba(255,255,255,.4);font-size:.65rem}
.pnl-meta-stat{font-size:.64rem;color:rgba(255,255,255,.8);font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;white-space:nowrap}
.pnl-meta-stat b{color:#fff;font-weight:600}
.pnl-close{background:transparent;border:1px solid rgba(255,255,255,.35);border-radius:3px;color:rgba(255,255,255,.7);cursor:pointer;font-size:.75rem;padding:3px 9px;font-family:'IBM Plex Mono',monospace;font-weight:600;transition:all .12s;flex-shrink:0;letter-spacing:.3px}
.pnl-close:hover{border-color:#fff;color:#fff}

/* KPI bar inline — compacto */
.pnl-kpis{display:none}

/* Table */
.pnl-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:thin;scrollbar-color:var(--border) transparent}
.pnl-scroll::-webkit-scrollbar{height:3px}
.pnl-scroll::-webkit-scrollbar-thumb{background:#cbd5e1}
.pnl-tbl{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;table-layout:fixed}
.pnl-tbl thead tr{background:#f1f5f9;border-bottom:1px solid #cbd5e1}
.pnl-tbl th{padding:5px 8px;font-size:.58rem;text-transform:uppercase;letter-spacing:.6px;color:#475569;font-weight:700;white-space:nowrap;text-align:right;border-right:1px solid #e2e8f0}
.pnl-tbl th:last-child{border-right:none}
.pnl-tbl th:first-child{text-align:left;width:90px}
.pnl-tbl th:nth-child(2){text-align:left;width:auto}
.pnl-tbl th:nth-child(3){width:62px}
.pnl-tbl th:nth-child(4){width:96px}
.pnl-tbl tbody tr{border-bottom:1px solid #f1f5f9;transition:background .08s}
.pnl-tbl tbody tr:nth-child(even) td{background:#fafafa}
.pnl-tbl tbody tr:hover td{background:#eff6f2!important}
.pnl-tbl td{padding:3px 8px;font-size:.66rem;color:var(--text);text-align:right;vertical-align:middle;font-variant-numeric:tabular-nums;border-right:1px solid #f1f5f9;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pnl-tbl td:last-child{border-right:none}
.pnl-tbl td:first-child{text-align:left;color:#64748b;font-size:.59rem}
.pnl-tbl td:nth-child(2){text-align:left;font-weight:500;color:var(--text)}
.pnl-tbl td.qty{color:#475569;font-size:.63rem}
.pnl-tbl td.amt{font-weight:600;font-size:.66rem}
.pnl-tbl td.amt.pos{color:#0a7c52}
.pnl-tbl td.amt.neg{color:#dc2626}

/* Group header rows */
.pnl-grp td{padding:3px 8px!important;font-size:.59rem!important;font-weight:700!important;letter-spacing:.4px!important;color:#1e293b!important;background:#e8edf2!important;border-top:1px solid #cbd5e1!important;border-bottom:1px solid #cbd5e1!important;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pnl-grp td:first-child{border-left:3px solid #0a7c52!important;padding-left:5px!important}
.pnl-grp-sub{font-weight:400!important;font-size:.56rem!important;color:#64748b!important;margin-left:5px}
.pnl-grp td.pnl-grp-amt{font-weight:700!important;font-size:.62rem!important;color:#0a7c52!important;text-align:right!important;font-style:normal!important}

/* Total row */
.pnl-tbl tr.pnl-total td{background:#e8f4ee!important;border-top:2px solid #0a7c52;padding:4px 8px!important;font-weight:700!important;font-size:.66rem!important;color:#0f2044!important}
.pnl-tbl tr.pnl-total td.amt{color:#0a7c52!important;font-size:.68rem!important}
.pnl-tbl tr.pnl-total td:first-child{color:#0a7c52!important;letter-spacing:.4px}

/* Footer strip — una sola línea */
.pnl-footer{display:flex;align-items:center;justify-content:flex-end;gap:16px;padding:4px 10px;border-top:1px solid var(--border);background:#f8fafc}
.pnl-footer-stat{display:flex;align-items:center;gap:5px}
.pnl-footer-stat-lbl{font-size:.58rem;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);font-family:'IBM Plex Mono',monospace}
.pnl-footer-stat-val{font-size:.67rem;font-weight:700;font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;color:#0a7c52}

/* Empty state */
.pnl-empty{padding:20px 16px;text-align:center;color:var(--dim);font-family:'IBM Plex Mono',monospace;font-size:.67rem;line-height:1.7}
.pnl-empty-ico{width:28px;height:28px;margin:0 auto 8px;opacity:.2;stroke:var(--muted);fill:none;stroke-width:1.5;stroke-linecap:round}
.no-prod{font-size:.68rem;font-family:'IBM Plex Mono',monospace;color:var(--dim);padding:16px 0;text-align:left}

/* ── PANEL SERVICIOS ───────────────────────────────────────── */
.serv-panel{display:none;margin:0 24px 16px;background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)}
.serv-panel.show{display:block;animation:pnl-in .15s ease}
.serv-panel-hdr{display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:rgba(10,124,82,.12);border-bottom:2px solid var(--green);gap:12px}
.serv-panel-title{font-size:.8rem;font-weight:800;color:var(--navy);font-family:'Syne',sans-serif;letter-spacing:.2px}
.serv-kpi-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;padding:14px 16px;background:var(--surface2);border-bottom:1px solid var(--border)}
.serv-kpi{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px}
.serv-kpi-label{font-size:.6rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-family:'IBM Plex Mono',monospace}
.serv-kpi-val{font-size:1.1rem;font-weight:800;font-family:'IBM Plex Mono',monospace;color:var(--green);margin-top:4px}
.serv-kpi-sub{font-size:.63rem;color:var(--dim);font-family:'IBM Plex Mono',monospace;margin-top:2px}
.serv-ranch-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px;padding:14px 16px}
.serv-ranch-card{border:1px solid var(--border);border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.serv-ranch-hdr{padding:10px 12px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border)}
.serv-ranch-name{font-size:.78rem;font-weight:800;font-family:'Syne',sans-serif;color:#fff}
.serv-ranch-total{font-size:.85rem;font-weight:800;font-family:'IBM Plex Mono',monospace;color:#fff}
.serv-subcat-list{padding:8px 12px}
.serv-subcat-row{display:flex;align-items:center;gap:8px;padding:4px 0;border-bottom:1px solid var(--border);font-size:.68rem;font-family:'IBM Plex Mono',monospace}
.serv-subcat-row:last-child{border-bottom:none}
.serv-subcat-icon{width:16px;height:16px;border-radius:4px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:.6rem}
.serv-subcat-name{flex:1;color:var(--text)}
.serv-subcat-val{font-weight:600;color:var(--navy)}
.serv-total-row{display:flex;justify-content:space-between;padding:8px 12px;background:var(--accent-soft);border-top:2px solid rgba(10,124,82,.2);font-family:'IBM Plex Mono',monospace;font-size:.72rem;font-weight:800;color:var(--green)}
.serv-bottom{display:grid;grid-template-columns:1fr 1fr;gap:14px;padding:14px 16px;border-top:1px solid var(--border)}
@media(max-width:700px){.serv-bottom{grid-template-columns:1fr}.serv-ranch-grid{grid-template-columns:1fr}}

</style>
</head>
<body>

<div id="loader">
  <div class="spin"></div>
  <div class="load-txt">Cargando datos…</div>
  <div class="load-sub">Procesando hojas semanales</div>
</div>

<div id="app" style="display:none">



<div style="position:relative;border-bottom:2px solid #c8e6d8;overflow:hidden;min-height:70px;">
  <!-- Imagen de fondo -->
  <img src="__FLORES_IMG__" alt="" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;object-position:center 40%;display:block;"
    onerror="this.src='https://images.unsplash.com/photo-1457530378978-8bac673b8062?w=1600&h=200&fit=crop&q=85'"/>
  <!-- Overlay suave para legibilidad -->
  <div style="position:absolute;inset:0;background:rgba(255,255,255,0.35);"></div>
  <!-- Controles encima -->
  <div class="cat-select-wrap" style="position:relative;z-index:2;border-bottom:none;padding:12px 24px;">
    <!-- Label Categoría -->
    <span class="cat-select-label" style="background:rgba(110,81,115,0.85);color:#fff;padding:3px 8px;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.18);">Categoría</span>
    <!-- Selector 1 -->
    <div class="cat-select-outer" style="background:rgba(110,81,115,0.88);border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.2);">
      <select class="cat-select" id="catSelect" onchange="selectCat(this.value)" style="background:transparent;font-weight:700;color:#fff;"></select>
      <span class="cat-arrow" style="color:#fff;">▾</span>
    </div>
    <!-- Selector 2 -->
    <div class="cat-select-outer" style="background:rgba(110,81,115,0.88);border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.2);">
      <select class="cat-select" id="moSelect" style="background:transparent;font-weight:700;color:#fff;">
        <option value="" disabled selected>COSTO DE MANO DE OBRA</option>
      </select>
      <span class="cat-arrow" style="color:#fff;">▾</span>
    </div>
    <!-- Selector 3 -->
    <div class="cat-select-outer" style="background:rgba(110,81,115,0.88);border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.2);">
      <select class="cat-select" id="servSelect" onchange="selectServ(this.value)" style="background:transparent;font-weight:700;color:#fff;">
        <option value="" disabled selected>COSTO DE SERVICIOS</option>
      </select>
      <span class="cat-arrow" style="color:#fff;">▾</span>
    </div>
    <!-- Contador -->
    <span class="cat-count" id="catCount" style="display:none;"></span>
    <!-- Controles derecha -->
    <div style="margin-left:auto;display:flex;align-items:center;gap:10px;flex-shrink:0;background:rgba(110,81,115,0.85);padding:6px 14px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.18);">
      <span class="ctrl-label" style="color:#f0e8f2;">Moneda</span>
      <div class="toggle-group" style="border-color:rgba(255,255,255,.3);">
        <button class="toggle-btn active" id="btnUSD" onclick="setCurrency('usd')">USD $</button>
        <button class="toggle-btn"        id="btnMXN" onclick="setCurrency('mxn')">MXN $</button>
      </div>
      <span class="ctrl-label" style="margin-left:4px;color:#f0e8f2;">Años</span>
      <div class="year-chips" id="yearChips" style="gap:4px;"></div>
    </div>
  </div>
</div>

<div class="view-tabs">
  <button class="view-tab"        id="vtAnual"     onclick="setView('anual')">📊 Por Año</button>
  <button class="view-tab active" id="vtSemana"    onclick="setView('semana')">📆 Por Semana</button>
  <button class="view-tab"        id="vtTendencia" onclick="setView('tendencia')">📈 Tendencia & Rango</button>
</div>

<!-- VIEW: ANUAL -->
<div id="viewAnual" style="display:none">
  <div class="main">
    <div class="kpi-strip" id="kpiStrip"></div>
    <div class="row2">
      <div class="card">
        <div class="card-hdr"><span class="card-title">Comparativo Anual</span><span class="card-note" id="barNote">USD</span></div>
        <div class="chart-wrap tall"><div id="chartBar"></div></div>
      </div>
      <div class="card">
        <div class="card-hdr"><span class="card-title">Desglose por Rancho</span><span class="card-note" id="ranchNote">USD</span></div>
        <div class="toggle-group" style="margin-bottom:12px;display:inline-flex">
          <button class="toggle-btn active" id="ranchAll" onclick="setRanchYear('all')" style="font-size:.68rem;padding:4px 10px">TODOS</button>
          <span id="ranchYrBtns" style="display:contents"></span>
        </div>
        <div id="ranchBars" class="ranch-grid"></div>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr"><span class="card-title">Tendencia Semanal (USD) — Años superpuestos</span><span class="card-note">línea = 1 año</span></div>
      <div class="chart-wrap tall"><div id="chartLine"></div></div>
    </div>
    <div class="row2">
      <div class="card">
        <div class="card-hdr"><span class="card-title">Barras Apiladas por Rancho</span><span class="card-note" id="stackNote">USD</span></div>
        <div class="chart-wrap medium"><div id="chartStack"></div></div>
      </div>
      <div class="card">
        <div class="card-hdr"><span class="card-title">Tabla Resumen Anual</span><span class="card-note" id="tableNote">USD · Δ anual</span></div>
        <div class="table-scroll-wrap">
          <div class="scroll-hint show" id="hintAnual">← desliza →</div>
          <div class="scroll-x scroll-fade" id="wrapAnual">
            <table class="data-table">
              <thead><tr>
                <th>Año</th><th>Total</th><th>Δ vs ant.</th>
                <th>Prop-RM</th><th>PosCo-RM</th><th>Campo-RM</th><th>Isabela</th><th>HOOPS</th><th>Cecilia</th><th>Cecilia 25</th><th>Christina</th>
              </tr></thead>
              <tbody id="tableBody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- VIEW: POR SEMANA -->
<div id="viewSemana">
  <div class="week-nav">
    <button class="week-nav-btn" onclick="prevWeek()">◀</button>
    <div class="week-info">
      <div class="week-num" id="weekNumLabel">Semana W—</div>
      <div class="week-date" id="weekDateLabel">—</div>
    </div>
    <input type="range" class="week-slider" id="weekSlider" min="1" max="52" value="1" oninput="onWeekSlider(this.value)">
    <button class="week-nav-btn" onclick="nextWeek()">▶</button>
    <span class="week-avail" id="weekAvail"></span>
  </div>
  <div class="main">
    <div class="card" id="swTableCard">
      <div class="card-hdr"><span class="card-title">Tabla — Misma Semana, Distintos Años</span><span class="card-note" id="swTableNote">USD</span></div>
      <div class="table-scroll-wrap">
        <div class="scroll-hint show" id="hintSemana">← desliza →</div>
        <div class="scroll-x scroll-fade" id="wrapSemana">
          <table class="data-table">
            <thead><tr>
              <th>Año</th><th>Semana</th><th>Fecha</th><th>Total USD</th><th>Δ vs mismo año ant.</th>
              <th>Prop-RM</th><th>PosCo-RM</th><th>Campo-RM</th><th>Isabela</th><th>Cecilia</th><th>Cecilia 25</th><th>Christina</th>
            </tr></thead>
            <tbody id="swTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- VIEW: TENDENCIA / RANGO -->
<div id="viewTendencia" style="display:none">
  <div class="range-ctrl">
    <span class="range-label">Desde</span>
    <span class="range-val" id="fromWeekLabel">W01</span>
    <input type="range" class="range-slider" id="fromSlider" min="1" max="52" value="1" oninput="onRangeChange()">
    <span class="range-sep">→</span>
    <span class="range-label">Hasta</span>
    <span class="range-val" id="toWeekLabel">W52</span>
    <input type="range" class="range-slider" id="toSlider" min="1" max="52" value="52" oninput="onRangeChange()">
    <span class="range-badge" id="rangeBadge">W01 → W52 · 52 semanas</span>
    <button class="btn-reload" style="margin-left:auto" onclick="resetRange()">↺ Reset</button>
  </div>
  <div class="main">
    <div class="stat-row" id="rangeStats"></div>
    <div class="card">
      <div class="card-hdr"><span class="card-title">Tabla Desglose por Semana</span><span class="card-note" id="rangeTableNote">USD · cada fila = 1 semana</span></div>
      <div style="display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center">
        <span style="font-size:.65rem;font-family:'IBM Plex Mono',monospace;color:var(--muted);text-transform:uppercase;letter-spacing:1px">Ver por</span>
        <div class="toggle-group">
          <button class="toggle-btn active" id="rtgYear" onclick="setRangeTableGroup('year')" style="font-size:.7rem;padding:5px 12px">Año → Semana</button>
          <button class="toggle-btn"        id="rtgWeek" onclick="setRangeTableGroup('week')" style="font-size:.7rem;padding:5px 12px">Semana → Año</button>
        </div>
        <span style="font-size:.62rem;font-family:'IBM Plex Mono',monospace;color:var(--dim)" id="rangeTableSub"></span>
      </div>
      <div class="table-scroll-wrap">
        <div class="scroll-hint show" id="hintRange">← desliza →</div>
        <div class="scroll-x scroll-fade" id="wrapRange">
          <table class="data-table" id="rangeDetailTable">
            <thead id="rangeDetailHead"></thead>
            <tbody id="rangeTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr"><span class="card-title">Tendencia Semanal en el Rango — Años Superpuestos</span><span class="card-note">USD · línea = 1 año</span></div>
      <div class="chart-wrap tall"><div id="chartRangeLine"></div></div>
    </div>
    <div class="row2">
      <div class="card">
        <div class="card-hdr"><span class="card-title">Total en el Rango por Año</span><span class="card-note" id="rangeBarNote">USD</span></div>
        <div class="chart-wrap medium"><div id="chartRangeBar"></div></div>
      </div>
      <div class="card">
        <div class="card-hdr"><span class="card-title">Acumulado Semanal</span><span class="card-note">USD · suma corrida por semana</span></div>
        <div class="chart-wrap medium"><div id="chartCumul"></div></div>
      </div>
    </div>

  </div>
</div>
<div id="productosSection" class="pnl">
  <div class="pnl-hdr">
    <div class="pnl-hdr-left">
      <div class="pnl-meta">
        <span class="pnl-title" id="productosTitle">PRODUCTOS</span>
      </div>
    </div>
    <button class="pnl-close" onclick="closeProductos()">✕</button>
  </div>
  <div id="pnlKpis" class="pnl-kpis"></div>
  <div class="pnl-scroll">
    <table class="pnl-tbl">
      <thead>
        <tr>
          <th>Ubicación</th>
          <th>Producto</th>
          <th>Unid.</th>
          <th>Monto</th>
        </tr>
      </thead>
      <tbody id="productosContent"></tbody>
    </table>
  </div>
  <div id="pnlFooter" class="pnl-footer" style="display:none">
    <div class="pnl-footer-stat" id="pnlFooterStats"></div>
  </div>
</div>

<div id="serviciosSection" class="serv-panel">
  <div class="serv-panel-hdr">
    <div class="serv-panel-title" id="servPanelTitle">COSTO DE SERVICIOS</div>
    <button class="pnl-close" onclick="closeServicios()">✕</button>
  </div>
  <div class="serv-kpi-strip" id="servKpiStrip"></div>
  <div class="serv-ranch-grid" id="servRanchGrid"></div>
  <div class="serv-bottom">
    <div>
      <div style="font-size:.68rem;font-family:'IBM Plex Mono',monospace;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Gastos por Rancho</div>
      <div id="servRanchList" class="ranch-grid"></div>
    </div>
    <div><div id="chartServ" style="height:200px"></div></div>
  </div>
</div>

</div><!-- /app -->

<script>
// ═══════════════════════════════════════════
// DATOS INYECTADOS DESDE PYTHON
// ═══════════════════════════════════════════
var _raw = atob('__DATA_JSON__');
var DATA = JSON.parse(_raw);

// ═══════════════════════════════════════════
// ESTADO GLOBAL
// ═══════════════════════════════════════════
var state = {
  cat:'', currency:'usd', activeYears:{}, ranchYear:'all',
  view:'semana', weekIdx:0, fromWeek:1, toWeek:52, servFilter:null
};
var lastProductos = null; // guarda {rancho, tipo, src} para refrescar con semana nueva
var allWeeks = [];
var YEAR_COLORS = {2021:'#67e8f9',2022:'#fde68a',2023:'#86efac',2024:'#c4b5fd',2025:'#6ee7b7',2026:'#fca5a5'};
var RANCH_COLORS = {
  'Prop-RM':'#047857','PosCo-RM':'#1d4ed8','Campo-RM':'#b45309',
  'Isabela':'#7c3aed','HOOPS':'#c2410c','Cecilia':'#be185d',
  'Cecilia 25':'#047857','Christina':'#0369a1','Albahaca-RM':'#6d28d9','Campo-VI':'#64748b'
};
var RANCH_ORDER = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','HOOPS','Cecilia','Cecilia 25','Christina','Albahaca-RM','Campo-VI'];
var KEY_RANCHES = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','Cecilia','Cecilia 25','Christina'];
var charts = {};

function recargar() { window.location.reload(); }

// ═══════════════════════════════════════════
// INICIALIZAR
// ═══════════════════════════════════════════
function inicializar() {
  // Event delegation para celdas de productos
  document.addEventListener('click', function(e){
    var td = e.target.closest('.prod-cell');
    if(!td) return;
    showProductos(td.dataset.r, td.dataset.t, parseInt(td.dataset.w), parseInt(td.dataset.y), td.dataset.src || 'pr');
  });
  var years = DATA.years, cats = DATA.categories;
  var prefCat = 'MATERIAL DE EMPAQUE';
  state.cat = cats.indexOf(prefCat) > -1 ? prefCat : cats[0];
  // Solo mostrar año más reciente al inicio
  state.activeYears = {};
  var latestYr = years[years.length-1];
  state.activeYears[latestYr] = true;

  var wSet = {};
  DATA.weekly_detail.forEach(function(r){ wSet[r.week] = 1; });
  allWeeks = Object.keys(wSet).map(Number).sort(function(a,b){return a-b;});
  state.weekIdx = allWeeks.length - 1;

  // ── Apuntar a la semana más reciente del año más reciente ──
  var latestYear = DATA.years[DATA.years.length-1];
  var weeksOfLatest = DATA.weekly_detail
    .filter(function(r){ return r.year === latestYear; })
    .map(function(r){ return r.week; })
    .filter(function(v,i,a){ return a.indexOf(v)===i; })
    .sort(function(a,b){ return a-b; });
  var curWeek  = weeksOfLatest[weeksOfLatest.length-1] || allWeeks[allWeeks.length-1] || 1;
  // Buscar el índice de esa semana en allWeeks y apuntar ahí
  var curIdx = allWeeks.indexOf(curWeek);
  if(curIdx >= 0) state.weekIdx = curIdx;
  var prevWeek2 = weeksOfLatest[weeksOfLatest.length-2] || weeksOfLatest[0] || curWeek;
  state.fromWeek = prevWeek2;
  state.toWeek   = curWeek;

  // header removido

  buildCatSelect(); buildServSelect(); buildYearChips(); updateWeekSlider(); updateRangeSliders(); renderView();
  document.getElementById('loader').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  var by=document.getElementById('badgeYears');
  if(by) by.textContent='📅 '+DATA.years[0]+' → '+DATA.years[DATA.years.length-1];
  setTimeout(initScrollHints, 100);
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════
function fmt(n) {
  if (n===null||n===undefined||n===0||isNaN(n)) return '—';
  var neg=n<0, s=Math.abs(n);
  return (neg?'-$':'$')+s.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function pct(a,b){ return (!b||b===0)?null:((a-b)/b*100).toFixed(1); }
function destroyChart(id){ var el=document.getElementById(id); if(el) Plotly.purge(el); }
function getAnnualVal(cat,yr){ var d=(DATA.summary[cat]||{})[yr]; if(!d) return 0; return state.currency==='usd'?d.usd:d.mxn; }
function activeYrList(){ return DATA.years.filter(function(y){return state.activeYears[y];}); }
function wFmt(n){ return 'W'+String(n).padStart(2,'0'); }

function getDetail(cat,weekNum,yearNum){
  return DATA.weekly_detail.filter(function(r){
    if(r.categoria!==cat) return false;
    if(weekNum!==undefined&&r.week!==weekNum) return false;
    if(yearNum!==undefined&&r.year!==yearNum) return false;
    return true;
  });
}
function aggregateDetail(records){
  var out={usd:0,mxn:0,ranches:{},ranches_mxn:{},date_range:''};
  records.forEach(function(r){
    out.usd+=r.usd_total; out.mxn+=r.mxn_total;
    if(r.date_range) out.date_range=r.date_range;
    Object.keys(r.usd_ranches).forEach(function(rn){out.ranches[rn]=(out.ranches[rn]||0)+r.usd_ranches[rn];});
    Object.keys(r.mxn_ranches).forEach(function(rn){out.ranches_mxn[rn]=(out.ranches_mxn[rn]||0)+r.mxn_ranches[rn];});
  });
  out.usd=Math.round(out.usd*100)/100; out.mxn=Math.round(out.mxn*100)/100;
  return out;
}
function getWeekByYear(cat,weekNum){
  var res={};
  activeYrList().forEach(function(yr){
    var recs=getDetail(cat,weekNum,yr);
    if(recs.length) res[yr]=aggregateDetail(recs);
  });
  return res;
}
function getRangeByYear(cat,fromW,toW){
  var res={};
  activeYrList().forEach(function(yr){
    if(isCombined(cat)){
      var d=getCombinedRange(fromW,toW,yr);
      if(d) res[yr]=d;
    } else {
      var recs=getDetail(cat,undefined,yr).filter(function(r){return r.week>=fromW&&r.week<=toW;});
      if(!recs.length) return;
      var ag=aggregateDetail(recs);
      ag.weekly={};
      recs.forEach(function(r){ag.weekly[r.week]=(ag.weekly[r.week]||0)+r.usd_total;});
      res[yr]=ag;
    }
  });
  return res;
}
function heatColor(ratio){
  if(!ratio||ratio<=0) return 'transparent';
  var g=Math.round(60+ratio*141), b=Math.round(50+ratio*30), a=0.15+ratio*0.7;
  return 'rgba(0,'+g+','+b+','+a+')';
}

// ═══════════════════════════════════════════
// UI BUILDERS
// ═══════════════════════════════════════════
var CAT_MIRFE = 'FERTILIZANTES';
var CAT_MIPE  = 'DESINFECCION / PLAGUICIDAS';
var CAT_COMBINED = 'MIRFE + MIPE';

function isCombined(cat){ return cat===CAT_MIRFE || cat===CAT_MIPE; }

function buildCatSelect(){
  var el=document.getElementById('catSelect');
  el.innerHTML='<option value="" disabled>COSTO DE MATERIALES</option>' +
  DATA.categories.map(function(c){
    var label = c===CAT_MIRFE ? c+' (MIRFE)' : c===CAT_MIPE ? c+' (MIPE)' : c;
    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+'>'+label+'</option>';
  }).join('');
  document.getElementById('catCount').textContent=(DATA.categories.indexOf(state.cat)+1)+' / '+DATA.categories.length;
}

// Obtener datos combinados MIRFE+MIPE para una semana y año
function getCombinedWeek(weekNum, yr){
  var r1 = getDetail(CAT_MIRFE, weekNum, yr);
  var r2 = getDetail(CAT_MIPE,  weekNum, yr);
  var mirfe = r1.length ? aggregateDetail(r1) : null;
  var mipe  = r2.length ? aggregateDetail(r2) : null;
  if(!mirfe && !mipe) return null;
  // Combined total
  var out = {
    usd: (mirfe?mirfe.usd:0) + (mipe?mipe.usd:0),
    mxn: (mirfe?mirfe.mxn:0) + (mipe?mipe.mxn:0),
    ranches:{}, ranches_mxn:{},
    date_range: (mirfe&&mirfe.date_range)?mirfe.date_range:((mipe&&mipe.date_range)?mipe.date_range:''),
    mirfe: mirfe, mipe: mipe
  };
  return out;
}

// Obtener datos combinados MIRFE+MIPE para un rango
function getCombinedRange(fromW, toW, yr){
  var recs1 = getDetail(CAT_MIRFE, undefined, yr).filter(function(r){return r.week>=fromW&&r.week<=toW;});
  var recs2 = getDetail(CAT_MIPE,  undefined, yr).filter(function(r){return r.week>=fromW&&r.week<=toW;});
  if(!recs1.length && !recs2.length) return null;
  var ag1 = recs1.length ? aggregateDetail(recs1) : {usd:0,mxn:0,ranches:{},ranches_mxn:{}};
  var ag2 = recs2.length ? aggregateDetail(recs2) : {usd:0,mxn:0,ranches:{},ranches_mxn:{}};
  var weekly = {};
  recs1.forEach(function(r){weekly[r.week]=(weekly[r.week]||0)+r.usd_total;});
  recs2.forEach(function(r){weekly[r.week]=(weekly[r.week]||0)+r.usd_total;});
  return {
    usd: ag1.usd + ag2.usd,
    mxn: ag1.mxn + ag2.mxn,
    ranches: ag1.ranches, ranches_mxn: ag1.ranches_mxn,
    weekly: weekly,
    mirfe_usd: ag1.usd, mirfe_mxn: ag1.mxn,
    mipe_usd:  ag2.usd, mipe_mxn:  ag2.mxn
  };
}
function buildYearChips(){
  var el=document.getElementById('yearChips');
  el.innerHTML=DATA.years.map(function(y){
    var col=YEAR_COLORS[y]||'#888', on=state.activeYears[y]?'on':'';
    return '<button class="yr-chip '+on+'" style="color:'+col+';border-color:'+(on?col:'rgba(255,255,255,.25)')+'" onclick="toggleYear('+y+')">'+y+'</button>';
  }).join('');
  var rb=document.getElementById('ranchYrBtns');
  if(rb) rb.innerHTML=DATA.years.map(function(y){
    return '<button class="toggle-btn" id="ranchYr'+y+'" onclick="setRanchYear('+y+')" style="font-size:.68rem;padding:4px 10px">'+y+'</button>';
  }).join('');
}
function updateWeekSlider(){
  if(!allWeeks.length) return;
  var wn=allWeeks[state.weekIdx];
  var sl=document.getElementById('weekSlider');
  sl.min=allWeeks[0]; sl.max=allWeeks[allWeeks.length-1]; sl.value=wn;
  document.getElementById('weekNumLabel').textContent='Semana '+wFmt(wn);
  var recs=DATA.weekly_detail.filter(function(r){return r.week===wn&&r.date_range;});
  recs.sort(function(a,b){return b.year-a.year;});
  document.getElementById('weekDateLabel').textContent=recs.length?recs[0].date_range:'';
  var avail=DATA.years.filter(function(yr){return DATA.weekly_detail.some(function(r){return r.week===wn&&r.year===yr;});});
  document.getElementById('weekAvail').textContent='Disponible en: '+avail.join(', ');
}
function updateRangeSliders(){
  var f=state.fromWeek, t=state.toWeek;
  document.getElementById('fromSlider').value=f;
  document.getElementById('toSlider').value=t;
  document.getElementById('fromWeekLabel').textContent=wFmt(f);
  document.getElementById('toWeekLabel').textContent=wFmt(t);
  var count=0; allWeeks.forEach(function(w){if(w>=f&&w<=t) count++;});
  document.getElementById('rangeBadge').textContent=wFmt(f)+' → '+wFmt(t)+' · '+count+' semanas';
}

// ═══════════════════════════════════════════
// VIEW SWITCHER
// ═══════════════════════════════════════════
function setView(v){
  closeProductos();
  closeServicios();
  state.view=v;
  // Al entrar a tendencia, activar solo 2026
  if(v==='tendencia'){
    state.activeYears={};
    if(DATA.years.indexOf(2026)>-1) state.activeYears[2026]=true;
    else state.activeYears[DATA.years[DATA.years.length-1]]=true;
    buildYearChips();
  }
  ['anual','semana','tendencia'].forEach(function(name){
    document.getElementById('view'+name.charAt(0).toUpperCase()+name.slice(1)).style.display=v===name?'':'none';
    document.getElementById('vt'+name.charAt(0).toUpperCase()+name.slice(1)).classList.toggle('active',v===name);
  });
  renderView();
}
function renderView(){
  if(state.view==='anual')           renderAnual();
  else if(state.view==='semana')     renderSemana();
  else if(state.view==='tendencia')  renderTendencia();
  setTimeout(initScrollHints,80);
}
function selectCat(cat){
  closeProductos();
  state.cat=cat;
  document.getElementById('catCount').textContent=(DATA.categories.indexOf(cat)+1)+' / '+DATA.categories.length;
  renderView();
}
function setCurrency(cur){
  state.currency=cur;
  document.getElementById('btnUSD').classList.toggle('active',cur==='usd');
  document.getElementById('btnMXN').classList.toggle('active',cur==='mxn');
  renderView();
  refreshServSiAbierto();
}
function toggleYear(y){
  var active=DATA.years.filter(function(yr){return state.activeYears[yr];});
  if(state.activeYears[y]&&active.length>1) delete state.activeYears[y];
  else state.activeYears[y]=true;
  buildYearChips(); renderView();
}
function setRanchYear(yr){
  state.ranchYear=yr;
  document.getElementById('ranchAll').classList.toggle('active',yr==='all');
  DATA.years.forEach(function(y){var b=document.getElementById('ranchYr'+y);if(b) b.classList.toggle('active',yr===y);});
  renderRanchBars();
}
function prevWeek(){ if(state.weekIdx>0){state.weekIdx--;updateWeekSlider();renderSemana();refreshProductosSiAbierto();refreshServSiAbierto();} }
function nextWeek(){ if(state.weekIdx<allWeeks.length-1){state.weekIdx++;updateWeekSlider();renderSemana();refreshProductosSiAbierto();refreshServSiAbierto();} }
function onWeekSlider(val){
  var wn=parseInt(val), idx=allWeeks.indexOf(wn);
  if(idx<0){ idx=0; var mn=Math.abs(allWeeks[0]-wn); allWeeks.forEach(function(w,i){var d=Math.abs(w-wn);if(d<mn){mn=d;idx=i;}});}
  state.weekIdx=idx; updateWeekSlider(); renderSemana(); refreshProductosSiAbierto(); refreshServSiAbierto();
}
function refreshServSiAbierto(){
  var section=document.getElementById('serviciosSection');
  if(section && section.style.display!=='none'){
    renderServPanel(state.servFilter);
  }
}

function refreshProductosSiAbierto(){
  var section = document.getElementById('productosSection');
  if(lastProductos && section.style.display !== 'none'){
    var weekNum = allWeeks[state.weekIdx];
    var activeYrs = DATA.years.filter(function(y){ return state.activeYears[y]; }).sort(function(a,b){return b-a;});
    var targetYr = activeYrs[0] || DATA.years[DATA.years.length-1];
    showProductos(lastProductos.rancho, lastProductos.tipo, weekNum, targetYr, lastProductos.src);
  }
}
function onRangeChange(){
  var f=parseInt(document.getElementById('fromSlider').value);
  var t=parseInt(document.getElementById('toSlider').value);
  if(f>t){var tmp=f;f=t;t=tmp;}
  state.fromWeek=f; state.toWeek=t; updateRangeSliders(); renderTendencia();
}
function resetRange(){
  var latestYear = DATA.years[DATA.years.length-1];
  var wks = DATA.weekly_detail.filter(function(r){return r.year===latestYear;}).map(function(r){return r.week;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort(function(a,b){return a-b;});
  state.toWeek   = wks[wks.length-1] || allWeeks[allWeeks.length-1] || 52;
  state.fromWeek = wks[wks.length-2] || wks[0] || state.toWeek;
  updateRangeSliders(); renderTendencia();
}

// ═══════════════════════════════════════════
// VIEW 1 — ANUAL
// ═══════════════════════════════════════════
function renderAnual(){ renderKPIs(); renderAnnualBar(); renderRanchBars(); renderLine(); renderStack(); renderAnnualTable(); }

function renderKPIs(){
  var yrs=activeYrList(), sym=state.currency==='usd'?'USD':'MXN';
  document.getElementById('kpiStrip').innerHTML=yrs.map(function(yr,i){
    var val=getAnnualVal(state.cat,yr);
    var prev=i>0?getAnnualVal(state.cat,yrs[i-1]):null;
    var delta=prev!==null?pct(val,prev):null;
    var cls=delta===null?'flat':parseFloat(delta)>0?'up':'down';
    var arrow=delta===null?'':parseFloat(delta)>0?'▲':'▼';
    var col=YEAR_COLORS[yr]||'#888';
    var note=yr===2026?' <small style="color:var(--dim);font-size:.58rem">(YTD)</small>':'';
    return '<div class="kpi" style="--accent:'+col+'">'+
      '<div class="kpi-yr">'+yr+note+' · '+sym+'</div>'+
      '<div class="kpi-val" style="color:'+col+'">'+fmt(val)+'</div>'+
      (delta!==null?'<div class="kpi-delta '+cls+'">'+arrow+' '+Math.abs(delta)+'% vs '+yrs[i-1]+'</div>':'<div class="kpi-delta flat">— base</div>')+
      '</div>';
  }).join('');
}

function renderAnnualBar(){
  destroyChart('chartBar');
  var yrs=activeYrList(), vals=yrs.map(function(y){return getAnnualVal(state.cat,y);});
  document.getElementById('barNote').textContent=(state.currency==='usd'?'USD':'MXN')+' · total anual';
  var trace={type:'bar',x:yrs.map(function(y){return y===2026?'2026*':String(y);}),y:vals,
    marker:{color:yrs.map(function(y){return YEAR_COLORS[y]||'#888';}),
    line:{color:yrs.map(function(y){return YEAR_COLORS[y]||'#888';}),width:1}},
    text:vals.map(fmt),textposition:'outside',hovertemplate:'<b>%{x}</b><br>%{text}<extra></extra>'};
  var layout=plotlyLayout({yaxis:{tickformat:'$,.0f'}});
  Plotly.newPlot('chartBar',[trace],layout,plotlyCfg());
}

function renderRanchBars(){
  var el=document.getElementById('ranchBars'), yr=state.ranchYear;
  var yrs=yr==='all'?activeYrList():[yr];
  var totals={};
  yrs.forEach(function(y){
    var d=(DATA.summary[state.cat]||{})[y]; if(!d) return;
    var src=state.currency==='usd'?d.ranches:d.ranches_mxn;
    Object.keys(src||{}).forEach(function(r){totals[r]=(totals[r]||0)+src[r];});
  });
  var total=Object.keys(totals).reduce(function(a,k){return a+totals[k];},0);
  if(!total){el.innerHTML='<div class="no-data">Sin datos de rancho</div>';return;}
  document.getElementById('ranchNote').textContent=(yr==='all'?'todos':yr)+' · '+(state.currency==='usd'?'USD':'MXN');
  var sorted=RANCH_ORDER.map(function(r){return [r,totals[r]||0];}).filter(function(a){return a[1]>0;}).sort(function(a,b){return b[1]-a[1];});
  var max=sorted[0][1];
  el.innerHTML=sorted.map(function(a){
    var r=a[0],v=a[1],col=RANCH_COLORS[r]||'#888';
    return '<div class="ranch-row">'+
      '<div class="ranch-lbl" style="color:'+col+'">'+r+'</div>'+
      '<div class="ranch-bar-outer"><div class="ranch-bar-inner" style="width:'+(v/max*100).toFixed(1)+'%;background:'+col+'"></div></div>'+
      '<div class="ranch-usd">'+fmt(v)+'</div>'+
      '<div class="ranch-pct">'+(v/total*100).toFixed(1)+'%</div></div>';
  }).join('');
}

function renderLine(){
  destroyChart('chartLine');
  var ws=DATA.weekly_series?DATA.weekly_series[state.cat]||{}:{};
  if(!DATA.weekly_series){
    DATA.weekly_detail.forEach(function(r){
      if(r.usd_total>0){var key=r.year+'-W'+String(r.week).padStart(2,'0');if(!ws[key]) ws[key]=0;ws[key]+=r.usd_total;}
    });
  }
  var yrs=activeYrList();
  var labels=[]; for(var i=1;i<=53;i++) labels.push(wFmt(i));
  var traces=yrs.map(function(yr){
    var data=new Array(53).fill(null);
    Object.keys(ws).forEach(function(key){
      var parts=key.split('-W');
      if(parseInt(parts[0])===yr){var w=parseInt(parts[1])-1;if(w>=0&&w<53) data[w]=ws[key];}
    });
    var col=YEAR_COLORS[yr]||'#888';
    return {type:'scatter',mode:'lines',name:String(yr),x:labels,y:data,
      line:{color:col,width:yr===2025?3:2,dash:yr===2026?'dash':'solid'},
      connectgaps:true,hovertemplate:'<b>'+yr+' %{x}</b><br>%{text}<extra></extra>',
      text:data.map(function(v){return v?fmt(v):'—';})};
  });
  if(!traces.length) return;
  var layout=plotlyLayout({yaxis:{tickformat:'$,.0f'},legend:{orientation:'h',y:1.12}});
  Plotly.newPlot('chartLine',traces,layout,plotlyCfg());
}

function renderStack(){
  destroyChart('chartStack');
  var yrs=activeYrList();
  document.getElementById('stackNote').textContent=state.currency==='usd'?'USD':'MXN';
  var rUsed={};
  yrs.forEach(function(y){var d=(DATA.summary[state.cat]||{})[y];if(d) Object.keys(d.ranches||{}).forEach(function(r){rUsed[r]=1;});});
  var rList=RANCH_ORDER.filter(function(r){return rUsed[r];});
  if(!rList.length) return;
  var xLabels=yrs.map(function(y){return y===2026?'2026*':String(y);});
  var traces=rList.map(function(r){
    var vals=yrs.map(function(y){
      var d=(DATA.summary[state.cat]||{})[y];if(!d) return 0;
      return state.currency==='mxn'&&d.usd>0?(d.ranches[r]||0)*d.mxn/d.usd:(d.ranches[r]||0);
    });
    return {type:'bar',name:r,x:xLabels,y:vals,
      marker:{color:RANCH_COLORS[r]||'#888'},
      hovertemplate:'<b>'+r+'</b><br>%{text}<extra></extra>',
      text:vals.map(fmt)};
  });
  var layout=plotlyLayout({barmode:'stack',yaxis:{tickformat:'$,.0f'},legend:{orientation:'v',x:1.02}});
  Plotly.newPlot('chartStack',traces,layout,plotlyCfg());
}

function renderAnnualTable(){
  var yrs=activeYrList(), sym=state.currency==='usd'?'USD':'MXN';
  document.getElementById('tableNote').textContent=sym+' · variación vs año anterior';
  document.getElementById('tableBody').innerHTML=yrs.map(function(yr,i){
    var d=(DATA.summary[state.cat]||{})[yr]||{usd:0,mxn:0,ranches:{},ranches_mxn:{}};
    var val=state.currency==='usd'?d.usd:d.mxn;
    var prev=i>0?getAnnualVal(state.cat,yrs[i-1]):null;
    var delta=prev!==null?pct(val,prev):null;
    var dStr=delta!==null?'<span class="'+(parseFloat(delta)>0?'chg-pos':'chg-neg')+'">'+(parseFloat(delta)>0?'+':'')+delta+'%</span>':'<span class="chg-0">—</span>';
    var col=YEAR_COLORS[yr]||'#888';
    var ranchSrc=state.currency==='usd'?d.ranches:d.ranches_mxn;
    var cells=KEY_RANCHES.map(function(r){var v=ranchSrc[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'#3a5a48')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
    return '<tr><td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+(yr===2026?'*':'')+'</strong></td>'+
      '<td style="color:'+col+';font-weight:600">'+fmt(val)+'</td><td>'+dStr+'</td>'+cells+'</tr>';
  }).join('');
}

// ═══════════════════════════════════════════
// VIEW 2 — POR SEMANA
// ═══════════════════════════════════════════
function renderSemana(){
  if(!allWeeks.length) return;
  var weekNum=allWeeks[state.weekIdx], yrs=activeYrList();
  var sym=state.currency==='usd'?'USD':'MXN';
  document.getElementById('swTableNote').textContent=sym;

  if(isCombined(state.cat)){
    // ── MODO MIRFE + MIPE: 2 filas por año ──
    var rows=[];
    yrs.forEach(function(yr){
      var col=YEAR_COLORS[yr]||'#888';
      // Fila MIRFE
      var rm=getDetail(CAT_MIRFE,weekNum,yr);
      var dm=rm.length?aggregateDetail(rm):null;
      var valm=dm?(state.currency==='usd'?dm.usd:dm.mxn):0;
      var ranchSrcM=dm?(state.currency==='usd'?dm.ranches:dm.ranches_mxn):{};
      var cellsM=KEY_RANCHES.map(function(r){var v=ranchSrcM[r]||0;if(v>0){return '<td class="prod-cell" data-r="'+r+'" data-t="MIRFE" data-w="'+weekNum+'" data-y="'+yr+'" style="color:'+(RANCH_COLORS[r]||'#888')+'cc">'+fmt(v)+'</td>';}return '<td style="color:#94a3b8">—</td>';}).join('');
      rows.push('<tr>'+
        '<td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong>'+
        '<span style="font-size:.6rem;margin-left:4px;color:#b45309;font-family:IBM Plex Mono,monospace">MIRFE</span></td>'+
        '<td style="color:'+col+'">'+wFmt(weekNum)+'</td>'+
        '<td style="color:var(--dim);font-size:.65rem">'+(dm&&dm.date_range?dm.date_range:'—')+'</td>'+
        '<td style="color:'+(valm>0?col:'var(--dim)')+';font-weight:600">'+fmt(valm)+'</td>'+
        '<td>—</td>'+cellsM+'</tr>');
      // Fila MIPE
      var rp=getDetail(CAT_MIPE,weekNum,yr);
      var dp=rp.length?aggregateDetail(rp):null;
      var valp=dp?(state.currency==='usd'?dp.usd:dp.mxn):0;
      var ranchSrcP=dp?(state.currency==='usd'?dp.ranches:dp.ranches_mxn):{};
      var cellsP=KEY_RANCHES.map(function(r){var v=ranchSrcP[r]||0;if(v>0){return '<td class="prod-cell" data-r="'+r+'" data-t="MIPE" data-w="'+weekNum+'" data-y="'+yr+'" style="color:'+(RANCH_COLORS[r]||'#888')+'cc">'+fmt(v)+'</td>';}return '<td style="color:#94a3b8">—</td>';}).join('');
      rows.push('<tr style="border-bottom:2px solid var(--border)">'+
        '<td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong>'+
        '<span style="font-size:.6rem;margin-left:4px;color:#2563eb;font-family:IBM Plex Mono,monospace">MIPE</span></td>'+
        '<td style="color:'+col+'">'+wFmt(weekNum)+'</td>'+
        '<td style="color:var(--dim);font-size:.65rem">'+(dp&&dp.date_range?dp.date_range:'—')+'</td>'+
        '<td style="color:'+(valp>0?col:'var(--dim)')+';font-weight:600">'+fmt(valp)+'</td>'+
        '<td>—</td>'+cellsP+'</tr>');
    });
    document.getElementById('swTableBody').innerHTML=rows.join('');
  } else {
    // ── MODO NORMAL ──
    var byYear=getWeekByYear(state.cat,weekNum);
    document.getElementById('swTableBody').innerHTML=yrs.map(function(yr,i){
      var d=byYear[yr], val=d?(state.currency==='usd'?d.usd:d.mxn):0;
      var prevD=i>0?byYear[yrs[i-1]]:null, prevVal=prevD?(state.currency==='usd'?prevD.usd:prevD.mxn):null;
      var delta=prevVal!==null?pct(val,prevVal):null;
      var col=YEAR_COLORS[yr]||'#888';
      var dStr=delta!==null?'<span class="'+(parseFloat(delta)>0?'chg-pos':'chg-neg')+'">'+(parseFloat(delta)>0?'+':'')+delta+'%</span>':'<span class="chg-0">—</span>';
      var ranchSrc=d?(state.currency==='usd'?d.ranches:d.ranches_mxn):{};
      var isMant=(state.cat==='MANTENIMIENTO');
      var isMatEmp=(state.cat==='MATERIAL DE EMPAQUE');
      var cells=KEY_RANCHES.map(function(r){var v=ranchSrc[r]||0;if(v>0&&isMant){return '<td class="prod-cell" data-r="'+r+'" data-t="ALL" data-w="'+weekNum+'" data-y="'+yr+'" data-src="mp" style="color:'+(RANCH_COLORS[r]||'#888')+'cc">'+fmt(v)+'</td>';}if(v>0&&isMatEmp){return '<td class="prod-cell" data-r="'+r+'" data-t="ALL" data-w="'+weekNum+'" data-y="'+yr+'" data-src="me" style="color:'+(RANCH_COLORS[r]||'#888')+'cc">'+fmt(v)+'</td>';}return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'#3a5a48')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
      return '<tr>'+
        '<td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong></td>'+
        '<td style="color:'+col+'">'+wFmt(weekNum)+'</td>'+
        '<td style="color:var(--dim);font-size:.65rem">'+(d&&d.date_range?d.date_range:'—')+'</td>'+
        '<td style="color:'+col+';font-weight:600">'+fmt(val)+'</td>'+
        '<td>'+dStr+'</td>'+cells+'</tr>';
    }).join('');
  }
}

// ═══════════════════════════════════════════
// VIEW 3 — TENDENCIA / RANGO
// ═══════════════════════════════════════════
function renderTendencia(){
  var f=state.fromWeek, t=state.toWeek, yrs=activeYrList(), byYear=getRangeByYear(state.cat,f,t);
  var sym=state.currency==='usd'?'USD':'MXN';
  document.getElementById('rangeStats').innerHTML=yrs.map(function(yr){
    var d=byYear[yr]; if(!d) return '';
    var val=state.currency==='usd'?d.usd:d.mxn;
    var wks=Object.keys(d.weekly||{}).length, avg=wks>0?(val/wks):0;
    var col=YEAR_COLORS[yr]||'#888';
    var subLine='';
    if(isCombined(state.cat)){
      var mv=state.currency==='usd'?(d.mirfe_usd||0):(d.mirfe_mxn||0);
      var pv=state.currency==='usd'?(d.mipe_usd||0) :(d.mipe_mxn||0);
      subLine='<div style="margin-top:4px;font-size:.6rem;font-family:IBM Plex Mono,monospace">'+
        '<span style="color:#f0b429">⬤ MIRFE '+fmt(mv)+'</span> '+
        '<span style="color:#3b9eff">⬤ MIPE '+fmt(pv)+'</span></div>';
    }
    return '<div class="stat-box" style="border-color:'+col+'33">'+
      '<div class="stat-label">'+yr+' · '+sym+'</div>'+
      '<div class="stat-val" style="color:'+col+'">'+fmt(val)+'</div>'+
      '<div style="font-size:.62rem;color:var(--dim);font-family:IBM Plex Mono,monospace">'+fmt(avg)+'/sem · '+wks+' semanas</div>'+
      subLine+'</div>';
  }).join('');

  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;}), rLabels=rangeWeeks.map(wFmt);

  destroyChart('chartRangeLine');
  var rTraces=yrs.map(function(yr){
    var d=byYear[yr], data=rangeWeeks.map(function(w){return d&&d.weekly[w]?d.weekly[w]:null;}), col=YEAR_COLORS[yr]||'#888';
    return {type:'scatter',mode:'lines'+(rangeWeeks.length<20?'+markers':''),name:String(yr),x:rLabels,y:data,
      line:{color:col,width:yr===2025?3:2,dash:yr===2026?'dash':'solid'},
      connectgaps:true,hovertemplate:'<b>'+yr+' %{x}</b><br>%{text}<extra></extra>',
      text:data.map(function(v){return v?fmt(v):'—';})};
  });
  if(rTraces.some(function(d){return d.y.some(function(v){return v!==null;});})){
    Plotly.newPlot('chartRangeLine',rTraces,plotlyLayout({yaxis:{tickformat:'$,.0f'},legend:{orientation:'h',y:1.12}}),plotlyCfg());
  }

  destroyChart('chartRangeBar');
  document.getElementById('rangeBarNote').textContent=sym;
  var rbVals=yrs.map(function(y){var d=byYear[y];return d?(state.currency==='usd'?d.usd:d.mxn):0;});
  Plotly.newPlot('chartRangeBar',[{type:'bar',x:yrs.map(String),y:rbVals,
    marker:{color:yrs.map(function(y){return YEAR_COLORS[y]||'#888';})},
    text:rbVals.map(fmt),textposition:'outside',
    hovertemplate:'<b>%{x}</b><br>%{text}<extra></extra>'}],
    plotlyLayout({yaxis:{tickformat:'$,.0f'}}),plotlyCfg());

  destroyChart('chartCumul');
  var cTraces=yrs.map(function(yr){
    var d=byYear[yr], cum=0, data=rangeWeeks.map(function(w){if(d&&d.weekly[w]) cum+=d.weekly[w];return cum||null;}), col=YEAR_COLORS[yr]||'#888';
    return {type:'scatter',mode:'lines',name:String(yr),x:rLabels,y:data,
      fill:'tozeroy',fillcolor:col+'25',line:{color:col,width:2,dash:yr===2026?'dash':'solid'},
      connectgaps:true,hovertemplate:'<b>'+yr+' %{x}</b><br>%{text}<extra></extra>',
      text:data.map(function(v){return v?fmt(v):'—';})};
  });
  Plotly.newPlot('chartCumul',cTraces,plotlyLayout({yaxis:{tickformat:'$,.0f'},legend:{orientation:'h',y:1.12}}),plotlyCfg());

  renderRangeTable(f,t,yrs,byYear);
}

function renderHeatmap(f,t,yrs){
  var el=document.getElementById('heatmapWrap');
  var wi=allWeeks.filter(function(w){return w>=f&&w<=t;});
  if(!wi.length){el.innerHTML='<div class="no-data">Sin semanas en el rango</div>';return;}
  var matrix={}, globalMax=0;
  yrs.forEach(function(yr){
    matrix[yr]={};
    wi.forEach(function(w){
      var recs=getDetail(state.cat,w,yr), val=recs.reduce(function(a,r){return a+r.usd_total;},0);
      matrix[yr][w]=val; if(val>globalMax) globalMax=val;
    });
  });
  var html='<table class="heatmap"><thead><tr><th class="yr-th">Año</th>';
  wi.forEach(function(w){html+='<th>'+wFmt(w)+'</th>';});
  html+='</tr></thead><tbody>';
  yrs.forEach(function(yr){
    var col=YEAR_COLORS[yr]||'#888';
    html+='<tr><td class="hm-yr-lbl" style="color:'+col+'">'+yr+'</td>';
    wi.forEach(function(w){
      var v=matrix[yr][w], ratio=globalMax>0?v/globalMax:0;
      var bg=heatColor(ratio), tc=ratio>0.5?'#fff':(ratio>0?'#0a7c52':'#94a3b8'), title=v>0?fmt(v):'—';
      html+='<td style="background:'+bg+';color:'+tc+'" title="'+yr+' W'+String(w).padStart(2,'0')+': '+title+'" onclick="jumpToWeek('+w+')">'+(v>0?fmt(v):'·')+'</td>';
    });
    html+='</tr>';
  });
  html+='</tbody></table>';
  el.innerHTML=html;
}

function jumpToWeek(w){var idx=allWeeks.indexOf(w);if(idx>=0){state.weekIdx=idx;setView('semana');updateWeekSlider();}}

var rangeTableGroup='year';
function setRangeTableGroup(g){
  rangeTableGroup=g;
  document.getElementById('rtgYear').classList.toggle('active',g==='year');
  document.getElementById('rtgWeek').classList.toggle('active',g==='week');
  renderRangeTable(state.fromWeek,state.toWeek,activeYrList(),getRangeByYear(state.cat,state.fromWeek,state.toWeek));
}
function deltaCell(val,prev){
  if(prev===null||prev===undefined||prev===0) return '<td class="delta-cell chg-0">—</td>';
  var diff=val-prev, p=((diff/prev)*100).toFixed(1), cls=diff>0?'chg-pos':diff<0?'chg-neg':'chg-0', sign=diff>0?'+':'';
  return '<td class="delta-cell '+cls+'"><span class="delta-amt">'+sign+fmt(diff)+'</span><span class="delta-pct">'+sign+p+'%</span></td>';
}
function renderRangeTable(f,t,yrs,byYear){
  var sym=state.currency==='usd'?'USD':'MXN';
  var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;});
  var ranchCols=['Prop-RM','PosCo-RM','Campo-RM','Isabela','Cecilia','Cecilia 25','Christina'];
  document.getElementById('rangeTableNote').textContent=sym+' · '+wFmt(f)+' → '+wFmt(t);
  document.getElementById('rangeTableSub').textContent=rangeWeeks.length+' semanas × '+yrs.length+' años = '+rangeWeeks.length*yrs.length+' filas';

  // Para MIRFE/MIPE: cargar datos de ambas categorías por separado
  var weekData={};
  if(isCombined(state.cat)){
    yrs.forEach(function(yr){
      weekData[yr]={};
      rangeWeeks.forEach(function(w){
        var rm=getDetail(CAT_MIRFE,w,yr), rp=getDetail(CAT_MIPE,w,yr);
        weekData[yr][w]={
          mirfe: rm.length?aggregateDetail(rm):null,
          mipe:  rp.length?aggregateDetail(rp):null
        };
      });
    });
  } else {
    yrs.forEach(function(yr){weekData[yr]={};rangeWeeks.forEach(function(w){var recs=getDetail(state.cat,w,yr);if(recs.length) weekData[yr][w]=aggregateDetail(recs);});});
  }
  var head, body;
  if(rangeTableGroup==='year'){
    head='<tr><th>Semana</th><th>Fecha</th><th>Total '+sym+'</th><th>Δ$ vs sem ant.</th>'+ranchCols.map(function(r){return '<th>'+r+'</th>';}).join('')+'</tr>';
    if(isCombined(state.cat)){
      // ── MIRFE+MIPE: por año → 2 sub-filas por semana ──
      body=yrs.map(function(yr,yi){
        var col=YEAR_COLORS[yr]||'#888';
        var hdr='<tr class="tr-group-hdr" style="--accent:'+col+'"><td colspan="'+(4+ranchCols.length)+'" style="color:'+col+'">📅 '+yr+'</td></tr>';
        var prevM=null, prevP=null;
        var wkRows=rangeWeeks.map(function(w){
          var dd=weekData[yr][w]||{}, dm=dd.mirfe, dp=dd.mipe;
          var vm=dm?(state.currency==='usd'?dm.usd:dm.mxn):0;
          var vp=dp?(state.currency==='usd'?dp.usd:dp.mxn):0;
          var dr=dm&&dm.date_range?dm.date_range:(dp&&dp.date_range?dp.date_range:'—');
          var dCellM=deltaCell(vm,prevM); prevM=vm>0?vm:prevM;
          var dCellP=deltaCell(vp,prevP); prevP=vp>0?vp:prevP;
          var rcM=ranchCols.map(function(r){if(!dm) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?dm.ranches:dm.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          var rcP=ranchCols.map(function(r){if(!dp) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?dp.ranches:dp.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          return '<tr class="tr-week">'+
            '<td style="color:'+col+';font-weight:600">'+wFmt(w)+' <span style="color:#f0b429;font-size:.6rem">MIRFE</span></td>'+
            '<td style="color:var(--dim);font-size:.63rem">'+dr+'</td>'+
            '<td style="color:'+(vm>0?col:'var(--dim)')+';font-weight:'+(vm>0?'600':'400')+'">'+fmt(vm)+'</td>'+
            dCellM+rcM+'</tr>'+
            '<tr class="tr-week" style="border-bottom:1px solid var(--border)">'+
            '<td style="color:'+col+';font-weight:600">'+wFmt(w)+' <span style="color:#3b9eff;font-size:.6rem">MIPE</span></td>'+
            '<td style="color:var(--dim);font-size:.63rem">'+dr+'</td>'+
            '<td style="color:'+(vp>0?col:'var(--dim)')+';font-weight:'+(vp>0?'600':'400')+'">'+fmt(vp)+'</td>'+
            dCellP+rcP+'</tr>';
        }).join('');
        return hdr+wkRows;
      }).join('');
    } else {
      body=yrs.map(function(yr,yi){
        var col=YEAR_COLORS[yr]||'#888', yearTotal=byYear[yr]?(state.currency==='usd'?byYear[yr].usd:byYear[yr].mxn):0;
        var prevYr=yi>0?byYear[yrs[yi-1]]:null, prevYrVal=prevYr?(state.currency==='usd'?prevYr.usd:prevYr.mxn):null;
        var yDiff=prevYrVal!==null?yearTotal-prevYrVal:null, yPct=prevYrVal!==null&&prevYrVal!==0?((yearTotal-prevYrVal)/prevYrVal*100).toFixed(1):null;
        var yCls=yDiff===null?'chg-0':yDiff>0?'chg-pos':'chg-neg', ySign=yDiff!==null&&yDiff>0?'+':'';
        var hdr='<tr class="tr-group-hdr" style="--accent:'+col+'"><td colspan="2" style="color:'+col+'">📅 '+yr+' — Total del rango</td>'+
          '<td style="color:'+col+'">'+fmt(yearTotal)+'</td>'+
          '<td class="delta-cell '+yCls+'">'+(yDiff!==null?'<span class="delta-amt">'+ySign+fmt(yDiff)+'</span><span class="delta-pct">'+ySign+(yPct||'0')+'%  vs '+yrs[yi-1]+'</span>':'<span class="delta-amt chg-0">— base</span>')+'</td>'+
          ranchCols.map(function(r){var d=byYear[yr];if(!d) return '<td>—</td>';var src=state.currency==='usd'?d.ranches:d.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+';font-size:.68rem">'+(v>0?fmt(v):'—')+'</td>';}).join('')+'</tr>';
        var prevWkVal=null;
        var wkRows=rangeWeeks.map(function(w){
          var d=weekData[yr][w], val=d?(state.currency==='usd'?d.usd:d.mxn):0, dCell=deltaCell(val,prevWkVal);
          prevWkVal=val>0?val:prevWkVal;
          var ranchCells=ranchCols.map(function(r){if(!d) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?d.ranches:d.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          return '<tr class="tr-week"><td style="color:'+col+';font-weight:600">'+wFmt(w)+'</td>'+
            '<td style="color:var(--dim);font-size:.63rem">'+(d&&d.date_range?d.date_range:'—')+'</td>'+
            '<td style="color:'+(val>0?col:'var(--dim)')+';font-weight:'+(val>0?'600':'400')+'">'+fmt(val)+'</td>'+
            dCell+ranchCells+'</tr>';
        }).join('');
        return hdr+wkRows;
      }).join('');
    }
  } else {
    head='<tr><th>Año</th><th>Total '+sym+'</th><th>Δ$ vs año ant.</th>'+ranchCols.map(function(r){return '<th>'+r+'</th>';}).join('')+'</tr>';
    if(isCombined(state.cat)){
      // ── MIRFE+MIPE: por semana → 2 sub-filas por año ──
      body=rangeWeeks.map(function(w){
        var dateEx='';
        yrs.forEach(function(yr){var dd=weekData[yr][w]||{};if(dd.mirfe&&dd.mirfe.date_range) dateEx=dd.mirfe.date_range;else if(dd.mipe&&dd.mipe.date_range) dateEx=dd.mipe.date_range;});
        var hdr='<tr class="tr-group-hdr" style="--accent:var(--green)"><td colspan="2" style="color:var(--green)">📆 '+wFmt(w)+(dateEx?' <span style="font-size:.6rem;color:var(--dim)">'+dateEx+'</span>':'')+'</td><td colspan="'+(2+ranchCols.length)+'"></td></tr>';
        var prevM=null, prevP=null;
        var yrRows=yrs.map(function(yr){
          var col=YEAR_COLORS[yr]||'#888', dd=weekData[yr][w]||{}, dm=dd.mirfe, dp=dd.mipe;
          var vm=dm?(state.currency==='usd'?dm.usd:dm.mxn):0;
          var vp=dp?(state.currency==='usd'?dp.usd:dp.mxn):0;
          var dCellM=deltaCell(vm,prevM); prevM=vm>0?vm:prevM;
          var dCellP=deltaCell(vp,prevP); prevP=vp>0?vp:prevP;
          var rcM=ranchCols.map(function(r){if(!dm) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?dm.ranches:dm.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          var rcP=ranchCols.map(function(r){if(!dp) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?dp.ranches:dp.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          return '<tr class="tr-week"><td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong> <span style="color:#f0b429;font-size:.6rem">MIRFE</span></td>'+
            '<td style="color:'+(vm>0?col:'var(--dim)')+';font-weight:'+(vm>0?'600':'400')+'">'+fmt(vm)+'</td>'+dCellM+rcM+'</tr>'+
            '<tr class="tr-week" style="border-bottom:1px solid var(--border)"><td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong> <span style="color:#3b9eff;font-size:.6rem">MIPE</span></td>'+
            '<td style="color:'+(vp>0?col:'var(--dim)')+';font-weight:'+(vp>0?'600':'400')+'">'+fmt(vp)+'</td>'+dCellP+rcP+'</tr>';
        }).join('');
        return hdr+yrRows;
      }).join('');
    } else {
      body=rangeWeeks.map(function(w){
        var dateEx='';
        yrs.forEach(function(yr){if(weekData[yr][w]&&weekData[yr][w].date_range) dateEx=weekData[yr][w].date_range;});
        var hdr='<tr class="tr-group-hdr" style="--accent:var(--green)"><td colspan="2" style="color:var(--green)">📆 '+wFmt(w)+(dateEx?' <span style="font-size:.6rem;color:var(--dim)">'+dateEx+'</span>':'')+'</td><td colspan="'+(2+ranchCols.length)+'"></td></tr>';
        var prevYrVal=null;
        var yrRows=yrs.map(function(yr){
          var col=YEAR_COLORS[yr]||'#888', d=weekData[yr][w], val=d?(state.currency==='usd'?d.usd:d.mxn):0, dCell=deltaCell(val,prevYrVal);
          prevYrVal=val>0?val:prevYrVal;
          var ranchCells=ranchCols.map(function(r){if(!d) return '<td style="color:var(--dim)">—</td>';var src=state.currency==='usd'?d.ranches:d.ranches_mxn,v=src[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'var(--dim)')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
          return '<tr class="tr-week"><td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong></td><td style="color:'+(val>0?col:'var(--dim)')+';font-weight:'+(val>0?'600':'400')+'">'+fmt(val)+'</td>'+dCell+ranchCells+'</tr>';
        }).join('');
        var wkTotal=yrs.reduce(function(acc,yr){var d=weekData[yr][w];return acc+(d?(state.currency==='usd'?d.usd:d.mxn):0);},0);
        var totalRow='<tr class="tr-total"><td style="color:var(--green)">TOTAL</td><td style="color:var(--green)">'+fmt(wkTotal)+'</td><td colspan="'+(1+ranchCols.length)+'"></td></tr>';
        return hdr+yrRows+totalRow;
      }).join('');
    }
  }
  document.getElementById('rangeDetailHead').innerHTML=head;
  document.getElementById('rangeTableBody').innerHTML=body;
  setTimeout(initScrollHints,80);
}

// ═══════════════════════════════════════════
// SCROLL HINTS
// ═══════════════════════════════════════════
function initScrollHints(){
  [{wrap:'wrapAnual',hint:'hintAnual'},{wrap:'wrapSemana',hint:'hintSemana'},{wrap:'wrapRange',hint:'hintRange'}].forEach(function(p){
    var el=document.getElementById(p.wrap);if(!el) return;
    function check(){
      var has=el.scrollWidth>el.clientWidth+4;
      el.classList.toggle('no-overflow',!has);
      if(p.hint){var h=document.getElementById(p.hint);if(h) h.classList.toggle('show',has);}
    }
    check();
    el.addEventListener('scroll',function(){if(p.hint){var h=document.getElementById(p.hint);if(h) h.classList.remove('show');}el.classList.add('no-overflow');},{once:true});
    window.addEventListener('resize',check);
  });
}

// ═══════════════════════════════════════════
// MODAL DE PRODUCTOS
// ═══════════════════════════════════════════
function showProductos(rancho, tipo, weekNum, yr, src) {
  src = src || 'pr';
  lastProductos = {rancho: rancho, tipo: tipo, src: src};
  var semCode = (yr % 100) * 100 + weekNum;
  var semCodeStr = String(semCode);

  var allProds = src === 'mp' ? (DATA.productos_mp || {})
               : src === 'me' ? (DATA.productos_me || {})
               : (DATA.productos || {});
  var debug    = src === 'mp' ? (DATA.productos_mp_debug || {})
               : src === 'me' ? (DATA.productos_me_debug || {})
               : (DATA.productos_debug || {});
  var prefix   = src === 'mp' ? 'MP' : src === 'me' ? 'ME' : 'PR';

  var prods = allProds[semCode] || allProds[semCodeStr] || null;

  var list = [];
  if(prods){
    var byRanch = prods[rancho];
    if(byRanch){
      if(tipo === 'ALL'){
        (byRanch['MIRFE'] || []).forEach(function(p){ list.push(p); });
        (byRanch['MIPE']  || []).forEach(function(p){ list.push(p); });
      } else {
        list = byRanch[tipo] || [];
      }
    }
  }

  var tipoNombre = src === 'mp' ? 'Mantenimiento'
    : src === 'me' ? 'Material de Empaque'
    : tipo === 'MIRFE' ? 'Riego / Fertilización' : 'Material de Protección';

  document.getElementById('productosTitle').innerHTML =
    rancho + ' <span>— ' + tipoNombre + '</span>';

  var wkLabel = 'W' + String(weekNum).padStart(2,'0') + ' · ' + yr;

  var footer = document.getElementById('pnlFooter');

  if(!list.length){
    var hojasKey = 'hojas_' + prefix.toLowerCase() + '_encontradas';
    var hojasEncontradas = debug[hojasKey] || [];
    var ranchosDisp = prods ? Object.keys(prods).join(', ') : 'ninguno';
    var msg = hojasEncontradas.length === 0 ? 'No se encontraron hojas ' + prefix + ' en el sheet'
      : !prods ? 'Sin datos · Semana ' + semCode + ' · Hojas: ' + hojasEncontradas.join(', ')
      : 'Rancho no encontrado · Disponibles: ' + ranchosDisp;

    document.getElementById('pnlKpis').innerHTML = '';
    document.getElementById('productosContent').innerHTML =
      '<tr><td colspan="4" style="border:none">' +
      '<div class="pnl-empty">' +
      '<svg class="pnl-empty-ico" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg>' +
      'Sin productos disponibles para este rancho.<br>' +
      '<span style="font-size:.6rem;opacity:.7">' + msg + '</span>' +
      '</div></td></tr>';
    footer.style.display = 'none';
  } else {
    var totalGasto = 0, totalUnidades = 0;

    // Agrupar por ubicacion
    var grupos = {}, gruposOrder = [];
    list.forEach(function(p){
      var det = p[3] ? p[3] : '—';
      if(!grupos[det]){ grupos[det]=[]; gruposOrder.push(det); }
      grupos[det].push(p);
    });

    // KPI strip
    var totalItems = list.length;
    list.forEach(function(p){
      totalUnidades += p[1] ? parseFloat(p[1]) : 0;
      totalGasto    += p[2] ? parseFloat(p[2]) : 0;
    });

    // Sin KPI strip separado — las métricas van en el header
    document.getElementById('pnlKpis').innerHTML = '';
    // Build rows
    var rows = '';
    gruposOrder.forEach(function(det){
      var grpGasto = 0, grpUnidades = 0;
      grupos[det].forEach(function(p){
        grpGasto    += p[2] ? parseFloat(p[2]) : 0;
        grpUnidades += p[1] ? parseFloat(p[1]) : 0;
      });

      if(gruposOrder.length > 1){
        var grpAmtStr = (grpGasto < 0 ? '-$' : '$') + Math.abs(grpGasto).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
        var grpUStr   = Math.abs(grpUnidades).toLocaleString('en-US',{maximumFractionDigits:0});
        rows += '<tr class="pnl-grp">' +
          '<td colspan="2">' + det + '<span class="pnl-grp-sub">' + grupos[det].length + ' productos</span></td>' +
          '<td class="qty" style="font-size:.6rem;color:var(--muted);font-weight:500">' + grpUStr + '</td>' +
          '<td class="pnl-grp-amt">' + grpAmtStr + '</td>' +
          '</tr>';
      }

      grupos[det].forEach(function(p){
        var gasto    = p[2] ? parseFloat(p[2]) : 0;
        var unidades = p[1] ? parseFloat(p[1]) : 0;
        var amtClass = 'amt ' + (gasto < 0 ? 'neg' : 'pos');
        var amtStr   = gasto !== 0
          ? (gasto < 0 ? '-$' : '$') + Math.abs(gasto).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})
          : '—';
        var qtyStr = unidades !== 0
          ? Math.abs(unidades).toLocaleString('en-US',{maximumFractionDigits:2})
          : '—';
        var ubicCell = gruposOrder.length > 1 ? '' : (det !== '—' ? det : '');
        rows += '<tr>' +
          '<td>' + ubicCell + '</td>' +
          '<td>' + p[0] + '</td>' +
          '<td class="qty">' + qtyStr + '</td>' +
          '<td class="' + amtClass + '">' + amtStr + '</td>' +
          '</tr>';
      });
    });

    // Fila TOTAL dentro de la tabla (Power BI style)
    var fSign = totalGasto < 0 ? '-$' : '$';
    var totalAmtStr = fSign + Math.abs(totalGasto).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
    var totalUStr   = Math.abs(totalUnidades).toLocaleString('en-US',{maximumFractionDigits:0});
    rows += '<tr class="pnl-total">' +
      '<td colspan="2">TOTAL GENERAL</td>' +
      '<td class="qty">' + totalUStr + '</td>' +
      '<td class="amt ' + (totalGasto<0?'neg':'pos') + '">' + totalAmtStr + '</td>' +
      '</tr>';

    document.getElementById('productosContent').innerHTML = rows;
    footer.style.display = 'none';
  }

  var section = document.getElementById('productosSection');
  section.classList.add('show');
  section.style.display = 'block';
  setTimeout(function(){ section.scrollIntoView({behavior:'smooth',block:'nearest'}); }, 100);
}

function closeProductos() {
  lastProductos = null;
  var section = document.getElementById('productosSection');
  section.classList.remove('show');
  setTimeout(function() { section.style.display = 'none'; }, 300);
}


// ═══════════════════════════════════════════
// COSTO DE SERVICIOS
// ═══════════════════════════════════════════
var SUBCAT_ICONS = {
  'Electricidad':'⚡','Fletes y Acarreos':'🚛','Gastos de Exportación':'📦',
  'Certificado Fitosanitario':'📋','Transporte de Personal':'🚌',
  'Compra de Flor a Terceros':'🌸','Comida para el Personal':'🍽','RO, TEL, RTA.Alim':'📡'
};
var SUBCAT_COLORS = {
  'Electricidad':'#f59e0b','Fletes y Acarreos':'#3b82f6','Gastos de Exportación':'#8b5cf6',
  'Certificado Fitosanitario':'#10b981','Transporte de Personal':'#f97316',
  'Compra de Flor a Terceros':'#ec4899','Comida para el Personal':'#14b8a6','RO, TEL, RTA.Alim':'#6366f1'
};
var SUBCATS_ORDER = ['Electricidad','Fletes y Acarreos','Gastos de Exportación','Certificado Fitosanitario','Transporte de Personal','Compra de Flor a Terceros','Comida para el Personal','RO, TEL, RTA.Alim'];

function buildServSelect(){
  var el=document.getElementById('servSelect');
  if(!el) return;
  var found=[];
  (DATA.servicios_data||[]).forEach(function(r){if(found.indexOf(r.subcat)<0) found.push(r.subcat);});
  el.innerHTML='<option value="" disabled selected>COSTO DE SERVICIOS</option>'+
    '<option value="">— TODAS —</option>'+
    SUBCATS_ORDER.filter(function(s){return found.indexOf(s)>=0;}).map(function(s){
      return '<option value="'+s+'">'+s+'</option>';
    }).join('');
}

function selectServ(val){
  state.servFilter = val||null;
  // Ocultar pestañas Por Año y Tendencia — solo dejar Por Semana
  document.getElementById('vtAnual').style.display='none';
  document.getElementById('vtTendencia').style.display='none';
  document.getElementById('vtSemana').style.display='';
  // Asegurar vista en Por Semana
  if(state.view!=='semana'){
    state.view='semana';
    document.getElementById('viewAnual').style.display='none';
    document.getElementById('viewTendencia').style.display='none';
    document.getElementById('viewSemana').style.display='';
    document.getElementById('vtSemana').classList.add('active');
    document.getElementById('vtAnual').classList.remove('active');
    document.getElementById('vtTendencia').classList.remove('active');
  }
  renderServPanel(state.servFilter);
}

function renderServPanel(subcatFilter){
  var weekNum=allWeeks[state.weekIdx];
  var yrs=activeYrList(), yr=yrs[yrs.length-1];
  var sym=state.currency==='usd'?'USD':'MXN';

  var data=(DATA.servicios_data||[]).filter(function(r){
    return r.week===weekNum&&r.year===yr&&(!subcatFilter||r.subcat===subcatFilter);
  });
  if(!data.length){
    data=(DATA.servicios_data||[]).filter(function(r){
      return r.year===yr&&(!subcatFilter||r.subcat===subcatFilter);
    });
  }

  var ranchData={};
  data.forEach(function(r){
    var src=state.currency==='usd'?r.usd_ranches:r.mxn_ranches;
    Object.keys(src).forEach(function(ranch){
      if(!src[ranch]) return;
      if(!ranchData[ranch]) ranchData[ranch]={};
      ranchData[ranch][r.subcat]=(ranchData[ranch][r.subcat]||0)+src[ranch];
    });
  });

  var ranchTotals={};
  Object.keys(ranchData).forEach(function(r){
    ranchTotals[r]=Object.keys(ranchData[r]).reduce(function(a,k){return a+ranchData[r][k];},0);
  });
  var grandTotal=Object.keys(ranchTotals).reduce(function(a,k){return a+ranchTotals[k];},0);
  var topRanch=Object.keys(ranchTotals).sort(function(a,b){return ranchTotals[b]-ranchTotals[a];})[0]||'—';
  var subcatTotals={};
  Object.keys(ranchData).forEach(function(r){Object.keys(ranchData[r]).forEach(function(sc){subcatTotals[sc]=(subcatTotals[sc]||0)+ranchData[r][sc];});});
  var topSubcat=Object.keys(subcatTotals).sort(function(a,b){return subcatTotals[b]-subcatTotals[a];})[0]||'—';

  var title=(subcatFilter?'SERVICIOS — '+subcatFilter:'COSTO DE SERVICIOS')+' · W'+String(weekNum).padStart(2,'0')+' '+yr;
  document.getElementById('servPanelTitle').textContent=title;

  document.getElementById('servKpiStrip').innerHTML=
    '<div class="serv-kpi"><div class="serv-kpi-label">Total de Gastos</div><div class="serv-kpi-val">'+fmt(grandTotal)+'</div><div class="serv-kpi-sub">'+sym+'</div></div>'+
    '<div class="serv-kpi"><div class="serv-kpi-label">Rancho con Más Gastos</div><div class="serv-kpi-val" style="color:'+(RANCH_COLORS[topRanch]||'#0a7c52')+'">'+topRanch+'</div><div class="serv-kpi-sub">'+fmt(ranchTotals[topRanch]||0)+'</div></div>'+
    '<div class="serv-kpi"><div class="serv-kpi-label">Categoría Más Costosa</div><div class="serv-kpi-val" style="color:'+(SUBCAT_COLORS[topSubcat]||'#f59e0b')+'">'+topSubcat+'</div><div class="serv-kpi-sub">'+fmt(subcatTotals[topSubcat]||0)+'</div></div>';

  var sortedRanches=RANCH_ORDER.filter(function(r){return ranchData[r];});
  if(!sortedRanches.length) sortedRanches=Object.keys(ranchData);
  document.getElementById('servRanchGrid').innerHTML=sortedRanches.length?
    sortedRanches.map(function(ranch){
      var col=RANCH_COLORS[ranch]||'#888', total=ranchTotals[ranch]||0;
      var subcats=ranchData[ranch]||{};
      var rows=(subcatFilter?[subcatFilter]:SUBCATS_ORDER).filter(function(sc){return sc in subcats;}).map(function(sc){
        var v=subcats[sc]||0, ic=SUBCAT_ICONS[sc]||'•', iCol=SUBCAT_COLORS[sc]||'#888';
        return '<div class="serv-subcat-row">'+
          '<div class="serv-subcat-icon" style="background:'+iCol+'22;color:'+iCol+'">'+ic+'</div>'+
          '<div class="serv-subcat-name">'+sc+'</div>'+
          '<div class="serv-subcat-val">'+fmt(v)+'</div></div>';
      }).join('');
      return '<div class="serv-ranch-card">'+
        '<div class="serv-ranch-hdr" style="background:'+col+'">'+
        '<div class="serv-ranch-name">'+ranch+'</div>'+
        '<div class="serv-ranch-total">'+fmt(total)+'</div></div>'+
        '<div class="serv-subcat-list">'+rows+'</div>'+
        '<div class="serv-total-row"><span>TOTAL:</span><span>'+fmt(total)+'</span></div></div>';
    }).join(''):
    '<div class="no-data" style="padding:30px;grid-column:1/-1">Sin datos para esta semana</div>';

  var maxV=Math.max.apply(null,Object.values(ranchTotals).concat([1]));
  document.getElementById('servRanchList').innerHTML=
    RANCH_ORDER.filter(function(r){return ranchTotals[r]>0;}).sort(function(a,b){return ranchTotals[b]-ranchTotals[a];}).map(function(r){
      var v=ranchTotals[r]||0, col=RANCH_COLORS[r]||'#888';
      return '<div class="ranch-row">'+
        '<div class="ranch-lbl" style="color:'+col+'">'+r+'</div>'+
        '<div class="ranch-bar-outer"><div class="ranch-bar-inner" style="width:'+(v/maxV*100).toFixed(1)+'%;background:'+col+'"></div></div>'+
        '<div class="ranch-usd">'+fmt(v)+'</div>'+
        '<div class="ranch-pct">'+(grandTotal>0?(v/grandTotal*100).toFixed(1):'0')+'%</div></div>';
    }).join('')+
    (grandTotal>0?'<div style="margin-top:8px;padding-top:8px;border-top:1px solid var(--border);display:flex;justify-content:space-between;font-family:IBM Plex Mono,monospace;font-size:.7rem;font-weight:700;color:var(--green)"><span>TOTAL</span><span>'+fmt(grandTotal)+'</span></div>':'');

  destroyChart('chartServ');
  var pieR=RANCH_ORDER.filter(function(r){return ranchTotals[r]>0;});
  if(pieR.length){
    Plotly.newPlot('chartServ',[{type:'pie',labels:pieR,
      values:pieR.map(function(r){return ranchTotals[r]||0;}),
      marker:{colors:pieR.map(function(r){return RANCH_COLORS[r]||'#888';})},
      textinfo:'label+percent',hole:0.35,
      hovertemplate:'<b>%{label}</b><br>%{text}<extra></extra>',
      text:pieR.map(function(r){return fmt(ranchTotals[r]||0);})
    }],plotlyLayout({margin:{t:35,r:30,b:15,l:50},height:200,showlegend:false}),plotlyCfg());
  }

  var section=document.getElementById('serviciosSection');
  section.classList.add('show');
  section.style.display='block';
  // Ocultar la tabla principal mientras se ve COSTO DE SERVICIOS
  var t=document.getElementById('swTableCard');
  if(t) t.style.display='none';
  setTimeout(function(){section.scrollIntoView({behavior:'smooth',block:'nearest'});},100);
}

function closeServicios(){
  var section=document.getElementById('serviciosSection');
  section.classList.remove('show');
  setTimeout(function(){section.style.display='none';},300);
  // Restaurar la tabla principal
  var t=document.getElementById('swTableCard');
  if(t) t.style.display='';
  // Restaurar las 3 pestañas
  document.getElementById('vtAnual').style.display='';
  document.getElementById('vtTendencia').style.display='';
  document.getElementById('vtSemana').style.display='';
  var el=document.getElementById('servSelect');
  if(el) el.selectedIndex=0;
}

// ═══════════════════════════════════════════
// PLOTLY HELPERS
// ═══════════════════════════════════════════
function plotlyLayout(extra){
  var base={
    paper_bgcolor:'rgba(0,0,0,0)',
    plot_bgcolor:'rgba(0,0,0,0)',
    font:{family:'IBM Plex Mono, monospace',color:'#64748b',size:11},
    margin:{t:10,r:10,b:40,l:60},
    xaxis:{gridcolor:'#f1f5f9',linecolor:'#e2e8f0',tickfont:{size:10}},
    yaxis:{gridcolor:'#f1f5f9',linecolor:'#e2e8f0',tickfont:{size:10},tickformat:'$,.0f'},
    hoverlabel:{bgcolor:'#fff',bordercolor:'#e2e8f0',font:{family:'IBM Plex Mono',size:12,color:'#1e293b'}},
    legend:{font:{size:10},bgcolor:'rgba(255,255,255,0.8)',bordercolor:'#e2e8f0',borderwidth:1},
    autosize:true
  };
  if(extra) Object.keys(extra).forEach(function(k){
    if(typeof extra[k]==='object'&&!Array.isArray(extra[k])&&base[k]){
      Object.assign(base[k],extra[k]);
    } else { base[k]=extra[k]; }
  });
  return base;
}
function plotlyCfg(){
  return {responsive:true,displayModeBar:true,
    modeBarButtonsToRemove:['select2d','lasso2d','autoScale2d'],
    displaylogo:false};
}

window.onerror = function(msg, src, line, col, err) {
  document.getElementById('loader').innerHTML =
    '<div style="color:#dc2626;font-family:IBM Plex Mono,monospace;padding:20px;max-width:600px;word-break:break-all;background:#fff;border-radius:12px;border:1px solid #fecaca">' +
    '<b>ERROR JS:</b><br>' + msg + '<br><small>línea ' + line + '</small></div>';
  return true;
};

// ═══════════════════════════════════════════
// ARRANCAR CON DATOS YA LISTOS
// ═══════════════════════════════════════════
// Reconstruir weekly_series desde weekly_detail si no existe
if(!DATA.weekly_series){
  DATA.weekly_series={};
  DATA.categories.forEach(function(cat){ DATA.weekly_series[cat]={}; });
  DATA.weekly_detail.forEach(function(r){
    if(r.usd_total>0){
      if(!DATA.weekly_series[r.categoria]) DATA.weekly_series[r.categoria]={};
      var key=r.year+'-W'+String(r.week).padStart(2,'0');
      DATA.weekly_series[r.categoria][key]=(DATA.weekly_series[r.categoria][key]||0)+r.usd_total;
    }
  });
}
inicializar();
</script>

<script>
function reportHeight() {
  var h = document.getElementById('app') 
    ? document.getElementById('app').getBoundingClientRect().bottom + window.scrollY + 40
    : document.body.scrollHeight + 40;
  window.parent.postMessage({type:'streamlit:setFrameHeight', height: Math.max(h, 600)}, '*');
}
var ro = new ResizeObserver(reportHeight);
ro.observe(document.body);
reportHeight();
setInterval(reportHeight, 300);
</script>
</body>
</html>"""

# Inyectar los datos JSON en el HTML
html_final = HTML.replace('__DATA_JSON__', data_json).replace('__FLORES_IMG__', _flores_b64)

# Renderizar
components.html(html_final, height=3500, scrolling=False)
