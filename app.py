 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/app.py b/app.py
index b573cb02e08096eb50cc46a2dcba804695420dae..decd9f4598f083dc6e197ec9aff0ea71bed2a00e 100644
--- a/app.py
+++ b/app.py
@@ -309,51 +309,51 @@ body{background:var(--bg);color:var(--text);font-family:'Syne',sans-serif;min-he
         </div>
       </div>
     </div>
   </div>
 </div>
 
 <!-- VIEW: POR SEMANA -->
 <div id="viewSemana" style="display:none">
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
     <div class="card">
       <div class="card-hdr"><span class="card-title">Tabla — Misma Semana, Distintos Años</span><span class="card-note" id="swTableNote">USD</span></div>
       <div class="table-scroll-wrap">
         <div class="scroll-hint show" id="hintSemana">← desliza →</div>
         <div class="scroll-x scroll-fade" id="wrapSemana">
           <table class="data-table">
-            <thead><tr>
+            <thead><tr id="swTableHead">
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
@@ -401,54 +401,62 @@ body{background:var(--bg);color:var(--text);font-family:'Syne',sans-serif;min-he
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
   view:'anual', weekIdx:0, fromWeek:1, toWeek:52
 };
 var allWeeks = [];
 var YEAR_COLORS = {2021:'#4ecdc4',2022:'#f7dc6f',2023:'#82e0aa',2024:'#f0b429',2025:'#00c97d',2026:'#ff6b6b'};
 var RANCH_COLORS = {
   'Prop-RM':'#00c97d','PosCo-RM':'#3b9eff','Campo-RM':'#f0b429',
   'Isabela':'#c084fc','HOOPS':'#fb923c','Cecilia':'#f472b6',
   'Cecilia 25':'#34d399','Christina':'#60a5fa','Albahaca-RM':'#a78bfa','Campo-VI':'#94a3b8'
 };
 var RANCH_ORDER = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','HOOPS','Cecilia','Cecilia 25','Christina','Albahaca-RM','Campo-VI'];
 var KEY_RANCHES = ['Prop-RM','PosCo-RM','Campo-RM','Isabela','Cecilia','Cecilia 25','Christina'];
+var CAT_MIRFE = 'FERTILIZANTES';
+var CAT_MIPE = 'DESINFECCION / PLAGUICIDAS';
 var charts = {};
 
 function recargar() { window.location.reload(); }
 
+function catLabel(cat){
+  if(cat===CAT_MIRFE) return 'FERTILIZANTES (MIRFE)';
+  if(cat===CAT_MIPE) return 'DESINFECCION / PLAGUICIDAS (MIPE)';
+  return cat;
+}
+
 // ═══════════════════════════════════════════
 // INICIALIZAR
 // ═══════════════════════════════════════════
 function inicializar() {
   var years = DATA.years, cats = DATA.categories;
   var prefCat = 'MATERIAL DE EMPAQUE';
   state.cat = cats.indexOf(prefCat) > -1 ? prefCat : cats[0];
   state.activeYears = {};
   years.forEach(function(y){ state.activeYears[y] = true; });
 
   var wSet = {};
   DATA.weekly_detail.forEach(function(r){ wSet[r.week] = 1; });
   allWeeks = Object.keys(wSet).map(Number).sort(function(a,b){return a-b;});
   state.weekIdx = allWeeks.length - 1;
 
   // Semana actual y anterior: tomar del año más reciente con datos
   var latestYear = DATA.years[DATA.years.length-1];
   var weeksOfLatest = DATA.weekly_detail
     .filter(function(r){ return r.year === latestYear; })
     .map(function(r){ return r.week; })
     .filter(function(v,i,a){ return a.indexOf(v)===i; })
     .sort(function(a,b){ return a-b; });
   var curWeek  = weeksOfLatest[weeksOfLatest.length-1] || allWeeks[allWeeks.length-1] || 1;
   var prevWeek2 = weeksOfLatest[weeksOfLatest.length-2] || weeksOfLatest[0] || curWeek;
   state.fromWeek = prevWeek2;
@@ -473,83 +481,95 @@ function fmt(n) {
 function pct(a,b){ return (!b||b===0)?null:((a-b)/b*100).toFixed(1); }
 function destroyChart(id){ if(charts[id]){charts[id].destroy();delete charts[id];} }
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
+function mergeAgg(a,b){
+  var out={usd:0,mxn:0,ranches:{},ranches_mxn:{},date_range:''};
+  [a,b].forEach(function(src){
+    if(!src) return;
+    out.usd+=(src.usd||0); out.mxn+=(src.mxn||0);
+    if(src.date_range && !out.date_range) out.date_range=src.date_range;
+    Object.keys(src.ranches||{}).forEach(function(rn){out.ranches[rn]=(out.ranches[rn]||0)+src.ranches[rn];});
+    Object.keys(src.ranches_mxn||{}).forEach(function(rn){out.ranches_mxn[rn]=(out.ranches_mxn[rn]||0)+src.ranches_mxn[rn];});
+  });
+  out.usd=Math.round(out.usd*100)/100; out.mxn=Math.round(out.mxn*100)/100;
+  return out;
+}
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
     var recs=getDetail(cat,undefined,yr).filter(function(r){return r.week>=fromW&&r.week<=toW;});
     if(!recs.length) return;
     var ag=aggregateDetail(recs);
     ag.weekly={};
     recs.forEach(function(r){ag.weekly[r.week]=(ag.weekly[r.week]||0)+r.usd_total;});
     res[yr]=ag;
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
 function buildCatSelect(){
   var el=document.getElementById('catSelect');
   el.innerHTML=DATA.categories.map(function(c){
-    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+'>'+c+'</option>';
+    return '<option value="'+c.replace(/"/g,'&quot;')+'"'+(c===state.cat?' selected':'')+'>'+catLabel(c)+'</option>';
   }).join('');
   document.getElementById('catCount').textContent=(DATA.categories.indexOf(state.cat)+1)+' / '+DATA.categories.length;
 }
 function buildYearChips(){
   var el=document.getElementById('yearChips');
   el.innerHTML=DATA.years.map(function(y){
     var col=YEAR_COLORS[y]||'#888', on=state.activeYears[y]?'on':'';
     return '<button class="yr-chip '+on+'" style="color:'+col+';background:'+(on?col+'22':'transparent')+';border-color:'+col+'44" onclick="toggleYear('+y+')">'+y+'</button>';
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
@@ -756,65 +776,87 @@ function renderStack(){
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
   var weekNum=allWeeks[state.weekIdx], yrs=activeYrList(), byYear=getWeekByYear(state.cat,weekNum);
-  document.getElementById('swTableNote').textContent=state.currency==='usd'?'USD':'MXN';
+  var showChemPair=(state.cat===CAT_MIRFE || state.cat===CAT_MIPE);
+  var byYearMirfe=showChemPair?getWeekByYear(CAT_MIRFE,weekNum):{};
+  var byYearMipe=showChemPair?getWeekByYear(CAT_MIPE,weekNum):{};
+  var sym=state.currency==='usd'?'USD':'MXN';
+  document.getElementById('swTableNote').textContent=showChemPair?(sym+' · QUIMICOS: MIRFE + MIPE'):sym;
+  document.getElementById('swTableHead').innerHTML=showChemPair
+    ? '<th>Año</th><th>Semana</th><th>Fecha</th><th>Total Químicos ('+sym+')</th><th>MIRFE ('+sym+')</th><th>MIPE ('+sym+')</th><th>Δ vs mismo año ant.</th>'+
+      '<th>Prop-RM</th><th>PosCo-RM</th><th>Campo-RM</th><th>Isabela</th><th>Cecilia</th><th>Cecilia 25</th><th>Christina</th>'
+    : '<th>Año</th><th>Semana</th><th>Fecha</th><th>Total '+sym+'</th><th>Δ vs mismo año ant.</th>'+
+      '<th>Prop-RM</th><th>PosCo-RM</th><th>Campo-RM</th><th>Isabela</th><th>Cecilia</th><th>Cecilia 25</th><th>Christina</th>';
   document.getElementById('swTableBody').innerHTML=yrs.map(function(yr,i){
-    var d=byYear[yr], val=d?(state.currency==='usd'?d.usd:d.mxn):0;
-    var prevD=i>0?byYear[yrs[i-1]]:null, prevVal=prevD?(state.currency==='usd'?prevD.usd:prevD.mxn):null;
+    var d=byYear[yr];
+    if(showChemPair){
+      d=mergeAgg(byYearMirfe[yr], byYearMipe[yr]);
+    }
+    var val=d?(state.currency==='usd'?d.usd:d.mxn):0;
+    var prevD=null;
+    if(i>0){
+      prevD=showChemPair?mergeAgg(byYearMirfe[yrs[i-1]], byYearMipe[yrs[i-1]]):byYear[yrs[i-1]];
+    }
+    var prevVal=prevD?(state.currency==='usd'?prevD.usd:prevD.mxn):null;
     var delta=prevVal!==null?pct(val,prevVal):null;
     var col=YEAR_COLORS[yr]||'#888';
     var dStr=delta!==null?'<span class="'+(parseFloat(delta)>0?'chg-pos':'chg-neg')+'">'+(parseFloat(delta)>0?'+':'')+delta+'%</span>':'<span class="chg-0">—</span>';
     var ranchSrc=d?(state.currency==='usd'?d.ranches:d.ranches_mxn):{};
     var cells=KEY_RANCHES.map(function(r){var v=ranchSrc[r]||0;return '<td style="color:'+(v>0?(RANCH_COLORS[r]||'#888')+'cc':'#3a5a48')+'">'+(v>0?fmt(v):'—')+'</td>';}).join('');
+    var mirfeVal=showChemPair&&byYearMirfe[yr]?(state.currency==='usd'?byYearMirfe[yr].usd:byYearMirfe[yr].mxn):0;
+    var mipeVal=showChemPair&&byYearMipe[yr]?(state.currency==='usd'?byYearMipe[yr].usd:byYearMipe[yr].mxn):0;
+    var chemCells=showChemPair
+      ? '<td style="font-weight:600;color:#34d399">'+(mirfeVal?fmt(mirfeVal):'—')+'</td><td style="font-weight:600;color:#60a5fa">'+(mipeVal?fmt(mipeVal):'—')+'</td>'
+      : '';
     return '<tr>'+
       '<td><span class="yr-dot" style="background:'+col+'"></span><strong style="color:'+col+'">'+yr+'</strong></td>'+
       '<td style="color:'+col+'">'+wFmt(weekNum)+'</td>'+
       '<td style="color:var(--dim);font-size:.65rem">'+(d&&d.date_range?d.date_range:'—')+'</td>'+
       '<td style="color:'+col+';font-weight:600">'+fmt(val)+'</td>'+
-      '<td>'+dStr+'</td>'+cells+'</tr>';
+      chemCells+'<td>'+dStr+'</td>'+cells+'</tr>';
   }).join('');
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
     return '<div class="stat-box" style="border-color:'+col+'33">'+
       '<div class="stat-label">'+yr+' · '+sym+'</div>'+
       '<div class="stat-val" style="color:'+col+'">'+fmt(val)+'</div>'+
       '<div style="font-size:.62rem;color:var(--dim);font-family:IBM Plex Mono,monospace">'+fmt(avg)+'/sem · '+wks+' semanas</div></div>';
   }).join('');
 
   var rangeWeeks=allWeeks.filter(function(w){return w>=f&&w<=t;}), rLabels=rangeWeeks.map(wFmt);
 
   destroyChart('rangeLine');
   var rDatasets=yrs.map(function(yr){
     var d=byYear[yr], data=rangeWeeks.map(function(w){return d&&d.weekly[w]?d.weekly[w]:null;}), col=YEAR_COLORS[yr]||'#888';
 
EOF
)
