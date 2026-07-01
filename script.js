/* ============================================================
   Inventra - Inventory Analytics Dashboard
   Vanilla JS + Chart.js. Reads data.json built from the
   uploaded Excel workbooks (Ageing sheets).
   ============================================================ */

(() => {
  'use strict';

   // ===============================
// Excel Upload Engine
// ===============================

let month1Data = [];

let month2Data = [];

   // Combined Dataset

let allData = [];

let month1File = null;

let month2File = null;

  // ---------- Globals ----------
  const PALETTE = ['#1e5fff','#0ea5e9','#10b981','#f59e0b','#ef4444','#8b5cf6','#6366f1','#14b8a6','#f97316','#ec4899'];
  const AGE_ORDER = ['<30days','<60days','< 90 days','< 180 days','<1 Year','>1 Year','>2 Year'];
  const AGE_COLORS = {
    '<30days':'#10b981','<60days':'#22c55e','< 90 days':'#0ea5e9',
    '< 180 days':'#6366f1','<1 Year':'#f59e0b','>1 Year':'#f97316','>2 Year':'#ef4444'
  };
  const STATE = {
    raw: [],
    filtered: [],
    filters: { customer:'', age:'', program:'', source:'', material:'' },
    table: { page:1, pageSize:15, sortKey:'Closing Value', sortDir:'desc', search:'' },
    charts: {}
  };

  // ---------- Utilities ----------
  const fmtINR = n => {
    if (n === null || n === undefined || isNaN(n)) return '—';
    const abs = Math.abs(n);
    if (abs >= 1e7) return '₹' + (n/1e7).toFixed(2) + ' Cr';
    if (abs >= 1e5) return '₹' + (n/1e5).toFixed(2) + ' L';
    if (abs >= 1e3) return '₹' + (n/1e3).toFixed(1) + 'K';
    return '₹' + Math.round(n).toLocaleString('en-IN');
  };
  const fmtNum = n => (n == null || isNaN(n)) ? '—' : Number(n).toLocaleString('en-IN');
  const fmtPct = n => (n==null||isNaN(n)) ? '—' : n.toFixed(1) + '%';
  const uniq = arr => Array.from(new Set(arr.filter(v => v !== null && v !== undefined && v !== ''))).sort();
  const sum = (arr, k) => arr.reduce((a,r) => a + (Number(r[k])||0), 0);
  const $ = s => document.querySelector(s);
  const $$ = s => document.querySelectorAll(s);

  // ---------- Data loading ----------
  async function loadData() {
    const res = await fetch('data.json');
    if (!res.ok) throw new Error('Failed to load data');
    const rows = await res.json();
    STATE.raw = rows.map(r => ({
      ...r,
      'Closing Stock': Number(r['Closing Stock']) || 0,
      'Closing Value': Number(r['Closing Value']) || 0,
    }));
  }

  // ---------- Filter pipeline ----------
  function applyFilters() {
    const f = STATE.filters;
    const m = f.material.toLowerCase().trim();
    STATE.filtered = STATE.raw.filter(r => {
      if (f.customer && r.Customer !== f.customer) return false;
      if (f.age && r['Ageing- Consm'] !== f.age) return false;
      if (f.program && r.Progm !== f.program) return false;
      if (f.source && r.Source !== f.source) return false;
      if (m) {
        const hay = ((r.Material||'') + ' ' + (r['Material Description']||'')).toLowerCase();
        if (!hay.includes(m)) return false;
      }
      return true;
    });
    STATE.table.page = 1;
    refreshAll();
  }

  function resetFilters() {
    STATE.filters = { customer:'', age:'', program:'', source:'', material:'' };
    $('#filterCustomer').value = '';
    $('#filterAge').value = '';
    $('#filterProgram').value = '';
    $('#filterSource').value = '';
    $('#filterMaterial').value = '';
    applyFilters();
  }

  // ---------- KPIs ----------
  function animateCounter(el, to, formatter) {
    const start = 0, dur = 900, t0 = performance.now();
    const step = t => {
      const p = Math.min(1, (t - t0) / dur);
      const eased = 1 - Math.pow(1 - p, 3);
      const val = start + (to - start) * eased;
      el.textContent = formatter(val);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  function renderKPIs() {
    const rows = STATE.filtered;
    const totalValue = sum(rows, 'Closing Value');
    const totalMaterials = new Set(rows.map(r => r.Material)).size;
    const slow = rows.filter(r => ['< 180 days','<1 Year','>1 Year','>2 Year'].includes(r['Ageing- Consm']));
    const dead = rows.filter(r => r['Ageing- Consm'] === '>2 Year');
    const fresh = rows.filter(r => ['<30days','<60days','< 90 days'].includes(r['Ageing- Consm']));
    const freshValue = sum(fresh, 'Closing Value');
    const health = totalValue > 0 ? (freshValue / totalValue) * 100 : 0;

    animateCounter($('#kpiValue'), totalValue, fmtINR);
    animateCounter($('#kpiMaterials'), totalMaterials, v => fmtNum(Math.round(v)));
    animateCounter($('#kpiSlow'), new Set(slow.map(r=>r.Material)).size, v => fmtNum(Math.round(v)));
    animateCounter($('#kpiDead'), new Set(dead.map(r=>r.Material)).size, v => fmtNum(Math.round(v)));
    animateCounter($('#kpiHealth'), health, v => v.toFixed(1) + '%');

    // Deltas vs raw baseline
    const baseTotal = sum(STATE.raw, 'Closing Value');
    const deltaPct = baseTotal > 0 ? ((totalValue - baseTotal)/baseTotal)*100 : 0;
    $('#kpiValueDelta').textContent = (deltaPct>=0?'+':'') + deltaPct.toFixed(1) + '% vs all';
    $('#kpiMaterialsDelta').textContent = totalMaterials + ' SKUs';
    const slowShare = rows.length ? (slow.length/rows.length)*100 : 0;
    $('#kpiSlowDelta').textContent = slowShare.toFixed(1) + '% of rows';
    const deadValue = sum(dead, 'Closing Value');
    $('#kpiDeadDelta').textContent = fmtINR(deadValue);
    $('#kpiHealthDelta').textContent = (100-health).toFixed(1) + '% aged';
  }

  // ---------- Charts ----------
  const defaultOpts = {
    responsive:true, maintainAspectRatio:false,
    plugins:{
      legend:{ position:'bottom', labels:{ usePointStyle:true, padding:14, font:{family:'Inter',size:11} } },
      tooltip:{ backgroundColor:'#0f172a', padding:12, cornerRadius:10, titleFont:{family:'Inter',weight:'600'}, bodyFont:{family:'Inter'} }
    },
    animation:{ duration:800, easing:'easeOutCubic' }
  };

  function makeChart(id, config) {
    const ctx = document.getElementById(id);
    if (STATE.charts[id]) STATE.charts[id].destroy();
    STATE.charts[id] = new Chart(ctx, config);
  }

  function renderCharts() {
    const rows = STATE.filtered;

    // 1) Ageing donut (count of records)
    const ageCount = {};
    AGE_ORDER.forEach(a => ageCount[a] = 0);
    rows.forEach(r => { const a=r['Ageing- Consm']; if(a in ageCount) ageCount[a]++; });
    makeChart('chartAgeingDonut', {
      type:'doughnut',
      data:{ labels:AGE_ORDER, datasets:[{ data:AGE_ORDER.map(a=>ageCount[a]), backgroundColor:AGE_ORDER.map(a=>AGE_COLORS[a]), borderWidth:2, borderColor:'#fff', hoverOffset:8 }] },
      options:{ ...defaultOpts, cutout:'65%' }
    });

    // 2) Value by age (bar)
    const ageValue = {}; AGE_ORDER.forEach(a => ageValue[a]=0);
    rows.forEach(r => { const a=r['Ageing- Consm']; if(a in ageValue) ageValue[a] += Number(r['Closing Value'])||0; });
    makeChart('chartAgeBar', {
      type:'bar',
      data:{ labels:AGE_ORDER, datasets:[{ label:'Closing Value (INR)', data:AGE_ORDER.map(a=>ageValue[a]), backgroundColor:AGE_ORDER.map(a=>AGE_COLORS[a]), borderRadius:8, maxBarThickness:48 }] },
      options:{ ...defaultOpts, plugins:{...defaultOpts.plugins, legend:{display:false}, tooltip:{...defaultOpts.plugins.tooltip, callbacks:{ label:c=>'Value: '+fmtINR(c.parsed.y) }}}, scales:{ y:{ ticks:{ callback:v=>fmtINR(v), font:{family:'Inter'} }, grid:{color:'#eef2f7'} }, x:{ grid:{display:false}, ticks:{font:{family:'Inter'}} } } }
    });

    // 3) Top customers
    const byCust = {};
    rows.forEach(r => { const c=r.Customer||'Unknown'; byCust[c]=(byCust[c]||0)+(Number(r['Closing Value'])||0); });
    const topCust = Object.entries(byCust).sort((a,b)=>b[1]-a[1]).slice(0,10);
    makeChart('chartCustomers', {
      type:'bar',
      data:{ labels:topCust.map(t=>t[0]), datasets:[{ label:'Inventory Value', data:topCust.map(t=>t[1]), backgroundColor:PALETTE, borderRadius:8 }] },
      options:{ ...defaultOpts, indexAxis:'y', plugins:{...defaultOpts.plugins,legend:{display:false}, tooltip:{...defaultOpts.plugins.tooltip, callbacks:{label:c=>fmtINR(c.parsed.x)}}}, scales:{ x:{ ticks:{callback:v=>fmtINR(v),font:{family:'Inter'}},grid:{color:'#eef2f7'} }, y:{grid:{display:false},ticks:{font:{family:'Inter'}}} } }
    });

    // 4) Top materials
    const byMat = {};
    rows.forEach(r => { const k=(r.Material||'?') + ' · ' + (r['Material Description']||'').slice(0,28); byMat[k]=(byMat[k]||0)+(Number(r['Closing Value'])||0); });
    const topMat = Object.entries(byMat).sort((a,b)=>b[1]-a[1]).slice(0,10);
    makeChart('chartMaterials', {
      type:'bar',
      data:{ labels:topMat.map(t=>t[0]), datasets:[{ label:'Closing Value', data:topMat.map(t=>t[1]), backgroundColor:PALETTE.slice().reverse(), borderRadius:8 }] },
      options:{ ...defaultOpts, indexAxis:'y', plugins:{...defaultOpts.plugins,legend:{display:false}, tooltip:{...defaultOpts.plugins.tooltip, callbacks:{label:c=>fmtINR(c.parsed.x)}}}, scales:{ x:{ ticks:{callback:v=>fmtINR(v),font:{family:'Inter'}},grid:{color:'#eef2f7'} }, y:{grid:{display:false},ticks:{font:{family:'Inter',size:10}}} } }
    });

    // 5) Program pie
    const byProg = {};
    rows.forEach(r => { const p=r.Progm||'Other'; byProg[p]=(byProg[p]||0)+(Number(r['Closing Value'])||0); });
    const progEntries = Object.entries(byProg).sort((a,b)=>b[1]-a[1]).slice(0,8);
    makeChart('chartPrograms', {
      type:'pie',
      data:{ labels:progEntries.map(t=>t[0]), datasets:[{ data:progEntries.map(t=>t[1]), backgroundColor:PALETTE, borderColor:'#fff', borderWidth:2 }] },
      options:{ ...defaultOpts, plugins:{...defaultOpts.plugins, tooltip:{...defaultOpts.plugins.tooltip, callbacks:{label:c=>c.label+': '+fmtINR(c.parsed)}}} }
    });

    // 6) Trend by source (snapshot) - if both periods present, draw line
    const bySource = {};
    rows.forEach(r => { const s=r.Source||'Unknown'; bySource[s]=(bySource[s]||0)+(Number(r['Closing Value'])||0); });
    // Order chronologically when possible
    const periodOrder = ['Jan 2026','Feb 2026','Mar 2026','Apr 2026'];
    const labels = periodOrder.filter(p=>p in bySource).concat(Object.keys(bySource).filter(p=>!periodOrder.includes(p)));
    makeChart('chartTrend', {
      type:'line',
      data:{ labels, datasets:[{
        label:'Closing Value',
        data:labels.map(l=>bySource[l]||0),
        borderColor:'#1e5fff', backgroundColor:'rgba(30,95,255,.12)',
        fill:true, tension:.4, pointRadius:6, pointHoverRadius:8,
        pointBackgroundColor:'#fff', pointBorderColor:'#1e5fff', pointBorderWidth:3
      }] },
      options:{ ...defaultOpts, plugins:{...defaultOpts.plugins,legend:{display:false}, tooltip:{...defaultOpts.plugins.tooltip, callbacks:{label:c=>fmtINR(c.parsed.y)}}}, scales:{ y:{ ticks:{callback:v=>fmtINR(v),font:{family:'Inter'}},grid:{color:'#eef2f7'} }, x:{grid:{display:false},ticks:{font:{family:'Inter'}}} } }
    });
  }

  // ---------- Insights ----------
  function renderInsights() {
    const rows = STATE.filtered;
    const grid = $('#insightsGrid');
    grid.innerHTML = '';

    const dead = rows.filter(r => r['Ageing- Consm']==='>2 Year');
    const deadVal = sum(dead, 'Closing Value');
    const aged1y = rows.filter(r => ['>1 Year','>2 Year'].includes(r['Ageing- Consm']));
    const agedVal = sum(aged1y, 'Closing Value');
    const noConsumption = rows.filter(r => !r['Last Consumptn']);
    const noConsValue = sum(noConsumption,'Closing Value');

    const byCust = {};
    rows.forEach(r => { const c=r.Customer||'Unknown'; byCust[c]=(byCust[c]||0)+(Number(r['Closing Value'])||0); });
    const topCust = Object.entries(byCust).sort((a,b)=>b[1]-a[1])[0];

    const topItem = rows.slice().sort((a,b)=>(b['Closing Value']||0)-(a['Closing Value']||0))[0];
    const fresh = rows.filter(r => ['<30days','<60days'].includes(r['Ageing- Consm']));
    const freshShare = rows.length ? (fresh.length/rows.length)*100 : 0;

    const insights = [
      { type:'danger', icon:'fa-skull-crossbones', title:'Dead Stock (>2 Years)',
        body:`<span class="value">${fmtNum(dead.length)}</span> records worth <span class="value">${fmtINR(deadVal)}</span> have not moved in over 2 years and are prime liquidation candidates.` },
      { type:'warn', icon:'fa-triangle-exclamation', title:'Aged Inventory (>1 Year)',
        body:`<span class="value">${fmtNum(aged1y.length)}</span> records totalling <span class="value">${fmtINR(agedVal)}</span> are aged beyond 12 months. Recommend a structured action plan.` },
      { type:'warn', icon:'fa-clock-rotate-left', title:'No Recent Consumption',
        body:`<span class="value">${fmtNum(noConsumption.length)}</span> SKUs (${fmtINR(noConsValue)}) have no last-consumption date recorded — review for obsolescence.` },
      { type:'default', icon:'fa-user-tie', title:'Highest-Value Customer',
        body: topCust ? `<span class="value">${topCust[0]}</span> holds the largest exposure at <span class="value">${fmtINR(topCust[1])}</span> in inventory value.` : 'No customer data in current view.' },
      { type:'default', icon:'fa-gem', title:'Top Inventory Item',
        body: topItem ? `<span class="value">${topItem.Material}</span> — ${topItem['Material Description']||''} valued at <span class="value">${fmtINR(topItem['Closing Value'])}</span>.` : 'No items in current view.' },
      { type:'success', icon:'fa-leaf', title:'Fresh Inventory Share',
        body:`<span class="value">${freshShare.toFixed(1)}%</span> of records are under 60 days old — healthy turnover signal.` },
    ];

    insights.forEach(i => {
      const div = document.createElement('div');
      div.className = 'insight ' + (i.type==='default' ? '' : i.type);
      div.innerHTML = `<i class="fa-solid ${i.icon} lead"></i><div class="insight-body"><strong>${i.title}</strong><p>${i.body}</p></div>`;
      grid.appendChild(div);
    });
  }

  // ---------- Table ----------
  function renderTable() {
    const t = STATE.table;
    let rows = STATE.filtered.slice();

    if (t.search) {
      const q = t.search.toLowerCase();
      rows = rows.filter(r => Object.values(r).some(v => String(v||'').toLowerCase().includes(q)));
    }

    rows.sort((a,b) => {
      const va = a[t.sortKey], vb = b[t.sortKey];
      const na = Number(va), nb = Number(vb);
      const numeric = !isNaN(na) && !isNaN(nb) && va !== null && vb !== null && va !== '' && vb !== '';
      let cmp = numeric ? na - nb : String(va||'').localeCompare(String(vb||''));
      return t.sortDir === 'asc' ? cmp : -cmp;
    });

    const total = rows.length;
    const pages = Math.max(1, Math.ceil(total / t.pageSize));
    t.page = Math.min(t.page, pages);
    const start = (t.page-1) * t.pageSize;
    const slice = rows.slice(start, start + t.pageSize);

    const tbody = $('#invTable tbody');
    tbody.innerHTML = slice.map(r => {
      const ageBadgeClass = (r['Ageing- Consm']==='>2 Year') ? 'red' : (r['Ageing- Consm']==='>1 Year' || r['Ageing- Consm']==='<1 Year') ? 'amber' : 'green';
      return `<tr>
        <td><strong>${r.Material||'—'}</strong></td>
        <td title="${(r['Material Description']||'').replace(/"/g,'&quot;')}">${r['Material Description']||'—'}</td>
        <td>${r.Customer||'—'}</td>
        <td><span class="badge gray">${r.Progm||'—'}</span></td>
        <td class="num">${fmtNum(r['Closing Stock'])}</td>
        <td class="num"><strong>${fmtINR(r['Closing Value'])}</strong></td>
        <td><span class="badge ${ageBadgeClass}">${r['Ageing- Consm']||'—'}</span></td>
        <td>${r['Last Consumptn']||'—'}</td>
        <td>${r.Remarks||'—'}</td>
      </tr>`;
    }).join('') || `<tr><td colspan="9" style="text-align:center;padding:40px;color:var(--muted)">No records match the current filters.</td></tr>`;

    $('#tableMeta').textContent = `${fmtNum(total)} rows · sorted by ${t.sortKey} (${t.sortDir})`;
    $('#pageInfo').textContent = `Page ${t.page} / ${pages}`;
    $('#pagePrev').disabled = t.page <= 1;
    $('#pageNext').disabled = t.page >= pages;
  }

  // ---------- Filters UI ----------
  function populateFilters() {
    const customers = uniq(STATE.raw.map(r => r.Customer));
    const programs = uniq(STATE.raw.map(r => r.Progm));
    const sources = uniq(STATE.raw.map(r => r.Source));
    const ages = AGE_ORDER.filter(a => STATE.raw.some(r => r['Ageing- Consm']===a));

    const fill = (id, opts) => {
      const sel = $(id);
      opts.forEach(o => { const e=document.createElement('option'); e.value=o; e.textContent=o; sel.appendChild(e); });
    };
    fill('#filterCustomer', customers);
    fill('#filterProgram', programs);
    fill('#filterSource', sources);
    fill('#filterAge', ages);
  }

  // ---------- Export / Print ----------
  function exportCSV() {
    const rows = STATE.filtered;
    if (!rows.length) return;
    const cols = ['Material','Material Description','Customer','Progm','Type','Closing Stock','BUn','Closing Value','Ageing- Consm','Last Consumptn','Remarks','Source'];
    const esc = v => { v = v==null?'':String(v); return /[",\n]/.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; };
    const csv = [cols.join(',')].concat(rows.map(r => cols.map(c=>esc(r[c])).join(','))).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `inventory_export_${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
  }

  // ---------- Refresh orchestrator ----------
  function refreshAll() {
    renderKPIs();
    renderCharts();
    renderInsights();
    renderTable();
  }

  // ---------- Events ----------
  function bindEvents() {
    $('#filterCustomer').addEventListener('change', e => { STATE.filters.customer = e.target.value; applyFilters(); });
    $('#filterAge').addEventListener('change', e => { STATE.filters.age = e.target.value; applyFilters(); });
    $('#filterProgram').addEventListener('change', e => { STATE.filters.program = e.target.value; applyFilters(); });
    $('#filterSource').addEventListener('change', e => { STATE.filters.source = e.target.value; applyFilters(); });

    let t1; $('#filterMaterial').addEventListener('input', e => { clearTimeout(t1); t1=setTimeout(()=>{ STATE.filters.material = e.target.value; applyFilters(); }, 220); });
    let t2; $('#globalSearch').addEventListener('input', e => { clearTimeout(t2); t2=setTimeout(()=>{ STATE.filters.material = e.target.value; $('#filterMaterial').value=e.target.value; applyFilters(); }, 220); });
    let t3; $('#tableSearch').addEventListener('input', e => { clearTimeout(t3); t3=setTimeout(()=>{ STATE.table.search = e.target.value; renderTable(); }, 200); });

    $('#btnReset').addEventListener('click', resetFilters);
    $('#btnRefresh').addEventListener('click', () => { refreshAll(); });
    $('#btnPrint').addEventListener('click', () => window.print());
    $('#btnExport').addEventListener('click', exportCSV);

    $('#pagePrev').addEventListener('click', () => { STATE.table.page--; renderTable(); });
    $('#pageNext').addEventListener('click', () => { STATE.table.page++; renderTable(); });

    $$('#invTable thead th').forEach(th => th.addEventListener('click', () => {
      const k = th.dataset.sort; if (!k) return;
      if (STATE.table.sortKey === k) STATE.table.sortDir = STATE.table.sortDir==='asc'?'desc':'asc';
      else { STATE.table.sortKey = k; STATE.table.sortDir = 'desc'; }
      renderTable();
    }));

    
  $$('.nav-item').forEach(item => {

    item.addEventListener('click', function () {

        document.querySelectorAll(".nav-item").forEach(nav =>
            nav.classList.remove("active")
        );

        this.classList.add("active");

        let target = null;

        switch(this.dataset.section){

            case "dashboard":
                target = document.querySelector(".kpis");
                break;

            case "inventory":
                target = document.querySelector(".table-card");
                break;

            case "ageing":
                target = document.getElementById("chartAgeingDonut");
                break;

            case "customers":
                target = document.getElementById("chartCustomers");
                break;

            case "materials":
                target = document.getElementById("chartMaterials");
                break;

            case "liquidation":
                target = document.getElementById("insightsGrid");
                break;

            case "reports":
                target = document.querySelector(".table-card");
                break;

            case "settings":
                alert("Settings coming soon!");
                return;
        }

        if(target){
            target.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });
        }

    });

});
  }

  // ---------- Init ----------
  async function init() {
    $('#todayDate').textContent = new Date().toLocaleDateString('en-IN', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
    try {
      await loadData();
      populateFilters();
      bindEvents();
      STATE.filtered = STATE.raw.slice();
      refreshAll();
    } catch (err) {
      console.error(err);
      alert('Failed to load inventory data: ' + err.message);
    } finally {
      setTimeout(() => $('#loader').classList.add('hidden'), 300);
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
function readExcel(file){

    return new Promise((resolve,reject)=>{

        const reader = new FileReader();

        reader.onload=(e)=>{

            const data = new Uint8Array(e.target.result);

            const workbook = XLSX.read(data,{
                type:"array"
            });

            const sheetName=workbook.SheetNames[0];

            const worksheet=workbook.Sheets[sheetName];

            const json=XLSX.utils.sheet_to_json(worksheet);

            resolve(json);

        };

        reader.onerror=reject;

        reader.readAsArrayBuffer(file);

    });

}
document
.getElementById("month1File")
.addEventListener("change",function(){

    month1File=this.files[0];

    document
    .getElementById("month1Name")
    .textContent=month1File.name;

});
document
.getElementById("month2File")
.addEventListener("change",function(){

    month2File=this.files[0];

    document
    .getElementById("month2Name")
    .textContent=month2File.name;

});
document
.getElementById("btnLoadData")
.addEventListener("click",async()=>{

    if(!month1File || !month2File){

        alert("Please select both Excel files.");

        return;

    }

    month1Data=await readExcel(month1File);

    month2Data=await readExcel(month2File);
    allData = [...month1Data, ...month2Data];

    console.log(month1Data);

    console.log(month2Data);

    alert("Dashboard data loaded successfully!");

});

function getUniqueValues(data, field){

    return [...new Set(

        data
        .map(row => row[field])
        .filter(value =>
            value !== undefined &&
            value !== null &&
            value !== ""
        )

    )].sort();

}

function fillDropdown(id, values){

    const dropdown = document.getElementById(id);

    dropdown.innerHTML = "";

    const option = document.createElement("option");

    option.value = "All";

    option.textContent = "All";

    dropdown.appendChild(option);

    values.forEach(value => {

        const op = document.createElement("option");

        op.value = value;

        op.textContent = value;

        dropdown.appendChild(op);

    });

}
fillDropdown(

    "customerFilter",

    getUniqueValues(allData,"Customer")

);

fillDropdown(

    "programFilter",

    getUniqueValues(allData,"Progm")

);

fillDropdown(

    "materialFilter",

    getUniqueValues(allData,"Material")

);

fillDropdown(

    "typeFilter",

    getUniqueValues(allData,"Type")

);

fillDropdown(

    "ageFilter",

    getUniqueValues(allData,"Ageing")

);




