(() => {
  'use strict';

  const COL = { supplier: 0, status: 5, tech: 7, region: 14, sent: 15, active: 17 };
  const CORE = new Set(['ATTIVO', 'ATTIVAZIONE', 'KO']);
  const STORAGE_KEY = 'connectivity-dashboard-history-v1';
  const DAY = 86400000;
  const nf = new Intl.NumberFormat('it-IT');
  const df = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const $ = id => document.getElementById(id);
  let report = null;
  let sourceRows = [];
  let charts = {};
  let sortState = {};

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    $('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));
    $('historyInput').addEventListener('change', e => importHistory(e.target.files[0]));
    $('exportHistoryBtn').addEventListener('click', exportHistory);
    $('clearHistoryBtn').addEventListener('click', clearHistory);
    $('exportExcelBtn').addEventListener('click', exportExcel);
    $('minSample').addEventListener('change', () => { if (report) { renderAlerts(); renderAllTables(); } });
    document.querySelectorAll('.tab').forEach(btn => btn.addEventListener('click', () => activateTab(btn.dataset.tab)));
    const dz = $('dropZone');
    ['dragenter','dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
    ['dragleave','drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
    dz.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
    updateHistoryButtons();
  }

  async function handleFile(file) {
    hideMessage();
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) return showMessage('Formato non valido. Seleziona un file .xlsx o .xls.');
    if (typeof XLSX === 'undefined') return showMessage('Modulo Excel non disponibile. Riapri la cartella completa dell’app e riprova.');
    try {
      setBusy(true);
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, dense: false });
      const selected = selectUsefulSheet(workbook);
      if (!selected) throw new Error('Nessun foglio utile trovato. Verifica che un foglio contenga una riga di intestazione con valori nelle colonne A, F, H, O, P e R.');
      const validation = selected.validation;
      const analysisDate = analysisDateFromFilename(file.name);
      sourceRows = parseRows(selected.rows, validation.headerRow, analysisDate);
      if (!sourceRows.length) throw new Error('Il foglio non contiene righe dati utilizzabili dopo l’intestazione.');
      report = buildReport(sourceRows, analysisDate, file.name, selected.name, validation.headerRow);
      saveSnapshot(report.summary, analysisDate);
      renderDashboard();
      if (report.summary.analyzed === 0) {
        showMessage('Il file è stato letto, ma non contiene pratiche con Stato = Attivo, Attivazione o KO. Controlla i valori della colonna F.');
      }
    } catch (err) {
      report = null;
      $('dashboard').classList.add('hidden');
      showMessage(err && err.message ? err.message : 'Impossibile elaborare il file.');
    } finally { setBusy(false); $('fileInput').value = ''; }
  }

  function selectUsefulSheet(workbook) {
    const diagnostics = [];
    for (const name of workbook.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
      if (!rows.some(r => Array.isArray(r) && r.some(v => !isBlank(v)))) continue;
      const validation = validateSheet(rows, name);
      if (validation.ok) return { name, rows, validation };
      diagnostics.push(validation.message);
    }
    return null;
  }

  function validateSheet(rows, sheetName) {
    let headerRow = -1;
    for (let i = 0; i < Math.min(rows.length, 100); i++) {
      const r = rows[i] || [];
      const hits = Object.values(COL).filter(ix => !isBlank(r[ix])).length;
      if (r.length >= 18 && hits === 6) { headerRow = i; break; }
    }
    if (headerRow < 0) return { ok: false, message: `Foglio “${sheetName}” non valido: non trovo una riga intestazione con le colonne richieste A, F, H, O, P e R.` };
    const row = rows[headerRow] || [];
    const labels = [['A','Fornitore',COL.supplier],['F','Stato linea',COL.status],['H','Tipologia connettività',COL.tech],['O','Regione',COL.region],['P','Data invio attivazione',COL.sent],['R','Data attivazione/collaudo',COL.active]];
    const missing = labels.filter(x => isBlank(row[x[2]])).map(x => `${x[0]} (${x[1]})`);
    if (missing.length) return { ok: false, message: `Colonne mancanti o senza intestazione nel foglio “${sheetName}”: ${missing.join(', ')}.` };
    return { ok: true, headerRow };
  }

  function parseRows(rows, headerRow, analysisDate) {
    return rows.slice(headerRow + 1).filter(r => Array.isArray(r) && r.some(v => !isBlank(v))).map((r, idx) => {
      const status = normalizeStatus(r[COL.status]);
      const p = parseDate(r[COL.sent]);
      const rr = parseDate(r[COL.active]);
      const active = status === 'ATTIVO';
      let dateIssue = null;
      if (active) {
        if (!p) dateIssue = 'P mancante';
        else if (!rr) dateIssue = 'R mancante';
        else if (rr < p) dateIssue = 'R < P';
        else if (rr > analysisDate) dateIssue = 'R futura';
      }
      const validActivation = active && !dateIssue;
      const activationDays = validActivation ? diffDays(rr, p) : null;
      const backlog = status === 'ATTIVAZIONE';
      const validAging = backlog && p && p <= analysisDate;
      return {
        rowNumber: headerRow + idx + 2,
        supplier: cleanText(r[COL.supplier]) || 'Non indicato',
        supplierMissing: !cleanText(r[COL.supplier]),
        status,
        tech: cleanText(r[COL.tech]) || 'Non indicata',
        techMissing: !cleanText(r[COL.tech]),
        region: cleanText(r[COL.region]) || 'Non indicata',
        regionMissing: !cleanText(r[COL.region]),
        p, r: rr, active, backlog, ko: status === 'KO', core: CORE.has(status),
        dateIssue, validActivation, activationDays, validAging,
        agingDays: validAging ? diffDays(analysisDate, p) : null
      };
    });
  }

  function buildReport(rows, analysisDate, fileName, sheetName, headerRow) {
    const summary = aggregate(rows);
    const suppliers = groupAggregate(rows, r => r.supplier, 'supplier');
    const regions = groupAggregate(rows, r => r.region, 'region');
    const technologies = groupAggregate(rows, r => r.tech, 'tech');
    const cross = groupAggregate(rows, r => `${r.supplier}\u0000${r.region}`, 'cross').map(x => {
      const [supplier, region] = x.key.split('\u0000'); return { ...x, supplier, region };
    });
    const quality = {
      totalRows: rows.length,
      supplierMissing: rows.filter(r => r.supplierMissing).length,
      regionMissing: rows.filter(r => r.regionMissing).length,
      techMissing: rows.filter(r => r.techMissing).length,
      pMissing: rows.filter(r => !r.p).length,
      activeRMissing: rows.filter(r => r.active && !r.r).length,
      rBeforeP: rows.filter(r => r.p && r.r && r.r < r.p).length,
      rFuture: rows.filter(r => r.r && r.r > analysisDate).length,
      dateAnomalies: rows.filter(r => r.active && !r.validActivation).length
    };
    return { fileName, sheetName, headerRow, analysisDate, summary, suppliers, regions, technologies, cross, quality };
  }

  function aggregate(rows) {
    const core = rows.filter(r => r.core);
    const actives = rows.filter(r => r.active);
    const backlog = rows.filter(r => r.backlog);
    const ko = rows.filter(r => r.ko);
    const activationTimes = actives.filter(r => r.validActivation).map(r => r.activationDays);
    const aging = backlog.filter(r => r.validAging).map(r => r.agingDays);
    return {
      analyzed: core.length, active: actives.length, backlog: backlog.length, ko: ko.length,
      backlogRate: rate(backlog.length, core.length), koRate: rate(ko.length, core.length),
      validActivations: activationTimes.length, mean: mean(activationTimes), median: quantile(activationTimes,.5),
      p75: quantile(activationTimes,.75), p90: quantile(activationTimes,.9), min: extrema(activationTimes,'min'), max: extrema(activationTimes,'max'),
      agingMean: mean(aging), agingMedian: quantile(aging,.5), agingValid: aging.length,
      agingBuckets: {
        '0–15 giorni': aging.filter(x => x <= 15).length,
        '16–30 giorni': aging.filter(x => x >= 16 && x <= 30).length,
        '31–60 giorni': aging.filter(x => x >= 31 && x <= 60).length,
        '61–90 giorni': aging.filter(x => x >= 61 && x <= 90).length,
        'Oltre 90 giorni': aging.filter(x => x > 90).length
      },
      backlog60: aging.filter(x => x > 60).length,
      backlog90: aging.filter(x => x > 90).length,
      dateAnomalies: actives.filter(r => !r.validActivation).length
    };
  }

  function groupAggregate(rows, keyFn, type) {
    const groups = new Map();
    rows.forEach(r => { const key = keyFn(r); if (!groups.has(key)) groups.set(key, []); groups.get(key).push(r); });
    return [...groups].map(([key, list]) => ({ key, type, ...aggregate(list) })).sort((a,b) => b.analyzed - a.analyzed || String(a.key).localeCompare(String(b.key), 'it'));
  }

  function renderDashboard() {
    $('fileName').textContent = report.fileName;
    $('analysisMeta').textContent = `Data analisi ${df.format(report.analysisDate)} · Foglio “${report.sheetName}” · ${nf.format(report.quality.totalRows)} righe`;
    $('dashboard').classList.remove('hidden');
    renderKpis(); renderAlerts(); renderAging(); renderQuality(); renderAllTables(); updateHistoryButtons();
    try {
      renderCharts();
    } catch (chartError) {
      console.error('Errore grafici:', chartError);
      showMessage('I dati sono stati elaborati correttamente, ma uno o più grafici non sono disponibili. KPI e tabelle restano consultabili.');
    }
    $('dashboard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderKpis() {
    const s = report.summary;
    $('primaryKpis').innerHTML = [
      kpi('Pratiche analizzate', fmtInt(s.analyzed), 'Attivo + Attivazione + KO', ''),
      kpi('Attive', fmtInt(s.active), `${fmtInt(s.validActivations)} con date valide`, 'positive'),
      kpi('Backlog', fmtInt(s.backlog), pct(s.backlogRate), 'attention'),
      kpi('KO', fmtInt(s.ko), pct(s.koRate), 'critical')
    ].join('');
    $('secondaryKpis').innerHTML = [
      kpi('Tempo medio', days(s.mean), '', ''), kpi('Mediana', days(s.median), '', ''), kpi('P90', days(s.p90), '', ''),
      kpi('Aging backlog mediano', days(s.agingMedian), `${fmtInt(s.agingValid)} con data valida`, ''),
      kpi('Backlog >60 giorni', fmtInt(s.backlog60), '', 'attention'), kpi('Backlog >90 giorni', fmtInt(s.backlog90), '', 'critical')
    ].join('');
  }

  function kpi(label, value, sub, cls) {
    const isPct = sub && /%$/.test(sub);
    return `<article class="kpi-card ${cls}"><div class="kpi-label">${esc(label)}</div><div class="kpi-value">${esc(value)}${isPct ? `<span class="kpi-sub">${esc(sub)}</span>` : ''}</div>${sub && !isPct ? `<span class="kpi-note">${esc(sub)}</span>` : ''}</article>`;
  }

  function renderAlerts() {
    const min = Math.max(1, Number($('minSample').value) || 20);
    $('sampleChip').textContent = `Campione minimo: ${min} pratiche`;
    const eligibleS = report.suppliers.filter(x => x.analyzed >= min);
    const eligibleR = report.regions.filter(x => x.analyzed >= min);
    const eligibleT = report.technologies.filter(x => x.analyzed >= min);
    const specs = [
      ['fornitore con KO rate più elevato', top(eligibleS,'koRate'), 'koRate', 'ko'],
      ['fornitore con backlog rate più elevato', top(eligibleS,'backlogRate'), 'backlogRate', 'backlog'],
      ['fornitore con P90 più elevato', top(eligibleS.filter(x => x.p90 != null),'p90'), 'p90', 'validActivations'],
      ['regione con KO rate più elevato', top(eligibleR,'koRate'), 'koRate', 'ko'],
      ['regione con backlog rate più elevato', top(eligibleR,'backlogRate'), 'backlogRate', 'backlog'],
      ['tecnologia con KO rate più elevato', top(eligibleT,'koRate'), 'koRate', 'ko']
    ];
    $('alerts').innerHTML = specs.map(([label,item,metric,countKey]) => {
      if (!item) return `<div class="alert-item neutral">Nessun ${esc(label)}: non ci sono gruppi con almeno ${min} pratiche analizzate.</div>`;
      const name = item.key;
      const detail = metric === 'p90' ? `${days(item.p90)} su ${fmtInt(item.validActivations)} attivazioni valide` : `${pct(item[metric])} (${fmtInt(item[countKey])} casi) su ${fmtInt(item.analyzed)} pratiche analizzate`;
      return `<div class="alert-item"><strong>${esc(name)}</strong> è il ${esc(label)}: ${esc(detail)}.</div>`;
    }).join('');
  }

  function renderAging() {
    $('agingMeta').textContent = `Aging medio: ${days(report.summary.agingMean)} · ${fmtInt(report.summary.agingValid)} date valide`;
    $('agingBuckets').innerHTML = Object.entries(report.summary.agingBuckets).map(([k,v]) => `<div class="aging-bucket"><span>${esc(k)}</span><strong>${fmtInt(v)}</strong></div>`).join('');
    const s = report.summary;
    $('activationDetails').innerHTML = [['Date valide',fmtInt(s.validActivations)],['P75',days(s.p75)],['Minimo',days(s.min)],['Massimo',days(s.max)]].map(([k,v]) => `<div class="aging-bucket"><span>${esc(k)}</span><strong>${esc(v)}</strong></div>`).join('');
  }

  function renderQuality() {
    const q = report.quality;
    const items = [['Totale righe Excel',q.totalRows],['Fornitore mancante',q.supplierMissing],['Regione mancante',q.regionMissing],['Tecnologia mancante',q.techMissing],['P mancante',q.pMissing],['R mancante sulle Attive',q.activeRMissing],['R precedente a P',q.rBeforeP],['R futura',q.rFuture],['Anomalie escluse dai tempi',q.dateAnomalies]];
    $('qualityGrid').innerHTML = items.map(([label,value],i) => `<div class="quality-item ${i>0 && value>0?'bad':''}"><span>${esc(label)}</span><strong>${fmtInt(value)}</strong></div>`).join('');
  }

  const tableDefs = {
    supplier: {
      element:'supplierTable', rows:() => report.suppliers,
      cols:[['key','Fornitore','text'],['analyzed','Pratiche analizzate','int'],['active','Attive','int'],['backlog','Backlog','int'],['backlogRate','Backlog %','pct'],['ko','KO','int'],['koRate','KO %','pct'],['validActivations','Attivazioni date valide','int'],['mean','Tempo medio','days'],['median','Mediana','days'],['p90','P90','days'],['agingMedian','Aging backlog mediano','days'],['backlog60','Backlog >60','int'],['backlog90','Backlog >90','int']]
    },
    region: {
      element:'regionTable', rows:() => report.regions,
      cols:[['key','Regione','text'],['analyzed','Pratiche analizzate','int'],['active','Attive','int'],['backlog','Backlog','int'],['backlogRate','Backlog %','pct'],['ko','KO','int'],['koRate','KO %','pct'],['mean','Tempo medio','days'],['median','Mediana','days'],['p90','P90','days'],['agingMedian','Aging backlog mediano','days']]
    },
    tech: {
      element:'techTable', rows:() => report.technologies,
      cols:[['key','Tecnologia','text'],['analyzed','Pratiche analizzate','int'],['active','Attive','int'],['backlog','Backlog','int'],['backlogRate','Backlog %','pct'],['ko','KO','int'],['koRate','KO %','pct'],['mean','Tempo medio','days'],['median','Mediana','days'],['p90','P90','days']]
    },
    cross: {
      element:'crossTable', rows:() => report.cross,
      cols:[['supplier','Fornitore','text'],['region','Regione','text'],['analyzed','Pratiche analizzate','int'],['backlog','Backlog','int'],['backlogRate','Backlog %','pct'],['ko','KO','int'],['koRate','KO %','pct'],['mean','Tempo medio','days'],['median','Mediana','days'],['p90','P90','days'],['agingMedian','Aging backlog mediano','days']]
    }
  };

  function renderAllTables() { Object.keys(tableDefs).forEach(renderTable); }
  function renderTable(name) {
    const def = tableDefs[name]; const table = $(def.element); const state = sortState[name] || { key:'analyzed', dir:-1 }; sortState[name] = state;
    const rows = [...def.rows()].sort((a,b) => compare(a[state.key],b[state.key],state.dir));
    table.innerHTML = `<thead><tr>${def.cols.map(([key,label,type]) => `<th class="${type==='text'?'':'numeric'}" data-key="${key}">${esc(label)}${state.key===key?(state.dir===1?' ↑':' ↓'):''}</th>`).join('')}</tr></thead><tbody>${rows.map(r => tableRow(r,def.cols,name)).join('')}</tbody>`;
    table.querySelectorAll('th').forEach(th => th.addEventListener('click', () => { const key=th.dataset.key; state.dir=state.key===key?-state.dir:-1; state.key=key; renderTable(name); }));
  }

  function tableRow(row, cols, tableName) {
    const min = Math.max(1, Number($('minSample').value) || 20);
    const critical = row.analyzed >= min && (row.koRate >= .2 || row.backlogRate >= .25);
    const warning = row.analyzed >= min && !critical && (row.koRate >= .1 || row.backlogRate >= .15);
    const cls = tableName==='region' || tableName==='cross' ? (critical?'row-critical':warning?'row-warning':'') : '';
    return `<tr class="${cls}">${cols.map(([key,,type]) => `<td class="${type==='text'?'':'numeric'}">${formatCell(row[key],type,key,row)}</td>`).join('')}</tr>`;
  }

  function formatCell(v,type,key,row) {
    if (type==='text') return esc(v == null ? '—' : v);
    if (type==='int') return fmtInt(v);
    if (type==='days') return days(v);
    if (type==='pct') {
      const cls = key==='koRate' ? (v>=.2?'red':v>=.1?'amber':'green') : (v>=.25?'red':v>=.15?'amber':'green');
      return `<span class="badge ${cls}">${pct(v)}</span>`;
    }
    return esc(v);
  }

  function renderCharts() {
    if (typeof Chart === 'undefined') return;
    Object.values(charts).forEach(c => c.destroy()); charts = {};
    const rank = (items,key,limit=12) => [...items].filter(x=>x.analyzed>0).sort((a,b)=>(b[key]??-1)-(a[key]??-1)).slice(0,limit).reverse();
    charts.sko = horizontalBar('chartSupplierKo',rank(report.suppliers,'koRate'),x=>x.key,[['KO %',x=>x.koRate*100,'#b42318']],true);
    charts.sback = horizontalBar('chartSupplierBacklog',rank(report.suppliers,'backlogRate'),x=>x.key,[['Backlog %',x=>x.backlogRate*100,'#d97706']],true);
    const stime = [...report.suppliers].filter(x=>x.validActivations>0).sort((a,b)=>(b.p90??-1)-(a.p90??-1)).slice(0,12).reverse();
    charts.stime = horizontalBar('chartSupplierTime',stime,x=>x.key,[['Tempo medio',x=>x.mean,'#2563a7'],['P90',x=>x.p90,'#0b1f3a']],false);
    charts.rko = horizontalBar('chartRegionKo',rank(report.regions,'koRate'),x=>x.key,[['KO %',x=>x.koRate*100,'#b42318']],true);
    charts.rback = horizontalBar('chartRegionBacklog',rank(report.regions,'backlogRate'),x=>x.key,[['Backlog %',x=>x.backlogRate*100,'#d97706']],true);
    const tech = [...report.technologies].sort((a,b)=>b.analyzed-a.analyzed).slice(0,10);
    charts.tech = new Chart($('chartTech'),{type:'doughnut',data:{labels:tech.map(x=>x.key),datasets:[{data:tech.map(x=>x.analyzed),backgroundColor:['#0b1f3a','#2563a7','#6287ad','#9db2c7','#147d55','#d97706','#7559a6','#506577','#b42318','#b2bfca'],borderWidth:2,borderColor:'#fff'}]},options:chartOptions(false,true)});
    renderHistoryChart();
  }

  function horizontalBar(id,items,labelFn,series,percent) {
    return new Chart($(id),{type:'bar',data:{labels:items.map(labelFn),datasets:series.map(([label,get,color])=>({label,data:items.map(get),backgroundColor:color,borderRadius:3,barThickness:14}))},options:{...chartOptions(true,false),indexAxis:'y',scales:{x:{beginAtZero:true,grid:{color:'#edf0f3'},ticks:{callback:v=>percent?`${v}%`:v}},y:{grid:{display:false},ticks:{color:'#526174'}}}}});
  }

  function renderHistoryChart() {
    const history = loadHistory().filter(x=>!report || Number(x.year)===report.analysisDate.getFullYear()).sort((a,b)=>a.analysisDate.localeCompare(b.analysisDate));
    if (charts.history) charts.history.destroy();
    charts.history = new Chart($('chartHistory'),{type:'line',data:{labels:history.map(x=>x.month),datasets:[
      {label:'KO %',data:history.map(x=>x.koRate*100),borderColor:'#b42318',backgroundColor:'#b42318',yAxisID:'yPct'},
      {label:'Backlog %',data:history.map(x=>x.backlogRate*100),borderColor:'#d97706',backgroundColor:'#d97706',yAxisID:'yPct'},
      {label:'Tempo medio',data:history.map(x=>x.mean),borderColor:'#2563a7',backgroundColor:'#2563a7',yAxisID:'yDays'},
      {label:'P90',data:history.map(x=>x.p90),borderColor:'#0b1f3a',backgroundColor:'#0b1f3a',yAxisID:'yDays'}
    ].map(x=>({...x,tension:.25,pointRadius:3,spanGaps:true}))},options:{...chartOptions(false,false),scales:{yPct:{beginAtZero:true,position:'left',grid:{color:'#edf0f3'},ticks:{callback:v=>`${v}%`}},yDays:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},ticks:{callback:v=>`${v} gg`}},x:{grid:{display:false}}}}});
  }

  function chartOptions(indexAxis,legend) { return {responsive:true,maintainAspectRatio:false,plugins:{legend:{display:legend!==false,position:'bottom',labels:{boxWidth:10,usePointStyle:true,font:{size:11}}},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label||ctx.label}: ${Number(ctx.raw).toLocaleString('it-IT',{maximumFractionDigits:1})}`}}},animation:{duration:300}}; }

  function saveSnapshot(summary, date) {
    const history = loadHistory();
    const month = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}`;
    const snap = { month, year:date.getFullYear(), analysisDate:isoDate(date), analyzed:summary.analyzed, active:summary.active, backlog:summary.backlog, backlogRate:summary.backlogRate, ko:summary.ko, koRate:summary.koRate, mean:summary.mean, median:summary.median, p90:summary.p90, backlog60:summary.backlog60, backlog90:summary.backlog90 };
    const ix = history.findIndex(x=>x.month===month); if(ix>=0) history[ix]=snap; else history.push(snap);
    localStorage.setItem(STORAGE_KEY,JSON.stringify(history));
  }
  function loadHistory(){ try { const x=JSON.parse(localStorage.getItem(STORAGE_KEY)||'[]'); return Array.isArray(x)?x.filter(validSnapshot):[]; } catch { return []; } }
  function validSnapshot(x){return x&&/^\d{4}-\d{2}$/.test(String(x.month))&&/^\d{4}-\d{2}-\d{2}$/.test(String(x.analysisDate))&&Number.isFinite(Number(x.analyzed));}
  function exportHistory(){downloadBlob(JSON.stringify({schema:'connectivity-dashboard-history-v1',exportedAt:new Date().toISOString(),snapshots:loadHistory()},null,2),'storico_connettivita.json','application/json');}
  async function importHistory(file){
    if(!file)return; hideMessage();
    try{const parsed=JSON.parse(await file.text());const list=Array.isArray(parsed)?parsed:parsed.snapshots;if(!Array.isArray(list)||!list.every(validSnapshot))throw new Error('Il file JSON non contiene uno storico valido.');
      const merged=new Map(loadHistory().map(x=>[x.month,x]));list.forEach(x=>merged.set(x.month,x));localStorage.setItem(STORAGE_KEY,JSON.stringify([...merged.values()]));updateHistoryButtons();if(report)renderHistoryChart();
    }catch(e){showMessage(e.message||'Importazione storico non riuscita.');}finally{$('historyInput').value='';}
  }
  function clearHistory(){if(!loadHistory().length)return;if(confirm('Cancellare definitivamente tutti gli snapshot mensili salvati in questo browser?')){localStorage.removeItem(STORAGE_KEY);updateHistoryButtons();if(report)renderHistoryChart();}}
  function updateHistoryButtons(){$('exportHistoryBtn').disabled=loadHistory().length===0;}

  function exportExcel() {
    if(!report||typeof XLSX==='undefined')return;
    const wb=XLSX.utils.book_new();
    const s=report.summary;
    appendSheet(wb,'Riepilogo',[['Indicatore','Valore'],['Data analisi',df.format(report.analysisDate)],['Pratiche analizzate',s.analyzed],['Attive',s.active],['Backlog',s.backlog],['Backlog %',s.backlogRate],['KO',s.ko],['KO %',s.koRate],['Attivazioni con date valide',s.validActivations],['Tempo medio',s.mean],['Mediana',s.median],['P75',s.p75],['P90',s.p90],['Minimo',s.min],['Massimo',s.max],['Aging medio',s.agingMean],['Aging mediano',s.agingMedian],['Backlog >60',s.backlog60],['Backlog >90',s.backlog90],['Anomalie date',s.dateAnomalies]]);
    appendDataSheet(wb,'Fornitori',report.suppliers,tableDefs.supplier.cols);
    appendDataSheet(wb,'Regioni',report.regions,tableDefs.region.cols);
    appendDataSheet(wb,'Tecnologie',report.technologies,tableDefs.tech.cols);
    appendDataSheet(wb,'Fornitore_Regione',report.cross,tableDefs.cross.cols);
    const histCols=[['month','Mese'],['year','Anno'],['analysisDate','Data analisi'],['analyzed','Pratiche analizzate'],['active','Attive'],['backlog','Backlog'],['backlogRate','Backlog %'],['ko','KO'],['koRate','KO %'],['mean','Media'],['median','Mediana'],['p90','P90'],['backlog60','Backlog >60'],['backlog90','Backlog >90']];
    appendDataSheet(wb,'Storico',loadHistory().sort((a,b)=>a.month.localeCompare(b.month)),histCols);
    appendSheet(wb,'Legenda',[['Voce','Definizione'],['Pratiche analizzate','Attivo + Attivazione + KO'],['Backlog','Esclusivamente Stato = Attivazione'],['Backlog %','Backlog / Pratiche analizzate'],['KO %','KO / Pratiche analizzate'],['Tempo attivazione','Data R - Data P, giorni calendario'],['Aging backlog','Data analisi - Data P'],['P90','Giorni entro cui si conclude il 90% delle attivazioni valide'],['Anomalie date','P o R mancanti sulle attive, R < P, R futura']]);
    XLSX.writeFile(wb,`report_connettivita_${isoDate(report.analysisDate).replaceAll('-','')}.xlsx`);
  }
  function appendDataSheet(wb,name,rows,cols){appendSheet(wb,name,[cols.map(c=>c[1]),...rows.map(r=>cols.map(c=>r[c[0]]??null))],cols);}
  function appendSheet(wb,name,aoa,cols){const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!freeze']={xSplit:0,ySplit:1};ws['!autofilter']={ref:ws['!ref']};ws['!cols']=(cols||aoa[0]).map((c,i)=>({wch:Math.min(34,Math.max(12,...aoa.map(r=>String(r[i]??'').length+2)))}));XLSX.utils.book_append_sheet(wb,ws,name);}

  function activateTab(tab){document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x.dataset.tab===tab));document.querySelectorAll('.tab-panel').forEach(x=>x.classList.toggle('active',x.id===`tab-${tab}`));}
  function analysisDateFromFilename(name){const m=name.match(/^(\d{4})(\d{2})(\d{2})/);if(m){const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12);if(d.getFullYear()===Number(m[1])&&d.getMonth()===Number(m[2])-1&&d.getDate()===Number(m[3]))return d;}const n=new Date();return new Date(n.getFullYear(),n.getMonth(),n.getDate(),12);}
  function parseDate(v){if(v instanceof Date&&!isNaN(v))return dayDate(v.getFullYear(),v.getMonth(),v.getDate());if(typeof v==='number'&&Number.isFinite(v)){const p=XLSX.SSF.parse_date_code(v);return p?dayDate(p.y,p.m-1,p.d):null;}if(typeof v!=='string'||!v.trim())return null;const s=v.trim();let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\s|$)/);if(m){let y=Number(m[3]);if(y<100)y+=2000;return validDate(y,Number(m[2])-1,Number(m[1]));}m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:T|\s|$)/);if(m)return validDate(Number(m[1]),Number(m[2])-1,Number(m[3]));const d=new Date(s);return isNaN(d)?null:dayDate(d.getFullYear(),d.getMonth(),d.getDate());}
  function validDate(y,m,d){const x=dayDate(y,m,d);return x.getFullYear()===y&&x.getMonth()===m&&x.getDate()===d?x:null;}function dayDate(y,m,d){return new Date(y,m,d,12);}
  function diffDays(a,b){return Math.round((Date.UTC(a.getFullYear(),a.getMonth(),a.getDate())-Date.UTC(b.getFullYear(),b.getMonth(),b.getDate()))/DAY);}
  function cleanText(v){return v==null?'':String(v).replace(/\s+/g,' ').trim();}function normalizeStatus(v){return cleanText(v).toLocaleUpperCase('it-IT');}function isBlank(v){return v==null||String(v).trim()==='';}
  function mean(a){return a.length?a.reduce((x,y)=>x+y,0)/a.length:null;}function quantile(a,p){if(!a.length)return null;const s=[...a].sort((x,y)=>x-y);return s[Math.max(0,Math.ceil(p*s.length)-1)];}function extrema(a,t){return a.length?Math[t](...a):null;}function rate(n,d){return d?n/d:0;}function top(a,k){return a.length?[...a].sort((x,y)=>(y[k]??-Infinity)-(x[k]??-Infinity))[0]:null;}
  function compare(a,b,dir){if(a==null&&b==null)return 0;if(a==null)return 1;if(b==null)return -1;if(typeof a==='string')return a.localeCompare(b,'it')*dir;return (a-b)*dir;}
  function fmtInt(v){return v==null?'—':nf.format(v);}function pct(v){return v==null?'—':`${(v*100).toLocaleString('it-IT',{minimumFractionDigits:1,maximumFractionDigits:1})}%`;}function days(v){return v==null?'—':`${Number(v).toLocaleString('it-IT',{maximumFractionDigits:1})} gg`;}
  function isoDate(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}function esc(v){return String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));}
  function downloadBlob(content,name,type){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
  function setBusy(on){$('fileInput').disabled=on;$('dropZone').querySelector('strong').textContent=on?'Elaborazione in corso…':'Seleziona il file';}
  function showMessage(text){$('message').textContent=text;$('message').classList.remove('hidden');}
  function hideMessage(){$('message').classList.add('hidden');$('message').textContent='';}
})();
