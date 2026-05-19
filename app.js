/* ═══════════════════════════════════════
   app.js — Main application logic
   ═══════════════════════════════════════ */

const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

const state = {
  workbook: null, sheetNames: [], headers: [], rows: [],
  uniqueCache: {},   // colIndex → string[]
  numericCols: {},    // colIndex → boolean
  charts: [],        // dashboard chart entries
  wiz: { step:1, primaryCol:null, primaryFilter:[], measureCol:null, measureFilter:[], aggregation:'count', chartTypes:[] }
};

let chartIdCounter = 0;

/* ═══════════ INIT ═══════════ */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  setupUpload();
  setupWizardNav();
  setupAggButtons();
  $('#wizard-close').onclick = closeWizard;
  $('#fullscreen-close').onclick = () => $('#fullscreen-overlay').classList.add('hidden');
  $('#fullscreen-overlay').addEventListener('click', e => { if(e.target.id==='fullscreen-overlay') $('#fullscreen-overlay').classList.add('hidden'); });
  $('#nav-new-file').onclick = resetAll;
  $('#nav-add-viz').onclick = openWizard;
  $('#create-viz-btn').onclick = openWizard;
  $('#dashboard-add-btn').onclick = openWizard;
});

/* ═══════════ THEME TOGGLE ═══════════ */
function initTheme() {
  const saved = localStorage.getItem('dataviz-theme') || 'dark';
  setTheme(saved);
  $('#theme-toggle').onclick = toggleTheme;
  $('#upload-theme-toggle').onclick = toggleTheme;
}

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || 'dark';
  setTheme(current === 'dark' ? 'light' : 'dark');
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('dataviz-theme', theme);
  const icon = theme === 'dark' ? '🌙' : '☀️';
  const toggles = document.querySelectorAll('.theme-toggle');
  toggles.forEach(t => t.textContent = icon);
  // Re-render all existing charts with new theme colors
  rebuildChartsForTheme();
}

function rebuildChartsForTheme() {
  if (!state.headers.length) return;
  state.charts.forEach(entry => {
    const card = $(`#chart-card-${entry.id}`);
    if (!card) return;
    const body = $(`#chart-body-${entry.id}`);
    if (!body) return;
    // Destroy old chart
    if (card._chartInstance) { card._chartInstance.destroy(); card._chartInstance = null; }
    // Re-aggregate and re-render
    const aggResult = aggregate(entry.config);
    if (entry.type === 'heatmap') {
      body.innerHTML = buildHeatmapHTML(aggResult);
    } else {
      body.innerHTML = '';
      const canvas = document.createElement('canvas');
      canvas.id = 'canvas-' + entry.id;
      body.appendChild(canvas);
      card._chartInstance = buildChart(canvas, entry.type, aggResult, entry.config, state.headers);
    }
  });
}

/* ═══════════ FILE UPLOAD ═══════════ */
function setupUpload() {
  const zone = $('#upload-zone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('drag-over'); if(e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  $('#file-input').addEventListener('change', e => { if(e.target.files[0]) handleFile(e.target.files[0]); });
}

function handleFile(file) {
  showLoading('Reading ' + file.name + '...');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      $('#loading-details').textContent = 'Parsing spreadsheet...';
      setTimeout(() => {
        const data = new Uint8Array(e.target.result);
        state.workbook = XLSX.read(data, { type:'array' });
        state.sheetNames = state.workbook.SheetNames;
        hideLoading();
        if (state.sheetNames.length > 1) showSheetSelector(file.name);
        else loadSheet(state.sheetNames[0], file.name);
      }, 50);
    } catch(err) { hideLoading(); alert('Error parsing file: ' + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

/* ═══════════ LOADING ═══════════ */
function showLoading(msg) { $('#loading-details').textContent = msg||''; $('#loading-overlay').classList.remove('hidden'); }
function hideLoading() { $('#loading-overlay').classList.add('hidden'); }

/* ═══════════ SHEET SELECTOR ═══════════ */
function showSheetSelector(fileName) {
  showSection('sheet-section');
  showNav(fileName, '');
  const list = $('#sheet-list');
  list.innerHTML = '';
  state.sheetNames.forEach(name => {
    const btn = document.createElement('button');
    btn.className = 'sheet-btn';
    btn.innerHTML = `<span class="sheet-icon">📄</span> ${esc(name)}`;
    btn.onclick = () => loadSheet(name, fileName);
    list.appendChild(btn);
  });
}

/* ═══════════ LOAD SHEET ═══════════ */
function loadSheet(sheetName, fileName) {
  showLoading('Loading sheet: ' + sheetName);
  setTimeout(() => {
    const ws = state.workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
    if (!raw.length) { hideLoading(); alert('Sheet is empty'); return; }
    state.headers = raw[0].map(h => h != null ? String(h).trim() : '');
    state.rows = raw.slice(1);
    state.uniqueCache = {};
    state.numericCols = {};
    // Pre-detect numeric columns
    state.headers.forEach((_, i) => { state.numericCols[i] = detectNumeric(i); });
    hideLoading();
    showNav(fileName, sheetName);
    showPreview(sheetName);
  }, 50);
}

function detectNumeric(colIdx) {
  let numCount = 0, total = 0;
  const limit = Math.min(state.rows.length, 200);
  for (let i = 0; i < limit; i++) {
    const v = state.rows[i]?.[colIdx];
    if (v === '' || v == null) continue;
    total++;
    if (!isNaN(Number(v))) numCount++;
  }
  return total > 0 && (numCount / total) > 0.7;
}

function getUnique(colIdx) {
  if (state.uniqueCache[colIdx]) return state.uniqueCache[colIdx];
  const set = new Set();
  state.rows.forEach(r => {
    const v = r[colIdx];
    if (v !== '' && v != null) set.add(String(v).trim());
  });
  state.uniqueCache[colIdx] = [...set].sort();
  return state.uniqueCache[colIdx];
}

/* ═══════════ NAVIGATION ═══════════ */
function showNav(fileName, sheetName) {
  $('#top-nav').classList.remove('hidden');
  $('#nav-file-name').textContent = fileName || '';
  $('#nav-sheet-name').textContent = sheetName || '';
}

function showSection(id) {
  ['upload-section','sheet-section','preview-section','dashboard-section'].forEach(s => {
    $('#'+s).classList.toggle('hidden', s !== id);
    if(s === id) $('#'+s).classList.add('active-section');
  });
}

/* ═══════════ PREVIEW ═══════════ */
function showPreview(sheetName) {
  showSection('preview-section');
  $('#nav-add-viz').classList.remove('hidden');
  $('#stat-rows').textContent = state.rows.length.toLocaleString();
  $('#stat-cols').textContent = state.headers.length;
  $('#stat-sheet').textContent = sheetName;

  // Header chips
  const chips = $('#header-chips');
  chips.innerHTML = '';
  state.headers.forEach(h => {
    if (!h) return;
    const chip = document.createElement('span');
    chip.className = 'header-chip';
    chip.textContent = h;
    chips.appendChild(chip);
  });

  // Preview table
  const table = $('#preview-table');
  let html = '<thead><tr>';
  state.headers.forEach(h => { html += `<th>${esc(h)}</th>`; });
  html += '</tr></thead><tbody>';
  const previewRows = state.rows.slice(0, 10);
  previewRows.forEach(r => {
    html += '<tr>';
    state.headers.forEach((_, i) => { html += `<td>${esc(String(r[i] ?? ''))}</td>`; });
    html += '</tr>';
  });
  html += '</tbody>';
  table.innerHTML = html;
}

/* ═══════════ WIZARD ═══════════ */
function openWizard() {
  state.wiz = { step:1, primaryCol:null, primaryFilter:[], measureCol:null, measureFilter:[], aggregation:'count', chartTypes:[] };
  updateWizardStep(1);
  populateStep1();
  $('#wizard-overlay').classList.remove('hidden');
}

function closeWizard() { $('#wizard-overlay').classList.add('hidden'); }

function setupWizardNav() {
  $('#wizard-back').onclick = () => {
    let prev = state.wiz.step - 1;
    if (prev === 4 && (state.wiz.measureCol === null)) prev = 3;
    if (prev === 3 && state.wiz.step === 4) prev = 3;
    if (prev < 1) prev = 1;
    updateWizardStep(prev);
  };
  $('#wizard-next').onclick = handleNext;
  $('#wizard-overlay').addEventListener('click', e => { if(e.target.id==='wizard-overlay') closeWizard(); });
}

function handleNext() {
  const s = state.wiz.step;
  if (s === 1) {
    if (state.wiz.primaryCol === null) return shake($('#wizard-step-1'));
    collectPrimaryFilter_init();
    updateWizardStep(2);
  } else if (s === 2) {
    collectTags('primary');
    populateStep3();
    updateWizardStep(3);
  } else if (s === 3) {
    if (state.wiz.measureCol === null && !$('#single-col-btn').classList.contains('selected')) return shake($('#wizard-step-3'));
    if ($('#single-col-btn').classList.contains('selected')) {
      state.wiz.measureCol = null;
      populateStep5();
      updateWizardStep(5);
    } else {
      collectMeasureFilter_init();
      updateWizardStep(4);
    }
  } else if (s === 4) {
    collectTags('measure');
    populateStep5();
    updateWizardStep(5);
  } else if (s === 5) {
    if (!state.wiz.chartTypes.length) return shake($('#wizard-step-5'));
    generateAllCharts();
    closeWizard();
  }
}

function updateWizardStep(n) {
  state.wiz.step = n;
  $$('.wizard-step').forEach(el => el.classList.toggle('active', el.id === 'wizard-step-'+n));
  $$('.progress-step').forEach(el => {
    const sn = +el.dataset.step;
    el.classList.toggle('active', sn === n);
    el.classList.toggle('done', sn < n);
  });
  // Fill progress lines
  $$('.progress-line').forEach((line, i) => {
    const fill = line.querySelector('.progress-fill');
    if(fill) fill.style.width = (i < n-1) ? '100%' : '0%';
  });
  $('#wizard-back').disabled = n === 1;
  $('#wizard-next').textContent = n === 5 ? '✨ Generate' : 'Next →';
}

/* ── Step 1: Primary column ── */
function populateStep1() {
  const grid = $('#primary-col-list');
  grid.innerHTML = '';
  state.headers.forEach((h, i) => {
    if (!h) return;
    const btn = document.createElement('button');
    btn.className = 'col-option';
    btn.innerHTML = `${esc(h)}<span class="col-type">${state.numericCols[i] ? '123 Numeric' : 'Abc Text'}</span>`;
    btn.onclick = () => {
      $$('#primary-col-list .col-option').forEach(b => b.classList.remove('selected'));
      btn.classList.add('selected');
      state.wiz.primaryCol = i;
    };
    grid.appendChild(btn);
  });
}

/* ── Step 2: Primary filter ── */
function collectPrimaryFilter_init() {
  const col = state.wiz.primaryCol;
  $('#primary-col-name').textContent = state.headers[col];
  clearTagInput('primary');
  setupTagInput('primary', col);
}

/* ── Step 3: Measure column ── */
function populateStep3() {
  $('#primary-col-name-2').textContent = state.headers[state.wiz.primaryCol];
  const grid = $('#measure-col-list');
  grid.innerHTML = '';
  const singleBtn = $('#single-col-btn');
  singleBtn.classList.remove('selected');
  singleBtn.onclick = () => {
    singleBtn.classList.toggle('selected');
    if (singleBtn.classList.contains('selected')) {
      $$('#measure-col-list .col-option').forEach(b => b.classList.remove('selected'));
      state.wiz.measureCol = null;
    }
  };
  state.headers.forEach((h, i) => {
    if (!h || i === state.wiz.primaryCol) return;
    const btn = document.createElement('button');
    btn.className = 'col-option';
    btn.innerHTML = `${esc(h)}<span class="col-type">${state.numericCols[i] ? '123 Numeric' : 'Abc Text'}</span>`;
    btn.onclick = () => {
      $$('#measure-col-list .col-option').forEach(b => b.classList.remove('selected'));
      singleBtn.classList.remove('selected');
      btn.classList.add('selected');
      state.wiz.measureCol = i;
    };
    grid.appendChild(btn);
  });
}

/* ── Step 4: Measure filter ── */
function collectMeasureFilter_init() {
  const col = state.wiz.measureCol;
  $('#measure-col-name').textContent = state.headers[col];
  clearTagInput('measure');
  setupTagInput('measure', col);
  const aggSec = $('#aggregation-section');
  if (state.numericCols[col]) {
    aggSec.classList.remove('hidden');
    state.wiz.aggregation = 'sum';
    $$('.agg-btn').forEach(b => b.classList.toggle('active', b.dataset.agg === 'sum'));
  } else {
    aggSec.classList.add('hidden');
    state.wiz.aggregation = 'count';
  }
}

function setupAggButtons() {
  $('#agg-options').addEventListener('click', e => {
    const btn = e.target.closest('.agg-btn');
    if(!btn) return;
    $$('.agg-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    state.wiz.aggregation = btn.dataset.agg;
  });
}

/* ── Step 5: Chart types ── */
function populateStep5() {
  const grid = $('#chart-type-grid');
  grid.innerHTML = '';
  state.wiz.chartTypes = [];
  const isSingle = state.wiz.measureCol === null;
  Object.entries(CHART_TYPES).forEach(([key, ct]) => {
    if (ct.mode === 'cross' && isSingle) return;
    if (ct.mode === 'single' && !isSingle) return;
    const btn = document.createElement('button');
    btn.className = 'chart-type-option';
    btn.innerHTML = `<span class="ct-icon">${ct.icon}</span><span class="ct-name">${ct.name}</span><span class="ct-desc">${ct.desc}</span>`;
    btn.onclick = () => {
      btn.classList.toggle('selected');
      if (btn.classList.contains('selected')) state.wiz.chartTypes.push(key);
      else state.wiz.chartTypes = state.wiz.chartTypes.filter(t => t !== key);
    };
    grid.appendChild(btn);
  });
}

/* ═══════════ TAG INPUT SYSTEM ═══════════ */
function setupTagInput(prefix, colIdx) {
  const input = $(`#${prefix}-filter-input`);
  const dropdown = $(`#${prefix}-autocomplete`);
  const addAllBtn = $(`#${prefix}-add-all`);
  const clearAllBtn = $(`#${prefix}-clear-all`);
  const unique = getUnique(colIdx);

  const getTags = () => prefix === 'primary' ? state.wiz.primaryFilter : state.wiz.measureFilter;
  const setTags = v => { if(prefix==='primary') state.wiz.primaryFilter=v; else state.wiz.measureFilter=v; };

  const renderTags = () => {
    const area = $(`#${prefix}-tags`);
    area.innerHTML = '';
    getTags().forEach(t => {
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.innerHTML = `${esc(t)}<button class="tag-remove" data-val="${esc(t)}">&times;</button>`;
      area.appendChild(tag);
    });
    area.querySelectorAll('.tag-remove').forEach(btn => {
      btn.onclick = () => { setTags(getTags().filter(t => t !== btn.dataset.val)); renderTags(); };
    });
    $(`#${prefix}-tag-count`).textContent = getTags().length + ' selected';
  };

  const addTag = val => {
    val = val.trim();
    if (!val || getTags().includes(val)) return;
    setTags([...getTags(), val]);
    renderTags();
    input.value = '';
    dropdown.classList.add('hidden');
  };

  input.oninput = () => {
    const q = input.value.toLowerCase();
    if (!q) { dropdown.classList.add('hidden'); return; }
    const matches = unique.filter(v => v.toLowerCase().includes(q) && !getTags().includes(v)).slice(0, 30);
    if (!matches.length) { dropdown.classList.add('hidden'); return; }
    dropdown.innerHTML = '';
    matches.forEach(m => {
      const item = document.createElement('div');
      item.className = 'ac-item';
      const idx = m.toLowerCase().indexOf(q);
      item.innerHTML = esc(m.slice(0,idx)) + `<span class="ac-match">${esc(m.slice(idx,idx+q.length))}</span>` + esc(m.slice(idx+q.length));
      item.onclick = () => addTag(m);
      dropdown.appendChild(item);
    });
    dropdown.classList.remove('hidden');
  };

  input.onkeydown = e => {
    if(e.key === 'Enter') { e.preventDefault(); const first = dropdown.querySelector('.ac-item'); if(first) first.click(); else if(input.value.trim()) addTag(input.value); }
    if(e.key === 'Escape') dropdown.classList.add('hidden');
  };

  addAllBtn.onclick = () => { setTags([...new Set([...getTags(), ...unique])]); renderTags(); };
  clearAllBtn.onclick = () => { setTags([]); renderTags(); };

  document.addEventListener('click', e => {
    if (!e.target.closest('.tag-input-wrapper')) dropdown.classList.add('hidden');
  });
  renderTags();
}

function clearTagInput(prefix) {
  $(`#${prefix}-tags`).innerHTML = '';
  $(`#${prefix}-filter-input`).value = '';
  $(`#${prefix}-autocomplete`).classList.add('hidden');
  if(prefix==='primary') state.wiz.primaryFilter = [];
  else state.wiz.measureFilter = [];
}

function collectTags(prefix) {
  // Tags are already stored in state.wiz via setTags
}

/* ═══════════ DATA AGGREGATION ═══════════ */
function aggregate(config) {
  const { primaryCol, primaryFilter, measureCol, measureFilter, aggregation } = config;
  const isSingle = measureCol === null || measureCol === undefined;
  const isNumeric = !isSingle && state.numericCols[measureCol];

  // Get primary labels
  let pLabels = [];
  const pCounts = {};

  if (primaryFilter.length) {
    pLabels = primaryFilter;
  } else {
    // Calculate total occurrences for all unique primary values to sort them
    state.rows.forEach(r => {
      const v = String(r[primaryCol] ?? '').trim();
      if (v) pCounts[v] = (pCounts[v] || 0) + 1;
    });
    pLabels = Object.keys(pCounts).sort((a, b) => pCounts[b] - pCounts[a]);
    
    // Take top 15, group rest
    if (pLabels.length > 15) {
      pLabels = pLabels.slice(0, 15);
      pLabels.push('Others (Combined)');
    }
  }

  if (isSingle) {
    // Count occurrences of each primary value
    const counts = {};
    pLabels.forEach(l => counts[l] = 0);
    const othersIdx = pLabels.indexOf('Others (Combined)');
    state.rows.forEach(r => {
      const v = String(r[primaryCol] ?? '').trim();
      if (v in counts) {
        counts[v]++;
      } else if (othersIdx !== -1 && v) {
        counts['Others (Combined)']++;
      }
    });
    return { primaryLabels: pLabels, measureLabels: [], matrix: [], singleData: pLabels.map(l => counts[l]) };
  }

  // Cross-tabulation
  let mLabels;
  if (measureFilter.length) mLabels = measureFilter;
  else mLabels = getUnique(measureCol).slice(0, 30);

  // Build matrix: [primaryIdx][measureIdx]
  const matrix = pLabels.map(() => mLabels.map(() => 0));
  const countMatrix = pLabels.map(() => mLabels.map(() => 0));

  const pMap = {}; pLabels.forEach((l, i) => pMap[l] = i);
  const mMap = {}; mLabels.forEach((l, i) => mMap[l] = i);
  const othersIdx = pLabels.indexOf('Others (Combined)');

  state.rows.forEach(r => {
    const pv = String(r[primaryCol] ?? '').trim();
    const mv_raw = r[measureCol];
    const mv = String(mv_raw ?? '').trim();
    
    let pi = pMap[pv];
    if (pi === undefined) {
      if (othersIdx !== -1 && pv) pi = othersIdx;
      else return; // filtered out
    }

    if (isNumeric && aggregation !== 'count') {
      // For numeric: aggregate the value
      const numVal = Number(mv_raw);
      if (isNaN(numVal)) return;
      // When aggregating numerically, we don't filter by mLabels — we sum all
      // But if user entered measure filters, they act as row filters
      if (measureFilter.length && !(mv in mMap)) return;
      // For numeric single-value aggregation, just put in index 0
      if (aggregation === 'sum' || aggregation === 'average') matrix[pi][0] = (matrix[pi][0] || 0) + numVal;
      if (aggregation === 'min') matrix[pi][0] = countMatrix[pi][0] ? Math.min(matrix[pi][0], numVal) : numVal;
      if (aggregation === 'max') matrix[pi][0] = countMatrix[pi][0] ? Math.max(matrix[pi][0], numVal) : numVal;
      countMatrix[pi][0]++;
    } else {
      // Categorical cross-tab: count
      if (!(mv in mMap)) return;
      const mi = mMap[mv];
      matrix[pi][mi]++;
    }
  });

  // Finalize averages
  if (isNumeric && aggregation === 'average') {
    matrix.forEach((row, pi) => { row.forEach((v, mi) => { if(countMatrix[pi][mi]) row[mi] = +(v / countMatrix[pi][mi]).toFixed(2); }); });
  }

  // For numeric aggregation, collapse to single-data format
  if (isNumeric && aggregation !== 'count') {
    return { primaryLabels: pLabels, measureLabels: [state.headers[measureCol]], matrix, singleData: pLabels.map((_, pi) => matrix[pi][0] || 0) };
  }

  return { primaryLabels: pLabels, measureLabels: mLabels, matrix, singleData: pLabels.map((_, pi) => matrix[pi].reduce((a,b)=>a+b, 0)) };
}

/* ═══════════ SMART COLUMN NORMALIZATION ═══════════ 
   Automatically assigns the optimal role to each column:
   Rule 1: Numeric primary + Text measure → SWAP (text should group, numeric should aggregate)
   Rule 2: Text + Text, primary has far more unique values → SWAP (lower cardinality = better axis labels)
   Rule 3: Both numeric → the one with fewer unique values becomes primary (grouping axis)
*/
function normalizeConfig(config) {
  if (config.measureCol === null || config.measureCol === undefined) return; // single-column, no swap needed
  
  const pIsNum = !!state.numericCols[config.primaryCol];
  const mIsNum = !!state.numericCols[config.measureCol];
  
  let shouldSwap = false;
  
  if (pIsNum && !mIsNum) {
    // Rule 1: User put numeric as primary, text as measure — always swap
    shouldSwap = true;
  } else if (!pIsNum && !mIsNum) {
    // Rule 2: Both text — put the one with fewer unique values as primary (axis labels)
    const pUnique = getUnique(config.primaryCol).length;
    const mUnique = getUnique(config.measureCol).length;
    // Swap if primary has significantly more unique values than measure
    if (pUnique > mUnique * 2 && mUnique <= 30) {
      shouldSwap = true;
    }
  } else if (pIsNum && mIsNum) {
    // Rule 3: Both numeric — the one with fewer unique values becomes primary
    const pUnique = getUnique(config.primaryCol).length;
    const mUnique = getUnique(config.measureCol).length;
    if (pUnique > mUnique) {
      shouldSwap = true;
    }
  }
  // If !pIsNum && mIsNum → already optimal, no swap
  
  if (shouldSwap) {
    // Swap column indices
    const tmpCol = config.primaryCol;
    config.primaryCol = config.measureCol;
    config.measureCol = tmpCol;
    
    // Swap filters too
    const tmpFilter = config.primaryFilter;
    config.primaryFilter = config.measureFilter;
    config.measureFilter = tmpFilter;
    
    // Re-detect aggregation for the new measure column
    if (state.numericCols[config.measureCol]) {
      // New measure is numeric — use sum if aggregation was count
      if (config.aggregation === 'count') config.aggregation = 'sum';
    } else {
      // New measure is text — force count
      config.aggregation = 'count';
    }
  }
}

/* ═══════════ GENERATE CHARTS ═══════════ */
function generateAllCharts() {
  const config = { ...state.wiz };
  
  // Smart normalization: auto-detect optimal column roles
  normalizeConfig(config);
  
  const aggResult = aggregate(config);
  const title = generateTitle(config, state.headers);

  config.chartTypes.forEach(ct => {
    const id = ++chartIdCounter;
    addChartCard(id, ct, title, aggResult, config);
    state.charts.push({ id, type: ct, title, config: { ...config } });
  });

  if (!$('#dashboard-section').classList.contains('hidden') || state.charts.length > 0) {
    showSection('dashboard-section');
    // Also keep nav viz button
    $('#nav-add-viz').classList.remove('hidden');
  }
}

function addChartCard(id, type, title, aggResult, config) {
  const grid = $('#dashboard-grid');
  const card = document.createElement('div');
  card.className = 'chart-card';
  card.id = 'chart-card-' + id;

  const typeName = CHART_TYPES[type]?.name || type;

  card.innerHTML = `
    <div class="chart-card-header">
      <h4>${esc(title)} — ${typeName}</h4>
      <div class="chart-card-actions">
        <button title="Fullscreen" data-action="expand">⛶</button>
        <button title="Download PNG" data-action="download">⬇</button>
        <button title="Remove" data-action="remove">✕</button>
      </div>
    </div>
    <div class="chart-card-body" id="chart-body-${id}"></div>
    <div class="chart-card-footer">
      <span class="chart-meta">📊 ${typeName}</span>
      <span class="chart-meta">📋 ${state.headers[config.primaryCol]}</span>
      ${config.measureCol !== null ? `<span class="chart-meta">↔ ${state.headers[config.measureCol]}</span>` : ''}
      ${config.aggregation !== 'count' ? `<span class="chart-meta">Σ ${config.aggregation}</span>` : ''}
    </div>`;

  grid.appendChild(card);

  // Render chart
  const body = $(`#chart-body-${id}`);
  if (type === 'heatmap') {
    body.innerHTML = buildHeatmapHTML(aggResult);
  } else {
    const canvas = document.createElement('canvas');
    canvas.id = 'canvas-' + id;
    body.appendChild(canvas);
    const chartInstance = buildChart(canvas, type, aggResult, config, state.headers);
    card._chartInstance = chartInstance;
  }

  // Actions
  card.querySelector('[data-action="remove"]').onclick = () => {
    if (card._chartInstance) card._chartInstance.destroy();
    card.remove();
    state.charts = state.charts.filter(c => c.id !== id);
    if (!state.charts.length) {
      showSection('preview-section');
    }
  };

  card.querySelector('[data-action="download"]').onclick = () => {
    const canvas = card.querySelector('canvas');
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = title.replace(/\s+/g, '_') + '.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  };

  card.querySelector('[data-action="expand"]').onclick = () => {
    const overlay = $('#fullscreen-overlay');
    const container = $('#fullscreen-chart-container');
    $('#fullscreen-title').textContent = title + ' — ' + typeName;
    container.innerHTML = '';
    if (type === 'heatmap') {
      container.innerHTML = buildHeatmapHTML(aggResult);
    } else {
      const wrapper = document.createElement('div');
      wrapper.style.width = '100%';
      const canvas2 = document.createElement('canvas');
      wrapper.appendChild(canvas2);
      container.appendChild(wrapper);
      buildChart(canvas2, type, aggResult, config, state.headers);
    }
    overlay.classList.remove('hidden');
  };
}

/* ═══════════ RESET ═══════════ */
function resetAll() {
  state.charts.forEach(c => {
    const card = $(`#chart-card-${c.id}`);
    if (card?._chartInstance) card._chartInstance.destroy();
  });
  state.workbook = null; state.sheetNames = []; state.headers = []; state.rows = [];
  state.uniqueCache = {}; state.numericCols = {}; state.charts = [];
  $('#dashboard-grid').innerHTML = '';
  $('#top-nav').classList.add('hidden');
  $('#nav-add-viz').classList.add('hidden');
  showSection('upload-section');
  $('#file-input').value = '';
}

/* ═══════════ UTILS ═══════════ */
function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
function shake(el) { el.style.animation = 'none'; el.offsetHeight; el.style.animation = 'shake .4s ease'; }
