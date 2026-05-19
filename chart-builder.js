/* ═══════════════════════════════════════════
   chart-builder.js — Chart creation helpers
   ═══════════════════════════════════════════ */

const PALETTE = [
  '#c67a4a','#4a7c59','#d4a057','#6b8cae','#a0522d',
  '#8b5e3c','#5f8a8b','#9e6b2f','#7b68ee','#2e8b57',
  '#cd853f','#708090','#b8860b','#6a9b6a','#d2691e',
  '#8fbc8f','#556b2f','#bc8f8f','#daa520','#708238'
];

const CHART_TYPES = {
  bar:          { icon: '📊', name: 'Bar Chart',         desc: 'Vertical bars',          mode: 'all' },
  barH:         { icon: '📶', name: 'Horizontal Bar',    desc: 'Horizontal bars',         mode: 'all' },
  stacked:      { icon: '📚', name: 'Stacked Bar',       desc: 'Stacked composition',     mode: 'cross' },
  stackedH:     { icon: '📑', name: 'Stacked Horiz.',    desc: 'Horizontal stacked',      mode: 'cross' },
  pie:          { icon: '🥧', name: 'Pie Chart',         desc: 'Proportions',             mode: 'single' },
  doughnut:     { icon: '🍩', name: 'Donut Chart',       desc: 'Ring proportions',        mode: 'single' },
  radar:        { icon: '🕸️', name: 'Radar Chart',       desc: 'Multi-axis comparison',   mode: 'cross' },
  polarArea:    { icon: '🎯', name: 'Polar Area',        desc: 'Radial proportions',      mode: 'single' },
  heatmap:      { icon: '🔥', name: 'Heatmap',           desc: 'Density grid',            mode: 'cross' },
};

function getColor(i, alpha) {
  const c = PALETTE[i % PALETTE.length];
  if (!alpha) return c;
  const r = parseInt(c.slice(1,3),16), g = parseInt(c.slice(3,5),16), b = parseInt(c.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

/* Format numbers — Indian system (lakhs, crores) */
function formatNum(n) {
  if (n === null || n === undefined) return '0';
  const num = Number(n);
  if (isNaN(num)) return String(n);
  const fixed = parseFloat(num.toPrecision(12));
  const abs = Math.abs(fixed);
  if (abs >= 1e7) return (fixed / 1e7).toFixed(2) + ' Cr';
  if (abs >= 1e5) return (fixed / 1e5).toFixed(2) + ' L';
  if (abs >= 1e3) return fixed.toLocaleString('en-IN', { maximumFractionDigits: 2 });
  return Number.isInteger(fixed) ? String(fixed) : fixed.toFixed(2);
}

function isDark() {
  return document.documentElement.getAttribute('data-theme') !== 'light';
}

function themeColors() {
  const dark = isDark();
  return {
    tick: dark ? '#b8a898' : '#5a4d3e',
    grid: dark ? 'rgba(255,248,240,.07)' : 'rgba(45,36,25,.08)',
    legend: dark ? '#b8a898' : '#5a4d3e',
    tooltip: dark ? 'rgba(35,30,24,.95)' : 'rgba(255,255,255,.95)',
    tooltipText: dark ? '#f5ede4' : '#2d2419',
    tooltipBorder: dark ? 'rgba(255,248,240,.15)' : 'rgba(45,36,25,.12)',
    axisTitle: dark ? '#d4a057' : '#a0522d'
  };
}

/* Calculate dynamic canvas height based on data */
function calcHeight(type, numLabels, numDatasets) {
  if (type === 'barH' || type === 'stackedH') {
    // Each category group needs enough height for all dataset bars
    let perLabel;
    if (numDatasets <= 1) perLabel = 36;
    else if (numDatasets <= 5) perLabel = 50;
    else if (numDatasets <= 10) perLabel = 65;
    else perLabel = 80;
    return Math.max(400, numLabels * perLabel + 100);
  }
  if (type === 'bar' || type === 'stacked') {
    return Math.max(400, 520);
  }
  return 420;
}

/* ─── Build a chart and return the Chart instance ─── */
function buildChart(canvas, type, aggResult, config, headers) {
  const { primaryLabels, measureLabels, matrix, singleData } = aggResult;
  const isSingle = !config.measureCol && config.measureCol !== 0;
  const tc = themeColors();

  const chartJsType = getChartJsType(type);
  if (!chartJsType) return null;

  const primaryName = headers ? headers[config.primaryCol] || 'Primary' : 'Primary';
  const measureName = (headers && config.measureCol != null) ? headers[config.measureCol] || 'Measure' : '';
  const aggLabel = config.aggregation === 'count' ? 'Count (Number of Records)' : config.aggregation.charAt(0).toUpperCase() + config.aggregation.slice(1);

  // Dynamic height
  const numDS = isSingle ? 1 : measureLabels.length;
  const h = calcHeight(type, primaryLabels.length, numDS);
  canvas.parentElement.style.height = h + 'px';
  canvas.style.height = h + 'px';

  // Custom plugin for data labels on bars
  const dataLabelPlugin = {
    id: 'barValueLabels',
    afterDatasetsDraw(chart) {
      const chartType = type;
      if (!['bar','barH','stacked','stackedH'].includes(chartType)) return;
      const { ctx } = chart;
      const dark = isDark();
      ctx.save();
      ctx.font = '600 11px Inter';
      ctx.fillStyle = dark ? '#f5ede4' : '#2d2419';

      if (chartType === 'stacked' || chartType === 'stackedH') {
        // For stacked: show total on top of the last (top) bar segment
        const meta = chart.getDatasetMeta(chart.data.datasets.length - 1);
        meta.data.forEach((bar, idx) => {
          let total = 0;
          chart.data.datasets.forEach(ds => { total += (ds.data[idx] || 0); });
          if (total === 0) return;
          const label = formatNum(total);
          if (chartType === 'stackedH') {
            const x = bar.x + 6;
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            ctx.fillText(label, x, bar.y);
          } else {
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(label, bar.x, bar.y - 4);
          }
        });
      } else {
        // For grouped: show value on each individual bar
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          meta.data.forEach((bar, idx) => {
            const val = ds.data[idx];
            if (!val || val === 0) return;
            const label = formatNum(val);
            if (chartType === 'barH') {
              ctx.textAlign = 'left';
              ctx.textBaseline = 'middle';
              ctx.fillText(label, bar.x + 6, bar.y);
            } else {
              ctx.textAlign = 'center';
              ctx.textBaseline = 'bottom';
              ctx.fillText(label, bar.x, bar.y - 4);
            }
          });
        });
      }
      ctx.restore();
    }
  };

  const isBarType = ['bar','barH','stacked','stackedH'].includes(type);

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 800, easing: 'easeOutQuart' },
    layout: { padding: { top: 24, bottom: 10, left: 5, right: isBarType && (type === 'barH' || type === 'stackedH') ? 80 : 20 } },
    plugins: {
      legend: {
        display: !isSingle || type === 'pie' || type === 'doughnut' || type === 'polarArea',
        position: 'bottom',
        labels: {
          color: tc.legend, padding: 14, usePointStyle: true, pointStyle: 'rectRounded',
          font: { family: 'Inter', size: 12, weight: '500' }
        }
      },
      tooltip: {
        backgroundColor: tc.tooltip,
        titleColor: tc.tooltipText,
        bodyColor: tc.tooltipText,
        titleFont: { family: 'Inter', weight: '700', size: 13 },
        bodyFont: { family: 'Inter', size: 12 },
        padding: 14,
        cornerRadius: 10,
        borderColor: tc.tooltipBorder,
        borderWidth: 1,
        displayColors: true,
        boxPadding: 6,
        callbacks: {
          title: function(items) {
            const idx = items[0].dataIndex;
            return primaryLabels[idx] || items[0].label || '';
          },
          label: function(ctx) {
            const ds = ctx.dataset.label || '';
            let rawVal;
            if (type === 'pie' || type === 'doughnut' || type === 'polarArea') {
              rawVal = ctx.parsed;  // pie/doughnut/polar: parsed IS the number
            } else if (type === 'barH' || type === 'stackedH') {
              rawVal = ctx.parsed.x;
            } else {
              rawVal = ctx.parsed.y;
            }
            const val = formatNum(rawVal);
            const unit = config.aggregation === 'count' ? ' records' : '';
            return `${ds ? ds + ': ' : ''}${val}${unit}`;
          }
        }
      }
    }
  };

  const displayLabels = primaryLabels.map(l => l.length > 25 ? l.substring(0, 22) + '...' : l);

  let data;
  if (isSingle) {
    data = buildSingleData(type, displayLabels, singleData);
  } else {
    data = buildCrossData(type, displayLabels, measureLabels, matrix);
  }

  addTypeOptions(type, options, isSingle, tc, primaryName, measureName, aggLabel);

  const plugins = isBarType ? [dataLabelPlugin] : [];
  return new Chart(canvas, { type: chartJsType, data, options, plugins });
}

function getChartJsType(t) {
  const map = { bar:'bar', barH:'bar', stacked:'bar', stackedH:'bar', pie:'pie', doughnut:'doughnut', radar:'radar', polarArea:'polarArea' };
  return map[t] || null;
}

/* ─── Single column data ─── */
function buildSingleData(type, labels, values) {
  if (type === 'pie' || type === 'doughnut' || type === 'polarArea') {
    return {
      labels,
      datasets: [{
        data: values,
        backgroundColor: labels.map((_, i) => getColor(i, 0.8)),
        borderColor: isDark() ? 'rgba(35,30,24,.6)' : 'rgba(255,255,255,.8)',
        borderWidth: 2,
        hoverOffset: 12
      }]
    };
  }
  return {
    labels,
    datasets: [{
      label: 'Count',
      data: values,
      backgroundColor: labels.map((_, i) => getColor(i, 0.8)),
      borderColor: labels.map((_, i) => getColor(i)),
      borderWidth: 1,
      borderRadius: 6,
      borderSkipped: false,
      barPercentage: 0.7,
      categoryPercentage: 0.8,
      maxBarThickness: 50
    }]
  };
}

/* ─── Cross-tabulation data ─── */
function buildCrossData(type, primaryLabels, measureLabels, matrix) {
  const datasets = measureLabels.map((ml, mi) => ({
    label: String(ml),
    data: primaryLabels.map((_, pi) => matrix[pi]?.[mi] || 0),
    backgroundColor: getColor(mi, 0.85),
    borderColor: getColor(mi),
    borderWidth: 1,
    borderRadius: (type === 'stacked' || type === 'stackedH') ? 0 : 4,
    borderSkipped: false,
    barPercentage: 0.8,
    categoryPercentage: 0.85,
    maxBarThickness: 40
  }));
  return { labels: primaryLabels, datasets };
}

/* ─── Type-specific Chart.js options with AXIS LABELS ─── */
function addTypeOptions(type, opts, isSingle, tc, primaryName, measureName, aggLabel) {
  const tickStyle = { color: tc.tick, font: { family: 'Inter', size: 12, weight: '500' } };
  const gridStyle = { color: tc.grid };
  const titleStyle = { display: true, color: tc.axisTitle, font: { family: 'Outfit', size: 13, weight: '700' }, padding: 8 };

  const valueAxisTitle = isSingle ? 'Count (Number of Records)' : aggLabel;
  const catAxisTitle = primaryName;

  if (type === 'bar' || type === 'stacked') {
    opts.indexAxis = 'x';
    opts.scales = {
      x: { title: { ...titleStyle, text: catAxisTitle }, ticks: { ...tickStyle, maxRotation: 45, minRotation: 45, autoSkip: false }, grid: gridStyle, stacked: type === 'stacked' },
      y: { title: { ...titleStyle, text: valueAxisTitle }, ticks: tickStyle, grid: gridStyle, stacked: type === 'stacked', beginAtZero: true }
    };
  } else if (type === 'barH' || type === 'stackedH') {
    opts.indexAxis = 'y';
    opts.scales = {
      x: { title: { ...titleStyle, text: valueAxisTitle }, ticks: tickStyle, grid: gridStyle, stacked: type === 'stackedH', beginAtZero: true },
      y: { title: { ...titleStyle, text: catAxisTitle }, ticks: { ...tickStyle, font: { family: 'Inter', size: 11, weight: '600' }, autoSkip: false }, grid: gridStyle, stacked: type === 'stackedH' }
    };
  } else if (type === 'radar') {
    opts.scales = {
      r: { ticks: { ...tickStyle, backdropColor: 'transparent' }, grid: gridStyle, pointLabels: { ...tickStyle, font: { family: 'Inter', size: 12, weight: '600' } }, beginAtZero: true }
    };
  } else if (type === 'polarArea') {
    opts.scales = { r: { ticks: { ...tickStyle, backdropColor: 'transparent' }, grid: gridStyle, beginAtZero: true } };
  }
}

/* ─── Heatmap (returns HTML string) ─── */
function buildHeatmapHTML(aggResult) {
  const { primaryLabels, measureLabels, matrix } = aggResult;
  const allVals = matrix.flat();
  const maxVal = Math.max(...allVals, 1);

  let html = '<div class="heatmap-container"><table class="heatmap-table"><thead><tr><th></th>';
  measureLabels.forEach(ml => { html += `<th>${escHtml(String(ml))}</th>`; });
  html += '</tr></thead><tbody>';

  primaryLabels.forEach((pl, pi) => {
    html += `<tr><td>${escHtml(String(pl))}</td>`;
    measureLabels.forEach((_, mi) => {
      const v = matrix[pi]?.[mi] || 0;
      const intensity = v / maxVal;
      const bg = heatColor(intensity);
      const textColor = intensity > 0.45 ? '#fff' : (isDark() ? '#b8a898' : '#5a4d3e');
      html += `<td style="background:${bg};color:${textColor}">${formatNum(v)}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  // Color legend
  html += '<div class="heatmap-legend">';
  html += '<span class="heatmap-legend-label">Low</span>';
  html += '<div class="heatmap-legend-bar"></div>';
  html += '<span class="heatmap-legend-label">High</span>';
  html += '<div class="heatmap-legend-info">Color intensity represents the magnitude of values. Darker/warmer colors indicate higher values.</div>';
  html += '</div>';
  html += '</div>';
  return html;
}

function heatColor(t) {
  const r = Math.round(200 + t * 50);
  const g = Math.round(190 - t * 130);
  const b = Math.round(150 - t * 110);
  return `rgba(${Math.min(r,255)},${Math.max(g,0)},${Math.max(b,0)},${Math.max(0.12, t * 0.95)})`;
}

function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

/* ─── Summary stats ─── */
function buildSummaryHTML(aggResult) {
  const { primaryLabels, singleData } = aggResult;
  const total = singleData.reduce((a, b) => a + b, 0);
  const maxIdx = singleData.indexOf(Math.max(...singleData));
  const minIdx = singleData.indexOf(Math.min(...singleData));

  let html = '<div class="summary-grid">';
  html += summaryCard(total, 'Total', '#c67a4a');
  html += summaryCard(primaryLabels.length, 'Unique Values', '#4a7c59');
  html += summaryCard(primaryLabels[maxIdx] || '-', `Highest (${singleData[maxIdx]})`, '#d4a057');
  html += summaryCard(primaryLabels[minIdx] || '-', `Lowest (${singleData[minIdx]})`, '#a0522d');
  html += '</div>';
  return html;
}

function summaryCard(value, label, color) {
  return `<div class="summary-item"><span class="s-value" style="color:${color}">${value}</span><span class="s-label">${label}</span></div>`;
}

/* ─── Generate chart title ─── */
function generateTitle(config, headers) {
  const primary = headers[config.primaryCol] || 'Column';
  if (config.measureCol === null || config.measureCol === undefined) {
    return `Distribution of ${primary}`;
  }
  const measure = headers[config.measureCol] || 'Column';
  const aggLabel = config.aggregation === 'count' ? '' : ` (${config.aggregation})`;
  return `${measure}${aggLabel} by ${primary}`;
}
