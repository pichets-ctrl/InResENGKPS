/* ============================================================
   Research Dashboard — app.js  v3
   ─────────────────────────────────────────────────────────────
   Modules:
     parseExcel()        – FileReader + SheetJS
     normalizeHeaders()  – merge row3 & row4 → "Main - Sub"
     buildDataset()      – map rows 5+ → object array
     enrichBudget()      – compute __budget_num / _internal / _external
     renderSummary()     – update stat cards with count-up
     renderCharts()      – Chart.js charts (4 types)
     renderTable()       – paginated sortable table
     applyFilters()      – filter + trigger re-render
   ============================================================ */

'use strict';

// ─────────────────────────────────────────────
//  Constants
// ─────────────────────────────────────────────
const TARGET_SHEET   = 'โครงการวิจัย';
const HEADER_ROW     = 3;   // 1-based: main field headers
const DATA_ROW_START = 5;   // 1-based: first data row

// Semantic column keyword candidates
const SEMANTIC_KEYS = {
  projectName:    ['ชื่อโครงการ','ชื่อโครงการวิจัย','project name','โครงการ'],
  dept:           ['หน่วยงาน','ภาควิชา','คณะ','สาขา','department','dept','faculty'],
  researcher:     ['หัวหน้าโครงการ','ชื่อหัวหน้า','นักวิจัย','ผู้วิจัย','researcher','pi'],
  budget:         ['งบประมาณรวม','รวมงบประมาณ','งบประมาณ','budget','เงินทุน','ทุนวิจัย','วงเงิน'],
  budgetInternal: ['งบประมาณภายใน','งบภายใน','เงินทุนภายใน','ภายใน'],
  budgetExternal: ['งบประมาณภายนอก','งบภายนอก','เงินทุนภายนอก','ภายนอก'],
  type:           ['ประเภทวิจัย','ประเภทโครงการ','ประเภท','type'],
  year:           ['ปีงบประมาณ','ปี','year'],
  status:         ['สถานะ','status'],
  fundSource:     ['แหล่งทุน','แหล่งเงินทุน','funding source','แหล่งงบ'],
};

// Chart color palette
const PALETTE = [
  '#f59e0b','#10b981','#3b82f6','#8b5cf6','#ef4444',
  '#06b6d4','#f97316','#84cc16','#ec4899','#6366f1',
  '#14b8a6','#f43f5e','#a78bfa','#34d399','#fbbf24',
];

// ─────────────────────────────────────────────
//  App State
// ─────────────────────────────────────────────
const state = {
  rawData:      [],
  filteredData: [],
  columns:      [],
  colMap:       {},
  charts:       {},
  sort:         { col: null, dir: 'asc' },
  page:         1,
  pageSize:     10,
};

// ─────────────────────────────────────────────
//  Utility helpers
// ─────────────────────────────────────────────
const $        = (id) => document.getElementById(id);
const showEl   = (id) => { const e=$( id); if (e) e.classList.remove('hidden'); };
const hideEl   = (id) => { const e=$(id); if (e) e.classList.add('hidden'); };
const uniqueVals = (arr, key) =>
  [...new Set(arr.map(r => r[key]).filter(Boolean))].sort((a,b) =>
    String(a).localeCompare(String(b), 'th'));

const toNum = (val) => {
  const n = parseFloat(String(val ?? '').replace(/,/g, '').trim());
  return isNaN(n) ? 0 : n;
};

const fmtBudget = (n) => {
  const num = parseFloat(String(n).replace(/,/g, ''));
  if (isNaN(num)) return '–';
  return (num / 10000).toLocaleString('th-TH', {
    minimumFractionDigits: 2, maximumFractionDigits: 2,
  }) + ' หมื่น';
};

function debounce(fn, ms) {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}

function showLoading(msg = 'กำลังโหลดข้อมูล...') {
  $('loadingMsg').textContent = msg;
  showEl('loadingOverlay');
}
function hideLoading() { hideEl('loadingOverlay'); }

// ─────────────────────────────────────────────
//  Count-up animation for stat cards
// ─────────────────────────────────────────────
function countUp(el, endVal, formatter, duration = 700) {
  const t0 = performance.now();
  const run = (now) => {
    const p    = Math.min((now - t0) / duration, 1);
    const ease = 1 - Math.pow(1 - p, 3);   // cubic ease-out
    el.textContent = formatter(endVal * ease);
    if (p < 1) requestAnimationFrame(run);
  };
  requestAnimationFrame(run);
}

// ═══════════════════════════════════════════════
//  MODULE 1 — parseExcel
//  Entry point: reads file via FileReader + SheetJS
// ═══════════════════════════════════════════════
function parseExcel(file) {
  showLoading('กำลังอ่านไฟล์ Excel...');
  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });

      if (!wb.SheetNames.includes(TARGET_SHEET)) {
        hideLoading();
        alert(`ไม่พบ sheet "${TARGET_SHEET}"\nSheet ที่พบ: ${wb.SheetNames.join(', ')}`);
        return;
      }

      showLoading(`กำลังอ่าน sheet: ${TARGET_SHEET}`);
      const ws      = wb.Sheets[TARGET_SHEET];
      const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      const headers = normalizeHeaders(ws, allRows);
      const dataset = buildDataset(allRows, headers);

      if (!dataset.length) {
        hideLoading();
        alert(`ไม่พบข้อมูลใน sheet "${TARGET_SHEET}" (ตั้งแต่แถวที่ ${DATA_ROW_START})`);
        return;
      }

      state.rawData = dataset;
      state.columns = headers;
      state.colMap  = detectSemanticColumns(headers);
      state.sort    = { col: null, dir: 'asc' };
      state.page    = 1;

      $('lastUpdated').textContent = `อัปเดต: ${new Date().toLocaleString('th-TH')}`;
      showEl('lastUpdated');

      showLoading('กำลังประมวลผลข้อมูล...');
      setTimeout(() => {
        enrichBudget();
        populateFilters();
        applyFilters();
        hideEl('uploadPrompt');
        showEl('dashboard');
        hideLoading();
      }, 50);

    } catch (err) {
      hideLoading();
      console.error(err);
      alert('เกิดข้อผิดพลาดในการอ่านไฟล์: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════
//  MODULE 2 — normalizeHeaders
//  Expands merged cells, combines row3 + row4 → "Main - Sub"
// ═══════════════════════════════════════════════
function normalizeHeaders(ws, allRows) {
  // Expand merged cells so every cell in a merge holds the top-left value
  const merges   = ws['!merges'] || [];
  const expanded = allRows.map(r => [...r]);

  merges.forEach(m => {
    const topVal = expanded[m.s.r]?.[m.s.c] ?? '';
    for (let r = m.s.r; r <= m.e.r; r++) {
      if (!expanded[r]) continue;
      for (let c = m.s.c; c <= m.e.c; c++) expanded[r][c] = topVal;
    }
  });

  const row3    = expanded[HEADER_ROW - 1] || [];  // index 2 = Excel row 3
  const row4    = expanded[HEADER_ROW]     || [];  // index 3 = Excel row 4
  const maxCols = Math.max(row3.length, row4.length);
  const headers = [];

  for (let i = 0; i < maxCols; i++) {
    const h3 = String(row3[i] ?? '').trim();
    const h4 = String(row4[i] ?? '').trim();
    let name;
    if      (h3 && h4 && h3 !== h4) name = `${h3} - ${h4}`;  // "Main - Sub"
    else if (h3)                     name = h3;
    else if (h4)                     name = h4;
    else                             name = `คอลัมน์_${i + 1}`;
    headers.push(name);
  }

  return headers;
}

// ═══════════════════════════════════════════════
//  MODULE 3 — buildDataset
//  Maps data rows (row 5 onward) to array of objects
// ═══════════════════════════════════════════════
function buildDataset(allRows, headers) {
  return allRows
    .slice(DATA_ROW_START - 1)                  // start at index 4 (row 5)
    .filter(row =>                              // drop completely empty rows
      row.some(c => c !== null && c !== undefined && String(c).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i] ?? '';
        if      (val instanceof Date)       val = val.toLocaleDateString('th-TH');
        else if (typeof val === 'string')   val = val.trim();
        obj[h] = val;
      });
      return obj;
    });
}

// ─────────────────────────────────────────────
//  Semantic column detection
//  Detects budgetInternal/External first to avoid
//  "งบประมาณ" matching sub-columns incorrectly
// ─────────────────────────────────────────────
function detectSemanticColumns(headers) {
  const map  = {};
  const used = new Set();
  const priority = ['budgetInternal', 'budgetExternal'];
  const rest     = Object.keys(SEMANTIC_KEYS).filter(k => !priority.includes(k));

  for (const sem of [...priority, ...rest]) {
    for (const h of headers) {
      if (used.has(h)) continue;
      if (SEMANTIC_KEYS[sem].some(c => h.toLowerCase().includes(c.toLowerCase()))) {
        map[sem] = h;
        used.add(h);
        break;
      }
    }
  }
  return map;
}

// ─────────────────────────────────────────────
//  Enrich rows with computed budget fields
//  Always sums internal + external when available
// ─────────────────────────────────────────────
function enrichBudget() {
  const { budget: bc, budgetInternal: bic, budgetExternal: bec } = state.colMap;

  state.rawData.forEach(row => {
    const internal = bic ? toNum(row[bic]) : 0;
    const external = bec ? toNum(row[bec]) : 0;

    row['__budget_num']      = (bic || bec) ? internal + external
                             : bc           ? toNum(row[bc])
                             : 0;
    row['__budget_internal'] = internal;
    row['__budget_external'] = external;
  });
}

// ─────────────────────────────────────────────
//  Populate filter dropdowns
// ─────────────────────────────────────────────
function populateFilters() {
  const { dept: dc, type: tc } = state.colMap;

  $('deptFilter').innerHTML = '<option value="">ทุกหน่วยงาน</option>';
  $('typeFilter').innerHTML = '<option value="">ทุกประเภทวิจัย</option>';

  if (dc) uniqueVals(state.rawData, dc).forEach(v =>
    $('deptFilter').insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`));
  if (tc) uniqueVals(state.rawData, tc).forEach(v =>
    $('typeFilter').insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`));
}

// ═══════════════════════════════════════════════
//  MODULE 7 — applyFilters
//  Filters rawData → filteredData then re-renders all
// ═══════════════════════════════════════════════
function applyFilters() {
  const search = $('searchInput').value.toLowerCase().trim();
  const dept   = $('deptFilter').value;
  const type   = $('typeFilter').value;
  const budget = $('budgetFilter').value;
  const { dept: dc, type: tc } = state.colMap;

  state.filteredData = state.rawData.filter(row => {
    if (search) {
      const txt = Object.entries(row)
        .filter(([k]) => !k.startsWith('__'))
        .map(([, v]) => String(v)).join(' ').toLowerCase();
      if (!txt.includes(search)) return false;
    }
    if (dept   && dc && row[dc] !== dept) return false;
    if (type   && tc && row[tc] !== type) return false;
    if (budget) {
      const [lo, hi] = budget.split('-').map(Number);
      const b = row['__budget_num'];
      if (b < lo || b > hi) return false;
    }
    return true;
  });

  state.page = 1;
  $('resultsCount').textContent =
    `พบ ${state.filteredData.length.toLocaleString('th-TH')} รายการ`;

  renderSummary(state.filteredData);
  renderCharts(state.filteredData);
  renderTable();
}

// ═══════════════════════════════════════════════
//  MODULE 4 — renderSummary
//  Updates stat cards with count-up animation
// ═══════════════════════════════════════════════
function renderSummary(data) {
  const { dept: dc, researcher: rc } = state.colMap;
  const totalBudget = data.reduce((s, r) => s + r['__budget_num'], 0);

  countUp($('totalProjects'), data.length,
    v => Math.round(v).toLocaleString('th-TH'));

  // Budget formatted directly (not count-up — complex format)
  $('totalBudget').textContent = fmtBudget(totalBudget);

  if (dc) {
    countUp($('totalDepts'), uniqueVals(data, dc).length,
      v => Math.round(v).toLocaleString('th-TH'));
  } else {
    $('totalDepts').textContent = '–';
  }

  if (rc) {
    countUp($('totalResearchers'), uniqueVals(data, rc).length,
      v => Math.round(v).toLocaleString('th-TH'));
  } else {
    $('totalResearchers').textContent = '–';
  }
}

// ═══════════════════════════════════════════════
//  MODULE 5 — renderCharts
// ═══════════════════════════════════════════════
function destroyChart(id) {
  if (state.charts[id]) { state.charts[id].destroy(); delete state.charts[id]; }
}

function groupBy(data, key) {
  return data.reduce((acc, r) => {
    const k = String(r[key] || '(ไม่ระบุ)').trim();
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {});
}

function groupBudgetBy(data, key) {
  return data.reduce((acc, r) => {
    const k = String(r[key] || '(ไม่ระบุ)').trim();
    acc[k] = (acc[k] || 0) + r['__budget_num'];
    return acc;
  }, {});
}

function renderEmpty(ctx, id) {
  const c = $(id);
  ctx.clearRect(0, 0, c.width, c.height);
  ctx.font         = '14px Kanit, sans-serif';
  ctx.fillStyle    = '#d1d5db';
  ctx.textAlign    = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('ไม่พบข้อมูลสำหรับแสดงผล', c.width / 2, c.height / 2);
}

function renderCharts(data) {
  const { dept: dc, type: tc, researcher: rc } = state.colMap;
  renderDeptBar(data, dc);
  renderBudgetDoughnut(data, dc);
  renderBudgetHBar(data, rc || dc);
  renderTypeBar(data, tc);
}

// Chart 1 — Bar: project count per department
function renderDeptBar(data, col) {
  destroyChart('deptBarChart');
  const ctx = $('deptBarChart').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'deptBarChart'); return; }

  const grouped = groupBy(data, col);
  const sorted  = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 15);
  const labels  = sorted.map(e => e[0]);
  const values  = sorted.map(e => e[1]);

  state.charts['deptBarChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'จำนวนโครงการ',
        data: values,
        backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length] + 'bb'),
        borderColor:     labels.map((_, i) => PALETTE[i % PALETTE.length]),
        borderWidth: 1.5,
        borderRadius: 6,
        borderSkipped: false,
      }],
    },
    options: barOpts('จำนวนโครงการ', false),
  });
}

// Chart 2 — Doughnut: budget proportion
function renderBudgetDoughnut(data, col) {
  destroyChart('budgetDoughnut');
  const ctx = $('budgetDoughnut').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'budgetDoughnut'); return; }

  const grouped = groupBudgetBy(data, col);
  const sorted  = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const labels  = sorted.map(e => e[0]);
  const values  = sorted.map(e => e[1]);
  const total   = values.reduce((s, v) => s + v, 0);

  state.charts['budgetDoughnut'] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length] + 'dd'),
        borderColor: '#ffffff',
        borderWidth: 2,
        hoverOffset: 12,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '60%',
      animation: { animateRotate: true, duration: 900 },
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { family: 'Kanit', size: 11 },
            boxWidth: 12,
            padding: 10,
            generateLabels: (chart) => {
              const ds = chart.data.datasets[0];
              return chart.data.labels.map((lbl, i) => {
                const pct   = total > 0 ? ((ds.data[i] / total) * 100).toFixed(1) : 0;
                const short = lbl.length > 15 ? lbl.slice(0, 14) + '…' : lbl;
                return { text: `${short} (${pct}%)`, fillStyle: ds.backgroundColor[i], hidden: false, index: i };
              });
            },
          },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${fmtBudget(ctx.parsed)} (${pct}%)`;
            },
          },
          bodyFont: { family: 'Kanit' },
          titleFont: { family: 'Kanit' },
        },
      },
    },
  });
}

// Chart 3 — Horizontal bar: budget per researcher
function renderBudgetHBar(data, col) {
  destroyChart('budgetHBar');
  const ctx = $('budgetHBar').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'budgetHBar'); return; }

  const grouped = groupBudgetBy(data, col);
  const sorted  = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 12);
  const labels  = sorted.map(e => e[0]);
  const values  = sorted.map(e => e[1]);

  state.charts['budgetHBar'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'งบประมาณ (หมื่นบาท)',
        data: values,
        backgroundColor: labels.map((_, i) => PALETTE[(i + 2) % PALETTE.length] + '99'),
        borderColor:     labels.map((_, i) => PALETTE[(i + 2) % PALETTE.length]),
        borderWidth: 1.5,
        borderRadius: 4,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 700, easing: 'easeOutQuart' },
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => ` งบประมาณ: ${fmtBudget(ctx.parsed.x)}`,
          },
          bodyFont: { family: 'Kanit' },
          titleFont: { family: 'Kanit' },
        },
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            font: { family: 'Kanit', size: 10 },
            callback: function(v) { return fmtBudget(v); },
          },
          grid: { color: '#f3f4f6' },
        },
        y: {
          title: {
            display: true,
            text: 'ชื่อหัวหน้าโครงการ',
            font: { family: 'Kanit', size: 11 },
            color: '#6b7280',
          },
          ticks: {
            font: { family: 'Kanit', size: 10 },
            callback: function(v) {
              const s = String(v);
              return s.length > 20 ? s.slice(0, 19) + '…' : s;
            },
          },
          grid: { color: '#f3f4f6' },
        },
      },
    },
  });
}

// Chart 4 — Grouped bar: internal vs external by research type
function renderTypeBar(data, col) {
  destroyChart('typeBarChart');
  const ctx = $('typeBarChart').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'typeBarChart'); return; }

  const agg = {};
  data.forEach(r => {
    const t = String(r[col] || '(ไม่ระบุ)').trim();
    if (!agg[t]) agg[t] = { internal: 0, external: 0 };
    agg[t].internal += r['__budget_internal'] || 0;
    agg[t].external += r['__budget_external'] || 0;
  });

  const hasBreakdown = data.some(r => r['__budget_internal'] > 0 || r['__budget_external'] > 0);

  if (!hasBreakdown) {
    // Fallback: count per type when no internal/external columns exist
    const sorted = Object.entries(groupBy(data, col)).sort((a, b) => b[1] - a[1]);
    state.charts['typeBarChart'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: sorted.map(e => e[0]),
        datasets: [{
          label: 'จำนวนโครงการ',
          data: sorted.map(e => e[1]),
          backgroundColor: sorted.map((_, i) => PALETTE[(i + 4) % PALETTE.length] + 'bb'),
          borderColor:     sorted.map((_, i) => PALETTE[(i + 4) % PALETTE.length]),
          borderWidth: 1.5, borderRadius: 6,
        }],
      },
      options: barOpts('จำนวนโครงการ', false),
    });
    return;
  }

  const sorted = Object.entries(agg)
    .sort((a, b) => (b[1].internal + b[1].external) - (a[1].internal + a[1].external));
  const labels = sorted.map(e => e[0]);

  state.charts['typeBarChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'งบประมาณภายใน',
          data: sorted.map(e => e[1].internal),
          backgroundColor: '#3b82f6bb',
          borderColor: '#2563eb',
          borderWidth: 1.5,
          borderRadius: 4,
        },
        {
          label: 'งบประมาณภายนอก',
          data: sorted.map(e => e[1].external),
          backgroundColor: '#f59e0bbb',
          borderColor: '#d97706',
          borderWidth: 1.5,
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 700, easing: 'easeOutQuart' },
      plugins: {
        legend: {
          display: true,
          position: 'top',
          labels: { font: { family: 'Kanit', size: 11 }, boxWidth: 14, padding: 12 },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${fmtBudget(ctx.parsed.y)}`,
          },
          bodyFont: { family: 'Kanit' },
          titleFont: { family: 'Kanit' },
        },
      },
      scales: {
        x: {
          ticks: { font: { family: 'Kanit', size: 10 }, maxRotation: 35 },
          grid: { color: '#f3f4f6' },
        },
        y: {
          beginAtZero: true,
          ticks: {
            font: { family: 'Kanit', size: 10 },
            callback: function(v) { return fmtBudget(v); },
          },
          grid: { color: '#f3f4f6' },
        },
      },
    },
  });
}

// Shared bar chart option factory
function barOpts(yLabel, isMoney) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 700, easing: 'easeOutQuart' },
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: (ctx) => {
            const v = ctx.parsed.x !== undefined ? ctx.parsed.x : ctx.parsed.y;
            return ` ${yLabel}: ${isMoney ? fmtBudget(v) : v}`;
          },
        },
        bodyFont: { family: 'Kanit' },
        titleFont: { family: 'Kanit' },
      },
    },
    scales: {
      x: {
        ticks: { font: { family: 'Kanit', size: 10 }, maxRotation: 35 },
        grid: { color: '#f3f4f6' },
      },
      y: {
        beginAtZero: true,
        ticks: {
          font: { family: 'Kanit', size: 10 },
          callback: function(v) { return isMoney ? fmtBudget(v) : v; },
        },
        grid: { color: '#f3f4f6' },
      },
    },
  };
}

// ═══════════════════════════════════════════════
//  MODULE 6 — renderTable
//  Paginated, sortable table with all rows
// ═══════════════════════════════════════════════
function renderTable() {
  const data  = getSortedData();
  const start = (state.page - 1) * state.pageSize;
  const paged = data.slice(start, start + state.pageSize);

  renderTableHead();
  renderTableBody(paged, start);
  renderPagination(data.length);
}

function renderTableHead() {
  const thead = $('tableHead');
  thead.innerHTML = '';
  const tr = document.createElement('tr');

  // Row number column
  const thN = document.createElement('th');
  thN.textContent = '#';
  thN.className   = 'th-num';
  tr.appendChild(thN);

  state.columns.filter(c => !c.startsWith('__')).forEach(col => {
    const th     = document.createElement('th');
    const sorted = state.sort.col === col;
    th.className = `th-col${sorted ? (state.sort.dir === 'asc' ? ' sort-asc' : ' sort-desc') : ''}`;
    th.innerHTML = `<span class="th-text" title="${col}">${col}</span><span class="sort-icon">▲</span>`;
    th.addEventListener('click', () => handleSort(col));
    tr.appendChild(th);
  });

  thead.appendChild(tr);
}

function renderTableBody(paged, startIdx) {
  const tbody = $('tableBody');
  const { budget: bc, budgetInternal: bic, budgetExternal: bec } = state.colMap;
  const budgetCols = new Set([bc, bic, bec].filter(Boolean));

  tbody.innerHTML = '';

  if (!paged.length) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row';
    tr.innerHTML = `<td colspan="${state.columns.length + 1}">ไม่พบข้อมูลที่ตรงกับเงื่อนไข</td>`;
    tbody.appendChild(tr);
    return;
  }

  paged.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.style.animationDelay = `${i * 12}ms`;

    // Row number
    const tdN = document.createElement('td');
    tdN.className   = 'td-num';
    tdN.textContent = startIdx + i + 1;
    tr.appendChild(tdN);

    state.columns.filter(c => !c.startsWith('__')).forEach(col => {
      const td  = document.createElement('td');
      const val = row[col];

      if (budgetCols.has(col)) {
        // Show budget column with formatted badge
        const raw = col === bc  ? row['__budget_num']
                  : col === bic ? row['__budget_internal']
                  :               row['__budget_external'];
        td.innerHTML = raw > 0
          ? `<span class="budget-badge">${fmtBudget(raw)}</span>`
          : `<span class="text-gray-300 text-xs">–</span>`;
      } else {
        const display = (val === '' || val === null || val === undefined) ? '–' : val;
        td.textContent = display;
        if (String(val).length > 0) td.title = String(val);
      }
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

function getSortedData() {
  if (!state.sort.col) return state.filteredData;
  return [...state.filteredData].sort((a, b) => {
    let va = a[state.sort.col], vb = b[state.sort.col];
    const na = parseFloat(String(va).replace(/,/g, ''));
    const nb = parseFloat(String(vb).replace(/,/g, ''));
    if (!isNaN(na) && !isNaN(nb)) { va = na; vb = nb; }
    if (va < vb) return state.sort.dir === 'asc' ? -1 : 1;
    if (va > vb) return state.sort.dir === 'asc' ?  1 : -1;
    return 0;
  });
}

function handleSort(col) {
  state.sort.dir = (state.sort.col === col && state.sort.dir === 'asc') ? 'desc' : 'asc';
  state.sort.col = col;
  renderTable();
}

function renderPagination(total) {
  const ps         = state.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / ps));
  const cur        = state.page;
  const start      = (cur - 1) * ps + 1;
  const end        = Math.min(cur * ps, total);

  $('paginationInfo').textContent = total
    ? `แสดง ${start.toLocaleString('th-TH')}–${end.toLocaleString('th-TH')} จาก ${total.toLocaleString('th-TH')} รายการ`
    : 'ไม่พบข้อมูล';

  const ctrl = $('paginationControls');
  ctrl.innerHTML = '';

  const mkBtn = (lbl, pg, disabled = false, active = false) => {
    const btn     = document.createElement('button');
    btn.className = `page-btn${active ? ' active' : ''}`;
    btn.innerHTML = lbl;
    btn.disabled  = disabled;
    if (!disabled && !active) btn.onclick = () => { state.page = pg; renderTable(); };
    return btn;
  };

  ctrl.appendChild(mkBtn('«', 1,           cur === 1));
  ctrl.appendChild(mkBtn('‹', cur - 1,     cur === 1));

  const w = 2;
  for (let p = Math.max(1, cur - w); p <= Math.min(totalPages, cur + w); p++)
    ctrl.appendChild(mkBtn(p, p, false, p === cur));

  ctrl.appendChild(mkBtn('›', cur + 1,     cur === totalPages));
  ctrl.appendChild(mkBtn('»', totalPages,  cur === totalPages));
}

// ─────────────────────────────────────────────
//  Export CSV
// ─────────────────────────────────────────────
function exportCSV() {
  const data = getSortedData();
  if (!data.length) { alert('ไม่มีข้อมูลสำหรับส่งออก'); return; }

  const cols = state.columns.filter(c => !c.startsWith('__'));
  const rows = [
    cols,
    ...data.map(r => cols.map(c => `"${String(r[c] ?? '').replace(/"/g, '""')}"`)),
  ];

  const csv = '﻿' + rows.map(r => r.join(',')).join('\r\n');
  const a   = document.createElement('a');
  a.href    = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  a.download = `research_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ─────────────────────────────────────────────
//  Event listeners + drag-and-drop
// ─────────────────────────────────────────────
function initListeners() {
  $('fileInput').addEventListener('change', e => {
    if (e.target.files[0]) parseExcel(e.target.files[0]);
    e.target.value = '';
  });

  $('exportBtn').addEventListener('click', exportCSV);

  $('searchInput').addEventListener('input',  debounce(applyFilters, 300));
  $('deptFilter').addEventListener('change',  applyFilters);
  $('typeFilter').addEventListener('change',  applyFilters);
  $('budgetFilter').addEventListener('change', applyFilters);

  $('resetFilters').addEventListener('click', () => {
    ['searchInput', 'deptFilter', 'typeFilter', 'budgetFilter'].forEach(id => $(id).value = '');
    applyFilters();
  });

  $('pageSizeSelect').addEventListener('change', e => {
    state.pageSize = parseInt(e.target.value, 10);
    state.page     = 1;
    renderTable();
  });

  // Drag-and-drop on upload zone
  const zone = $('dropZone');
  if (zone) {
    zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drop-active'); });
    zone.addEventListener('dragenter', e => { e.preventDefault(); zone.classList.add('drop-active'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drop-active'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drop-active');
      const file = e.dataTransfer?.files?.[0];
      if (file && /\.(xlsx|xls|csv)$/i.test(file.name)) {
        parseExcel(file);
      } else if (file) {
        alert('กรุณาอัปโหลดไฟล์ .xlsx หรือ .xls เท่านั้น');
      }
    });
    // Also allow clicking the zone to open file picker
    zone.addEventListener('click', e => {
      if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
        $('fileInput').click();
      }
    });
  }
}

// ─────────────────────────────────────────────
//  Bootstrap
// ─────────────────────────────────────────────
(function init() {
  $('footerYear').textContent = new Date().getFullYear();
  initListeners();
  hideLoading();
  showEl('uploadPrompt');
})();
