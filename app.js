/* ============================================================
   Research Dashboard — app.js
   Dependencies: SheetJS (xlsx), Chart.js, TailwindCSS
   ============================================================ */

'use strict';

// ── State ──────────────────────────────────────────────────
const state = {
  rawData:      [],   // all rows from Excel
  filteredData: [],   // rows after filters applied
  columns:      [],   // detected column headers
  colMap:       {},   // { semantic: actualColName }
  charts:       {},   // chart instances keyed by id
  sort:         { col: null, dir: 'asc' },
  page:         1,
  pageSize:     10,
};

// Semantic column mapping — tries to find these in the sheet
const SEMANTIC_KEYS = {
  projectName:     ['ชื่อโครงการ', 'ชื่อโครงการวิจัย', 'project name', 'โครงการ'],
  dept:            ['หน่วยงาน', 'ภาควิชา', 'คณะ', 'สาขา', 'department', 'dept', 'faculty'],
  researcher:      ['หัวหน้าโครงการ', 'ชื่อหัวหน้า', 'นักวิจัย', 'ผู้วิจัย', 'researcher', 'pi'],
  budget:          ['งบประมาณรวม', 'รวมงบประมาณ', 'งบประมาณ', 'budget', 'เงินทุน', 'ทุนวิจัย', 'วงเงิน'],
  budgetInternal:  ['งบประมาณภายใน', 'งบภายใน', 'เงินทุนภายใน', 'ภายใน'],
  budgetExternal:  ['งบประมาณภายนอก', 'งบภายนอก', 'เงินทุนภายนอก', 'ภายนอก'],
  type:            ['ประเภท', 'ประเภทวิจัย', 'type', 'ประเภทโครงการ'],
  year:            ['ปี', 'ปีงบประมาณ', 'year'],
  status:          ['สถานะ', 'status'],
};

// Color palette
const PALETTE = [
  '#f59e0b','#10b981','#3b82f6','#8b5cf6','#ef4444',
  '#06b6d4','#f97316','#84cc16','#ec4899','#6366f1',
  '#14b8a6','#f43f5e','#a78bfa','#34d399','#fbbf24',
];

// ── Utility helpers ─────────────────────────────────────────
const $ = (id) => document.getElementById(id);
const fmt = (n) => {
  if (n === null || n === undefined || n === '') return '–';
  const num = parseFloat(String(n).replace(/,/g, ''));
  if (isNaN(num)) return n;
  return num.toLocaleString('th-TH');
};
const fmtBudget = (n) => {
  const num = parseFloat(String(n).replace(/,/g, ''));
  if (isNaN(num)) return '–';
  if (num >= 1e6)  return (num / 1e6).toFixed(2) + ' ล้าน';
  if (num >= 1e3)  return (num / 1e3).toFixed(1) + ' พัน';
  return num.toLocaleString('th-TH');
};
const showEl  = (id) => { const el = $(id); if (el) el.classList.remove('hidden'); };
const hideEl  = (id) => { const el = $(id); if (el) el.classList.add('hidden'); };
const uniqueVals = (arr, key) =>
  [...new Set(arr.map(r => r[key]).filter(Boolean))].sort((a, b) =>
    String(a).localeCompare(String(b), 'th'));

// ── Loading helpers ─────────────────────────────────────────
function showLoading(msg = 'กำลังโหลดข้อมูล...') {
  $('loadingMsg').textContent = msg;
  showEl('loadingOverlay');
}
function hideLoading() { hideEl('loadingOverlay'); }

// ── Column detection ────────────────────────────────────────
function detectColumns(headers) {
  const map = {};
  for (const [sem, candidates] of Object.entries(SEMANTIC_KEYS)) {
    for (const h of headers) {
      if (candidates.some(c => h.toLowerCase().includes(c.toLowerCase()))) {
        map[sem] = h;
        break;
      }
    }
  }
  return map;
}

// ── Excel / CSV loading ─────────────────────────────────────
// Layout: row 1-2 = title/meta, row 3 = headers, row 4 = sub-header/empty, row 5+ = data
const TARGET_SHEET   = 'โครงการวิจัย';
const HEADER_ROW     = 3;  // 1-based Excel row for column headers
const DATA_ROW_START = 5;  // 1-based Excel row where actual data begins

function loadData(file) {
  showLoading('กำลังอ่านไฟล์ Excel...');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type: 'array', cellDates: true });

      // Must use sheet "โครงการวิจัย" — no fallback
      if (!wb.SheetNames.includes(TARGET_SHEET)) {
        hideLoading();
        const found = wb.SheetNames.join(', ');
        alert(`ไม่พบ sheet "${TARGET_SHEET}" ในไฟล์นี้\nSheet ที่พบ: ${found}`);
        return;
      }

      $('loadingMsg').textContent = `อ่าน sheet: ${TARGET_SHEET}`;

      const ws = wb.Sheets[TARGET_SHEET];

      // Read all rows as raw arrays (no header inference)
      const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Row 3 (index 2) = headers; skip empty header cells
      const rawHeaders = allRows[HEADER_ROW - 1] || [];
      const headers = rawHeaders.map((h, i) =>
        (h !== null && h !== undefined && String(h).trim() !== '')
          ? String(h).trim()
          : `คอลัมน์_${i + 1}`
      );

      // Rows from row 4 (index 3) onward = data
      // Skip rows that are entirely empty
      const dataRows = allRows.slice(DATA_ROW_START - 1).filter(row =>
        row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
      );

      if (!dataRows.length) {
        hideLoading();
        alert(`ไม่พบข้อมูลใน sheet "${TARGET_SHEET}" (ตั้งแต่แถวที่ ${DATA_ROW_START})`);
        return;
      }

      // Build array of objects using detected headers
      const json = dataRows.map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
        return obj;
      });

      state.rawData  = json;
      state.columns  = headers;
      state.colMap   = detectColumns(headers);
      state.page     = 1;

      $('lastUpdated').textContent = `อัปเดต: ${new Date().toLocaleString('th-TH')}`;
      showEl('lastUpdated');

      processData();
    } catch (err) {
      hideLoading();
      console.error(err);
      alert('เกิดข้อผิดพลาดในการอ่านไฟล์: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Process & render ────────────────────────────────────────
function toNum(val) {
  const n = parseFloat(String(val ?? '').replace(/,/g, '').trim());
  return isNaN(n) ? 0 : n;
}

function processData() {
  showLoading('กำลังประมวลผลข้อมูล...');
  setTimeout(() => {
    const bc  = state.colMap.budget;
    const bic = state.colMap.budgetInternal;
    const bec = state.colMap.budgetExternal;

    state.rawData.forEach(row => {
      const internal = bic ? toNum(row[bic]) : 0;
      const external = bec ? toNum(row[bec]) : 0;

      // Total budget: prefer explicit total column, else sum internal+external
      if (bc) {
        row['__budget_num'] = toNum(row[bc]);
      } else if (bic || bec) {
        row['__budget_num'] = internal + external;
      } else {
        row['__budget_num'] = 0;
      }

      row['__budget_internal'] = internal;
      row['__budget_external'] = external;
    });

    populateFilters();
    applyFilters();
    hideEl('uploadPrompt');
    showEl('dashboard');
    hideLoading();
  }, 50);
}

// ── Populate filter dropdowns ───────────────────────────────
function populateFilters() {
  const deptCol = state.colMap.dept;
  const typeCol = state.colMap.type;

  const deptSel = $('deptFilter');
  const typeSel = $('typeFilter');

  deptSel.innerHTML = '<option value="">ทุกหน่วยงาน</option>';
  typeSel.innerHTML = '<option value="">ทุกประเภทวิจัย</option>';

  if (deptCol) {
    uniqueVals(state.rawData, deptCol).forEach(v => {
      deptSel.insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`);
    });
  }
  if (typeCol) {
    uniqueVals(state.rawData, typeCol).forEach(v => {
      typeSel.insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`);
    });
  }
}

// ── Apply filters ───────────────────────────────────────────
function applyFilters() {
  const search  = $('searchInput').value.toLowerCase().trim();
  const dept    = $('deptFilter').value;
  const type    = $('typeFilter').value;
  const budget  = $('budgetFilter').value;
  const deptCol = state.colMap.dept;
  const typeCol = state.colMap.type;

  state.filteredData = state.rawData.filter(row => {
    // Keyword search across all columns
    if (search) {
      const combined = Object.values(row).join(' ').toLowerCase();
      if (!combined.includes(search)) return false;
    }
    // Dept filter
    if (dept && deptCol && row[deptCol] !== dept) return false;
    // Type filter
    if (type && typeCol && row[typeCol] !== type) return false;
    // Budget range
    if (budget) {
      const [lo, hi] = budget.split('-').map(Number);
      const b = row['__budget_num'];
      if (b < lo || b > hi) return false;
    }
    return true;
  });

  state.page = 1;
  $('resultsCount').textContent = `พบ ${state.filteredData.length} รายการ`;
  renderSummaryCards();
  renderCharts();
  renderTable();
}

// ── Summary Cards ───────────────────────────────────────────
function renderSummaryCards() {
  const data = state.filteredData;
  const deptCol       = state.colMap.dept;
  const researcherCol = state.colMap.researcher;

  $('totalProjects').textContent    = data.length.toLocaleString('th-TH');
  $('totalBudget').textContent      = fmtBudget(data.reduce((s, r) => s + r['__budget_num'], 0));
  $('totalDepts').textContent       = deptCol
    ? uniqueVals(data, deptCol).length.toLocaleString('th-TH') : '–';
  $('totalResearchers').textContent = researcherCol
    ? uniqueVals(data, researcherCol).length.toLocaleString('th-TH') : '–';
}

// ── Chart helpers ───────────────────────────────────────────
function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
    delete state.charts[id];
  }
}

function groupBy(data, key) {
  return data.reduce((acc, row) => {
    const k = row[key] || '(ไม่ระบุ)';
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {});
}
function groupBudgetBy(data, key) {
  return data.reduce((acc, row) => {
    const k = row[key] || '(ไม่ระบุ)';
    acc[k] = (acc[k] || 0) + row['__budget_num'];
    return acc;
  }, {});
}

// ── Render all charts ───────────────────────────────────────
function renderCharts() {
  const data           = state.filteredData;
  const deptCol        = state.colMap.dept;
  const typeCol        = state.colMap.type;
  const researcherCol  = state.colMap.researcher;

  renderDeptBar(data, deptCol);
  renderBudgetDoughnut(data, deptCol);
  // Y-axis = researcher name (ชื่อหัวหน้าโครงการ), fallback to dept
  renderBudgetHBar(data, researcherCol || deptCol);
  renderTypeBar(data, typeCol);
}

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
        backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length] + 'cc'),
        borderColor:     labels.map((_, i) => PALETTE[i % PALETTE.length]),
        borderWidth: 1.5,
        borderRadius: 6,
      }],
    },
    options: chartOptions('จำนวนโครงการ', false),
  });
}

function renderBudgetDoughnut(data, col) {
  destroyChart('budgetDoughnut');
  const ctx = $('budgetDoughnut').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'budgetDoughnut'); return; }

  const grouped  = groupBudgetBy(data, col);
  const sorted   = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const labels   = sorted.map(e => e[0]);
  const values   = sorted.map(e => e[1]);
  const total    = values.reduce((s, v) => s + v, 0);

  state.charts['budgetDoughnut'] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length] + 'dd'),
        borderColor: '#ffffff',
        borderWidth: 2,
        hoverOffset: 10,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { animateRotate: true, duration: 800 },
      plugins: {
        legend: {
          position: 'right',
          labels: { font: { family: 'Kanit', size: 11 }, boxWidth: 12, padding: 10,
            generateLabels: (chart) => {
              const ds = chart.data.datasets[0];
              return chart.data.labels.map((lbl, i) => {
                const pct = total > 0 ? ((ds.data[i] / total) * 100).toFixed(1) : 0;
                const short = lbl.length > 16 ? lbl.slice(0, 15) + '…' : lbl;
                return {
                  text: `${short} (${pct}%)`,
                  fillStyle: ds.backgroundColor[i],
                  hidden: false,
                  index: i,
                };
              });
            },
          },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed.toLocaleString('th-TH')} บาท (${pct}%)`;
            },
          },
          bodyFont: { family: 'Kanit' },
          titleFont: { family: 'Kanit' },
        },
      },
    },
  });
}

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
        label: 'งบประมาณ (บาท)',
        data: values,
        backgroundColor: '#f59e0b99',
        borderColor: '#d97706',
        borderWidth: 1.5,
        borderRadius: 4,
      }],
    },
    options: {
      ...chartOptions('งบประมาณ (บาท)', true),
      indexAxis: 'y',
    },
  });
}

function renderTypeBar(data, col) {
  destroyChart('typeBarChart');
  const ctx = $('typeBarChart').getContext('2d');
  if (!col || !data.length) { renderEmpty(ctx, 'typeBarChart'); return; }

  // Aggregate internal & external budget per type
  const agg = {};
  data.forEach(row => {
    const t = row[col] || '(ไม่ระบุ)';
    if (!agg[t]) agg[t] = { internal: 0, external: 0 };
    agg[t].internal += row['__budget_internal'] || 0;
    agg[t].external += row['__budget_external'] || 0;
  });

  const hasBudgetBreakdown = data.some(r => r['__budget_internal'] > 0 || r['__budget_external'] > 0);

  if (!hasBudgetBreakdown) {
    // Fallback: simple project-count bar when no internal/external columns found
    const sorted = Object.entries(groupBy(data, col)).sort((a, b) => b[1] - a[1]);
    state.charts['typeBarChart'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: sorted.map(e => e[0]),
        datasets: [{
          label: 'จำนวนโครงการ',
          data: sorted.map(e => e[1]),
          backgroundColor: sorted.map((_, i) => PALETTE[(i + 4) % PALETTE.length] + 'cc'),
          borderColor:     sorted.map((_, i) => PALETTE[(i + 4) % PALETTE.length]),
          borderWidth: 1.5, borderRadius: 6,
        }],
      },
      options: chartOptions('จำนวนโครงการ', false),
    });
    return;
  }

  // Grouped bar: internal vs external budget per research type
  const sorted  = Object.entries(agg).sort((a, b) => (b[1].internal + b[1].external) - (a[1].internal + a[1].external));
  const labels  = sorted.map(e => e[0]);
  const inVals  = sorted.map(e => e[1].internal);
  const exVals  = sorted.map(e => e[1].external);

  state.charts['typeBarChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'งบประมาณภายใน',
          data: inVals,
          backgroundColor: '#3b82f6cc',
          borderColor: '#2563eb',
          borderWidth: 1.5,
          borderRadius: 4,
        },
        {
          label: 'งบประมาณภายนอก',
          data: exVals,
          backgroundColor: '#f59e0bcc',
          borderColor: '#d97706',
          borderWidth: 1.5,
          borderRadius: 4,
        },
      ],
    },
    options: {
      ...chartOptions('งบประมาณ (บาท)', true),
      plugins: {
        ...chartOptions('งบประมาณ (บาท)', true).plugins,
        legend: {
          display: true,
          position: 'top',
          labels: { font: { family: 'Kanit', size: 11 }, boxWidth: 14, padding: 12 },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('th-TH')} บาท`,
          },
          bodyFont: { family: 'Kanit' },
          titleFont: { family: 'Kanit' },
        },
      },
    },
  });
}

function renderEmpty(ctx, id) {
  const canvas = $(id);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = '14px Kanit, sans-serif';
  ctx.fillStyle = '#9ca3af';
  ctx.textAlign = 'center';
  ctx.fillText('ไม่พบข้อมูลที่เกี่ยวข้อง', canvas.width / 2, canvas.height / 2);
}

function chartOptions(yLabel, isMoney) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 600, easing: 'easeOutQuart' },
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: (ctx) => {
            const v = ctx.parsed.x !== undefined ? ctx.parsed.x : ctx.parsed.y;
            return ` ${yLabel}: ${isMoney ? v.toLocaleString('th-TH') + ' บาท' : v}`;
          },
        },
        bodyFont: { family: 'Kanit' },
        titleFont: { family: 'Kanit' },
      },
    },
    scales: {
      x: {
        ticks: {
          font: { family: 'Kanit', size: 10 },
          maxRotation: 35,
          callback: (v, i, ticks) => {
            const lbl = typeof v === 'number'
              ? (isMoney ? fmtBudget(v) : v)
              : v;
            return typeof lbl === 'string' && lbl.length > 14 ? lbl.slice(0, 13) + '…' : lbl;
          },
        },
        grid: { color: '#f3f4f6' },
      },
      y: {
        ticks: {
          font: { family: 'Kanit', size: 10 },
          callback: (v) => {
            const lbl = isMoney ? fmtBudget(v) : v;
            return typeof lbl === 'string' && lbl.length > 18 ? lbl.slice(0, 17) + '…' : lbl;
          },
        },
        grid: { color: '#f3f4f6' },
        beginAtZero: true,
      },
    },
  };
}

// ── Table ───────────────────────────────────────────────────
function renderTable() {
  const data   = getSortedData();
  const ps     = state.pageSize;
  const page   = state.page;
  const start  = (page - 1) * ps;
  const end    = start + ps;
  const paged  = data.slice(start, end);

  renderTableHead();
  renderTableBody(paged, start);
  renderPagination(data.length);
}

function renderTableHead() {
  const cols = state.columns;
  const thead = $('tableHead');
  thead.innerHTML = '';
  const tr = document.createElement('tr');

  // Row number column
  const thNum = document.createElement('th');
  thNum.textContent = '#';
  thNum.className = 'w-10 text-center text-xs font-semibold text-gray-500 px-3 py-3';
  tr.appendChild(thNum);

  cols.filter(c => !c.startsWith('__')).forEach(col => {
    const th = document.createElement('th');
    const isSorted = state.sort.col === col;
    th.className = `px-4 py-3 text-left text-xs font-semibold text-gray-500 uppercase tracking-wide whitespace-nowrap cursor-pointer select-none hover:bg-gold-50 transition-colors ${isSorted ? (state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc') : ''}`;
    th.innerHTML = `${col} <span class="sort-icon">▲</span>`;
    th.dataset.col = col;
    th.addEventListener('click', () => handleSort(col));
    tr.appendChild(th);
  });

  thead.appendChild(tr);
}

function renderTableBody(paged, startIdx) {
  const tbody   = $('tableBody');
  const budgetC = state.colMap.budget;
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
    tr.style.animationDelay = `${i * 15}ms`;

    // Row number
    const tdNum = document.createElement('td');
    tdNum.textContent = startIdx + i + 1;
    tdNum.className = 'text-center text-xs text-gray-400 px-3';
    tr.appendChild(tdNum);

    state.columns.filter(c => !c.startsWith('__')).forEach(col => {
      const td = document.createElement('td');
      const val = row[col];

      if (col === budgetC && val !== '' && val !== undefined) {
        td.innerHTML = `<span class="budget-badge">${Number(row['__budget_num']).toLocaleString('th-TH')}</span>`;
      } else {
        td.title = String(val);
        td.textContent = val === '' || val === null || val === undefined ? '–' : val;
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
    if (va > vb) return state.sort.dir === 'asc' ? 1 : -1;
    return 0;
  });
}

function handleSort(col) {
  if (state.sort.col === col) {
    state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    state.sort.col = col;
    state.sort.dir = 'asc';
  }
  renderTable();
}

// ── Pagination ──────────────────────────────────────────────
function renderPagination(total) {
  const ps        = state.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / ps));
  const cur       = state.page;
  const start     = (cur - 1) * ps + 1;
  const end       = Math.min(cur * ps, total);

  $('paginationInfo').textContent = total
    ? `แสดง ${start}–${end} จาก ${total.toLocaleString('th-TH')} รายการ`
    : 'ไม่พบข้อมูล';

  const ctrl = $('paginationControls');
  ctrl.innerHTML = '';

  const mkBtn = (label, page, disabled = false, active = false) => {
    const btn = document.createElement('button');
    btn.className = `page-btn${active ? ' active' : ''}`;
    btn.innerHTML = label;
    btn.disabled  = disabled;
    if (!disabled && !active) btn.onclick = () => { state.page = page; renderTable(); };
    return btn;
  };

  ctrl.appendChild(mkBtn('«', 1, cur === 1));
  ctrl.appendChild(mkBtn('‹', cur - 1, cur === 1));

  // Page number window
  const window = 2;
  for (let p = Math.max(1, cur - window); p <= Math.min(totalPages, cur + window); p++) {
    ctrl.appendChild(mkBtn(p, p, false, p === cur));
  }

  ctrl.appendChild(mkBtn('›', cur + 1, cur === totalPages));
  ctrl.appendChild(mkBtn('»', totalPages, cur === totalPages));
}

// ── Export CSV ──────────────────────────────────────────────
function exportCSV() {
  const data = getSortedData();
  if (!data.length) { alert('ไม่มีข้อมูลสำหรับส่งออก'); return; }

  const cols = state.columns.filter(c => !c.startsWith('__'));
  const rows = [cols, ...data.map(r => cols.map(c => {
    const v = String(r[c] ?? '').replace(/"/g, '""');
    return `"${v}"`;
  }))];

  const csv = '﻿' + rows.map(r => r.join(',')).join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `research_projects_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Event Listeners ─────────────────────────────────────────
function initListeners() {
  $('fileInput').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) loadData(file);
    e.target.value = '';
  });

  $('exportBtn').addEventListener('click', exportCSV);

  $('searchInput').addEventListener('input', debounce(applyFilters, 300));
  $('deptFilter').addEventListener('change', applyFilters);
  $('typeFilter').addEventListener('change', applyFilters);
  $('budgetFilter').addEventListener('change', applyFilters);

  $('resetFilters').addEventListener('click', () => {
    $('searchInput').value = '';
    $('deptFilter').value  = '';
    $('typeFilter').value  = '';
    $('budgetFilter').value = '';
    applyFilters();
  });

  $('pageSizeSelect').addEventListener('change', (e) => {
    state.pageSize = parseInt(e.target.value);
    state.page = 1;
    renderTable();
  });
}

function debounce(fn, delay) {
  let timer;
  return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); };
}

// ── Bootstrap ───────────────────────────────────────────────
(function init() {
  $('footerYear').textContent = new Date().getFullYear();
  initListeners();

  // Hide loading, show upload prompt
  hideLoading();
  showEl('uploadPrompt');
})();
