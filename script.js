'use strict';

const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
const REPORT_WIDTH = 1600; // px width of .rpt-wrap
const nf = new Intl.NumberFormat('en-IN');
const fmt = n => nf.format(n);
const $ = id => document.getElementById(id);
const REDUCED_MOTION = !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);

let workbook = null;
let currentHeaders = [];
let currentRows = [];
let computed = null; // {deptList, offByDept, grand, reportDateStr, reportDateObj, totalRows, unparsedDates, futureDates, statusOrder, deptStatusAgg}
let mapSelections = {};

const dropZone = $('dropZone');
const fileInput = $('fileInput');
const fileNameEl = $('fileName');
const uploadMsg = $('uploadMsg');
const mapCard = $('mapCard');
const mapTable = $('mapTable');
const sheetSelect = $('sheetSelect');
const reportDateInput = $('reportDateInput');
const mapMsg = $('mapMsg');
const genMsg = $('genMsg');

/* ==========================================================================
   Helpers
   ========================================================================== */

function showMsg(el, type, text) {
  if (!el) return;
  el.className = 'msg ' + type;
  el.textContent = text;
}

function clearMsg(el) {
  if (!el) return;
  el.className = 'msg';
  el.textContent = '';
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function safeStorageGet(key) {
  try { return localStorage.getItem(key); } catch (e) { return null; }
}

function safeStorageSet(key, value) {
  try { localStorage.setItem(key, value); } catch (e) { /* storage unavailable */ }
}

function localDateInputValue(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

/* Nav links to cards that are still hidden are disabled */
function refreshNav() {
  document.querySelectorAll('.nav-link[href^="#"]').forEach(a => {
    const href = a.getAttribute('href') || '';
    if (href.charAt(0) !== '#') return; // normal page links (e.g. pricing.html) are always enabled
    const target = document.querySelector(href);
    const off = !target || target.classList.contains('hidden');
    a.classList.toggle('is-disabled', off);
    a.setAttribute('aria-disabled', off ? 'true' : 'false');
    a.tabIndex = off ? -1 : 0;
  });
}

/* Hide previous results so stale numbers are never shown after a new upload / sheet change */
function resetResults() {
  computed = null;
  ['summaryCard', 'dashPreviewCard', 'offPreviewCard'].forEach(id => $(id).classList.add('hidden'));
  $('dashboardRender').innerHTML = '';
  $('officerRender').innerHTML = '';
  clearMsg(genMsg);
  refreshNav();
}

/* ==========================================================================
   File upload
   ========================================================================== */

function isExcelFile(file) {
  return !!file && /\.(xlsx|xls)$/i.test(file.name);
}

// Allow picking the same file again (otherwise "change" does not fire)
fileInput.addEventListener('click', () => { fileInput.value = ''; });

fileInput.addEventListener('change', () => {
  if (fileInput.files && fileInput.files.length) handleFile(fileInput.files[0]);
});

['dragenter', 'dragover'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.add('dragging');
  });
});

['dragleave', 'dragend', 'drop'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.remove('dragging');
  });
});

dropZone.addEventListener('drop', e => {
  const files = e.dataTransfer && e.dataTransfer.files;
  if (files && files.length) handleFile(files[0]);
});

// Stop the browser from opening files dropped outside the drop zone
['dragover', 'drop'].forEach(evt => {
  window.addEventListener(evt, e => { e.preventDefault(); }, false);
});

function handleFile(file) {
  if (!isExcelFile(file)) {
    showMsg(uploadMsg, 'error', 'Unsupported file type. Please choose an .xlsx or .xls file.');
    return;
  }

  workbook = null;
  currentHeaders = [];
  currentRows = [];
  resetResults();
  mapCard.classList.add('hidden');
  clearMsg(mapMsg);
  refreshNav();

  fileNameEl.textContent = '📁 ' + file.name;
  dropZone.classList.remove('has-file');
  showMsg(uploadMsg, 'info', 'Reading file…');

  const reader = new FileReader();
  reader.onerror = function () {
    showMsg(uploadMsg, 'error', 'Could not read this file. Please try again.');
  };
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array', cellDates: true });
      if (!workbook.SheetNames.length) throw new Error('The workbook has no sheets.');

      sheetSelect.innerHTML = '';
      workbook.SheetNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name; opt.textContent = name;
        sheetSelect.appendChild(opt);
      });
      const preferred = workbook.SheetNames.find(n => /complaint/i.test(n)) || workbook.SheetNames[0];
      sheetSelect.value = preferred;

      mapCard.classList.remove('hidden');
      refreshNav();
      loadSheet(preferred);

      dropZone.classList.add('has-file');
      showMsg(uploadMsg, 'ok', 'File loaded successfully: ' + file.name);
      mapCard.scrollIntoView({ behavior: REDUCED_MOTION ? 'auto' : 'smooth', block: 'start' });
    } catch (err) {
      showMsg(uploadMsg, 'error', 'Could not read this file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener('change', () => {
  resetResults();
  loadSheet(sheetSelect.value);
});

function loadSheet(name) {
  const ws = workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  if (!json.length) {
    currentHeaders = [];
    currentRows = [];
    mapTable.innerHTML = '';
    showMsg(mapMsg, 'error', 'The selected sheet has no data rows. Pick another sheet.');
    return false;
  }
  currentHeaders = Object.keys(json[0]);
  currentRows = json;
  buildMapTable();
  return true;
}

/* ==========================================================================
   Column mapping
   ========================================================================== */

function normKey(s) { return s.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ''); }

function guessColumn(candidates) {
  const norm = currentHeaders.map(normKey);
  for (const c of candidates) {
    const nc = normKey(c);
    if (!nc) continue;
    const idx = norm.indexOf(nc);
    if (idx >= 0) return currentHeaders[idx];
  }
  for (const c of candidates) {
    const nc = normKey(c);
    if (!nc) continue;
    // Guard against empty / very short headers matching everything
    const idx = norm.findIndex(h => h && (h.includes(nc) || (h.length >= 4 && nc.includes(h))));
    if (idx >= 0) return currentHeaders[idx];
  }
  return ''; // not detected
}

const FIELD_DEFS = [
  { key: 'date', label: 'Complaint Date column', required: true, candidates: ['Date', 'Complaint Date'] },
  { key: 'dept', label: 'Department column', required: true, candidates: ['Department'] },
  { key: 'officer', label: 'Currently Assigned To (Officer) column', required: true, candidates: ['Currently Assigned To'] },
  { key: 'esc', label: 'Is Escalated column', required: true, candidates: ['Is Escalated'] },
  { key: 'mobile', label: 'Assigned To Mobile column (pivot list export)', required: false, candidates: ['Assigned To Mobile', 'Mobile'] },
  { key: 'status', label: 'Current Status column (status pivot export)', required: false, candidates: ['Current Satus', 'Current Status', 'Status'] },
];

function missingRequired() {
  return FIELD_DEFS.filter(f => f.required && !mapSelections[f.key]);
}

function updateMapHint() {
  const missing = missingRequired();
  if (missing.length) {
    showMsg(mapMsg, 'error', 'Select a column for: ' + missing.map(f => f.label.replace(' column', '')).join(', ') + '.');
  } else {
    showMsg(mapMsg, 'info', 'Columns look good. Adjust if needed, verify the report date, then click Generate.');
  }
}

function buildMapTable() {
  mapTable.innerHTML = '';
  FIELD_DEFS.forEach(f => {
    const guess = guessColumn(f.candidates);
    mapSelections[f.key] = guess;

    const tr = document.createElement('tr');
    const td1 = document.createElement('td');
    td1.textContent = f.label;
    const tag = document.createElement('span');
    tag.className = 'req-tag';
    tag.textContent = f.required ? 'required' : 'optional';
    td1.appendChild(tag);

    const td2 = document.createElement('td');
    const sel = document.createElement('select');
    sel.setAttribute('aria-label', f.label);

    const none = document.createElement('option');
    none.value = '';
    none.textContent = f.required ? '— Not detected: choose a column —' : '— Skip (optional) —';
    sel.appendChild(none);

    currentHeaders.forEach(h => {
      const opt = document.createElement('option');
      opt.value = h; opt.textContent = h;
      sel.appendChild(opt);
    });
    sel.value = guess;

    const syncRow = () => tr.classList.toggle('row-missing', f.required && !sel.value);
    syncRow();

    sel.addEventListener('change', () => {
      mapSelections[f.key] = sel.value;
      syncRow();
      updateMapHint();
    });

    td2.appendChild(sel);
    tr.appendChild(td1); tr.appendChild(td2);
    mapTable.appendChild(tr);
  });
  updateMapHint();
}

/* ==========================================================================
   Parsing & aggregation
   ========================================================================== */

function parseDateVal(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return new Date(d.y, d.m - 1, d.d);
    } catch (e) { /* fall through */ }
  }
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    let m = s.match(/^(\d{1,2})[-\/]([A-Za-z]{3,})[-\/](\d{4})/);
    if (m) {
      const day = parseInt(m[1], 10);
      const mon = MONTHS[m[2].toLowerCase().slice(0, 3)];
      const year = parseInt(m[3], 10);
      if (mon !== undefined) return new Date(year, mon, day);
    }
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (m) {
      return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
    const d2 = new Date(s);
    if (!isNaN(d2)) return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  }
  return null;
}

function bucketOf(days) {
  if (days <= 15) return 0;
  if (days <= 30) return 1;
  if (days <= 45) return 2;
  if (days <= 90) return 3;
  return 4;
}

function fmtDateDDMMMYYYY(d) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return String(d.getDate()).padStart(2, '0') + '-' + months[d.getMonth()] + '-' + d.getFullYear();
}

function cellText(r, col) {
  if (!col) return '';
  const v = r[col];
  return (v === undefined || v === null) ? '' : v.toString().trim();
}

$('generateBtn').addEventListener('click', generateReport);

function generateReport() {
  try {
    if (!currentRows.length) { showMsg(mapMsg, 'error', 'No data loaded. Upload a file and pick a sheet with data.'); return; }
    const missing = missingRequired();
    if (missing.length) { updateMapHint(); return; }
    if (!reportDateInput.value) { showMsg(mapMsg, 'error', 'Please choose a report date.'); return; }

    const [ry, rm, rd] = reportDateInput.value.split('-').map(x => parseInt(x, 10));
    const reportDate = new Date(ry, rm - 1, rd);
    const reportDateStr = fmtDateDDMMMYYYY(reportDate);

    const { date: colDate, dept: colDept, officer: colOfficer, esc: colEsc, mobile: colMobile, status: colStatus } = mapSelections;

    const deptAgg = {};
    const offAgg = {};
    const statusOrder = [];
    const deptStatusAgg = {};
    let totalRows = 0, unparsedDates = 0, futureDates = 0;

    currentRows.forEach(r => {
      let dept = cellText(r, colDept);
      if (dept === '') dept = 'Unspecified Department';
      let officer = cellText(r, colOfficer);
      if (officer === '' || officer === '-') officer = 'Unassigned';
      let mobile = cellText(r, colMobile);
      if (mobile === '-') mobile = '';
      let status = cellText(r, colStatus);
      if (status === '') status = 'Unspecified';
      const escRaw = cellText(r, colEsc).toLowerCase();
      const isEsc = (escRaw === 'yes' || escRaw === 'y' || escRaw === 'true' || escRaw === '1');

      const dateVal = parseDateVal(r[colDate]);
      let age = 0;
      if (dateVal) {
        const diff = Math.round((reportDate - dateVal) / 86400000);
        if (diff < 0) { futureDates++; } else { age = diff; }
      } else {
        unparsedDates++;
      }
      const b = bucketOf(age);

      totalRows++;
      if (!deptAgg[dept]) deptAgg[dept] = { total: 0, escalated: 0, buckets: [0, 0, 0, 0, 0] };
      deptAgg[dept].total++;
      if (isEsc) deptAgg[dept].escalated++;
      // Dashboard aging buckets must reconcile to Total Pending, just like the officer report.
      deptAgg[dept].buckets[b]++;

      const key = dept + '||' + officer;
      if (!offAgg[key]) offAgg[key] = { dept, officer, total: 0, buckets: [0, 0, 0, 0, 0], mobile: '' };
      offAgg[key].total++;
      offAgg[key].buckets[b]++;
      if (!offAgg[key].mobile && mobile) offAgg[key].mobile = mobile;

      if (!statusOrder.includes(status)) statusOrder.push(status);
      if (!deptStatusAgg[dept]) deptStatusAgg[dept] = {};
      deptStatusAgg[dept][status] = (deptStatusAgg[dept][status] || 0) + 1;
    });

    const deptList = Object.keys(deptAgg).map(name => {
      const a = deptAgg[name];
      return { name, total: a.total, escalated: a.escalated, withinSla: a.total - a.escalated, buckets: a.buckets };
    });
    deptList.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));

    const offByDept = {};
    let officerRowSum = 0;
    Object.values(offAgg).forEach(o => {
      officerRowSum += o.total;
      if (!offByDept[o.dept]) offByDept[o.dept] = [];
      offByDept[o.dept].push(o);
    });
    Object.keys(offByDept).forEach(d => {
      offByDept[d].sort((a, b) => b.total - a.total || a.officer.localeCompare(b.officer));
    });

    const grand = deptList.reduce((acc, d) => {
      acc.total += d.total; acc.escalated += d.escalated; acc.withinSla += d.withinSla;
      for (let i = 0; i < 5; i++) acc.buckets[i] += d.buckets[i];
      return acc;
    }, { total: 0, escalated: 0, withinSla: 0, buckets: [0, 0, 0, 0, 0] });

    computed = {
      deptList, offByDept, grand, reportDateStr, reportDateObj: reportDate,
      totalRows, unparsedDates, futureDates, officerRowSum,
      statusOrder: statusOrder.sort((a, b) => a.localeCompare(b)), deptStatusAgg
    };

    renderSummary();
    renderDashboard();
    renderOfficerWise();

    ['summaryCard', 'dashPreviewCard', 'offPreviewCard'].forEach(id => $(id).classList.remove('hidden'));
    refreshNav();
    if (zoomers.dash) zoomers.dash.fit();
    if (zoomers.off) zoomers.off.fit();
    clearMsg(genMsg);
    $('summaryCard').scrollIntoView({ behavior: REDUCED_MOTION ? 'auto' : 'smooth', block: 'start' });
  } catch (err) {
    showMsg(mapMsg, 'error', 'Error while generating: ' + err.message);
    console.error(err);
  }
}

/* One ordering used by previews, PDFs and the Excel report exports */
function orderedDepts(deptList) {
  const sel = $('dashboardSortSelect');
  const dir = sel ? sel.value : 'desc';
  return [...deptList].sort((a, b) => {
    const byEsc = dir === 'asc' ? a.escalated - b.escalated : b.escalated - a.escalated;
    return byEsc || a.name.localeCompare(b.name);
  });
}

/* ==========================================================================
   Rendering
   ========================================================================== */

function animateValue(id, start, end, duration) {
  const obj = $(id);
  if (!obj) return;
  if (start === end || REDUCED_MOTION) { obj.textContent = fmt(end); return; }
  const range = end - start;
  const startTime = performance.now();

  function step(now) {
    const progress = Math.min((now - startTime) / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    obj.textContent = fmt(Math.floor(start + range * eased));
    if (progress < 1) requestAnimationFrame(step);
    else obj.textContent = fmt(end);
  }
  requestAnimationFrame(step);
}

function renderSummary() {
  const g = computed.grand;
  animateValue('sTotal', 0, g.total, 800);
  animateValue('sSla', 0, g.withinSla, 800);
  animateValue('sEsc', 0, g.escalated, 800);

  const slaRate = g.total > 0 ? (g.withinSla / g.total) * 100 : 0;
  const escRate = g.total > 0 ? (g.escalated / g.total) * 100 : 0;

  $('slaProgressBar').style.width = slaRate.toFixed(1) + '%';
  $('slaRateBadge').textContent = slaRate.toFixed(1) + '% SLA compliance';
  $('sSlaPct').textContent = slaRate.toFixed(1) + '% (' + fmt(g.withinSla) + ')';
  $('sEscPct').textContent = escRate.toFixed(1) + '% (' + fmt(g.escalated) + ')';

  const deptRowSum = computed.deptList.reduce((s, d) => s + d.total, 0);
  const checks = [
    {
      ok: deptRowSum === computed.totalRows,
      text: `Rows in file (${fmt(computed.totalRows)}) = sum of department totals (${fmt(deptRowSum)})`
    },
    {
      ok: computed.officerRowSum === computed.totalRows,
      text: `Rows in file (${fmt(computed.totalRows)}) = sum of officer totals (${fmt(computed.officerRowSum)})`
    },
    {
      ok: computed.unparsedDates === 0,
      text: computed.unparsedDates === 0
        ? 'All complaint dates were read successfully'
        : `${fmt(computed.unparsedDates)} row(s) have a blank or unreadable date and were aged as 0 days`
    },
    {
      ok: computed.futureDates === 0,
      text: computed.futureDates === 0
        ? 'No complaints are dated after the report date'
        : `${fmt(computed.futureDates)} row(s) are dated after the report date and were aged as 0 days`
    },
    {
      ok: !(g.total > 0 && g.escalated === 0),
      text: (g.total > 0 && g.escalated === 0)
        ? 'No rows are flagged as escalated — confirm the "Is Escalated" column mapping'
        : `${fmt(g.escalated)} row(s) flagged as escalated`
    }
  ];

  $('checklist').innerHTML = checks.map(c =>
    `<div class="chk-line"><span class="${c.ok ? 'ok' : 'bad'}">${c.ok ? 'OK' : 'CHECK'}</span><span>${escapeHtml(c.text)}</span></div>`
  ).join('');
}

function renderDashboard() {
  const { deptList, grand, reportDateStr } = computed;
  const rows = orderedDepts(deptList).map(d => `
    <tr class="${d.buckets[4] > 0 ? 'over-90-row' : ''}">
      <td class="dept-name-cell">${escapeHtml(d.name)}</td>
      <td class="lightblue">${d.total}</td>
      <td class="green">${d.withinSla}</td>
      <td class="red">${d.escalated}</td>
      <td>${d.buckets[0]}</td><td>${d.buckets[1]}</td><td>${d.buckets[2]}</td><td>${d.buckets[3]}</td><td class="${d.buckets[4] > 0 ? 'over-90-highlight' : ''}">${d.buckets[4]}</td>
    </tr>`).join('');

  $('dashboardRender').innerHTML = `
  <div class="rpt-wrap" id="dashboardCanvasTarget">
    <div class="rpt-title">311 APP-PENDING COMPLAINTS &ndash; MANAGEMENT DASHBOARD</div>
    <div class="rpt-asof">As on ${reportDateStr}</div>
    <table class="rpt-table">
      <tr><td colspan="3" class="section-bar">OVERALL</td></tr>
      <tr>
        <td class="lightblue ov-head">TOTAL PENDING</td>
        <td class="green ov-head">Within SLA</td>
        <td class="red ov-head">ESCALATED</td>
      </tr>
      <tr>
        <td class="lightblue ov-val">${grand.total}</td>
        <td class="green ov-val">${grand.withinSla} (OUT OF ${grand.total})</td>
        <td class="red ov-val">${grand.escalated} (OUT OF ${grand.total})</td>
      </tr>
    </table>
    <div class="rpt-gap"></div>
    <table class="rpt-table">
      <colgroup>
        <col style="width:19%;"><col style="width:9%;"><col style="width:9%;"><col style="width:9%;">
        <col style="width:10.8%;"><col style="width:10.8%;"><col style="width:10.8%;"><col style="width:10.8%;"><col style="width:10.8%;">
      </colgroup>
      <tr><td colspan="9" class="section-bar">DEPARTMENT-WISE PENDENCY</td></tr>
      <tr class="col-head">
        <td>Department</td><td>Total<br>Pending</td><td>Within<br>SLA</td><td>Escalated</td>
        <td>0&ndash;15<br>Days</td><td>16&ndash;30<br>Days</td><td>31&ndash;45<br>Days</td><td>46&ndash;90<br>Days</td><td>More than<br>90 Days</td>
      </tr>
      ${rows}
      <tr class="total-row">
        <td>TOTAL</td><td>${grand.total}</td><td>${grand.withinSla}</td><td>${grand.escalated}</td>
        <td>${grand.buckets[0]}</td><td>${grand.buckets[1]}</td><td>${grand.buckets[2]}</td><td>${grand.buckets[3]}</td><td>${grand.buckets[4]}</td>
      </tr>
    </table>
  </div>`;
  if (zoomers.dash) zoomers.dash.refresh();
}

$('dashboardSortSelect').addEventListener('change', () => {
  if (!computed) return;
  renderDashboard();
  renderOfficerWise();
});

function renderOfficerWise() {
  const { deptList, offByDept, reportDateStr } = computed;
  let blocks = '';
  orderedDepts(deptList).forEach(d => {
    const officers = offByDept[d.name] || [];
    const rows = officers.map((o, i) => `
      <tr>
        <td>${i + 1}</td>
        <td class="off-name-cell">${escapeHtml(o.officer)}</td>
        <td>${o.total}</td>
        <td>${o.buckets[0]}</td><td>${o.buckets[1]}</td><td>${o.buckets[2]}</td><td>${o.buckets[3]}</td><td>${o.buckets[4]}</td>
      </tr>`).join('');
    const totals = officers.reduce((acc, o) => { acc.total += o.total; for (let i = 0; i < 5; i++) acc.b[i] += o.buckets[i]; return acc; }, { total: 0, b: [0, 0, 0, 0, 0] });
    blocks += `
    <div class="off-block" data-dept-block="1">
      <div class="off-dept-bar">${escapeHtml(d.name)}</div>
      <table class="off-table">
        <tr class="off-col-head">
          <td style="width:6%;">S.No.</td><td style="width:38%;">Officer</td><td style="width:11%;">Total Pending</td>
          <td style="width:11%;">0&ndash;15 Days</td><td style="width:11%;">16&ndash;30 Days</td><td style="width:11%;">31&ndash;45 Days</td><td style="width:11%;">46&ndash;90 Days</td><td style="width:12%;">More than 90 Days</td>
        </tr>
        ${rows}
        <tr class="off-total-row">
          <td colspan="2">TOTAL</td><td>${totals.total}</td><td>${totals.b[0]}</td><td>${totals.b[1]}</td><td>${totals.b[2]}</td><td>${totals.b[3]}</td><td>${totals.b[4]}</td>
        </tr>
      </table>
    </div>`;
  });

  $('officerRender').innerHTML = `
  <div class="rpt-wrap" id="officerCanvasTarget">
    <div class="off-title" id="officerTitleBlock">OFFICER-WISE PENDENCY AS ON ${reportDateStr.toUpperCase()}</div>
    ${blocks}
  </div>`;
  if (zoomers.off) zoomers.off.refresh();
}

/* ==========================================================================
   Preview zoom (transform-based; sizer keeps scrollbars correct)
   ========================================================================== */

const zoomers = {};

function initPreviewControls() {
  [
    { key: 'dash', prefix: 'dash', targetId: 'dashboardRender', outerId: 'dashPreviewOuter', sizerId: 'dashSizer' },
    { key: 'off', prefix: 'off', targetId: 'officerRender', outerId: 'offPreviewOuter', sizerId: 'offSizer' }
  ].forEach(cfg => {
    const target = $(cfg.targetId);
    const outer = $(cfg.outerId);
    const sizer = $(cfg.sizerId);
    const zoomVal = $(cfg.prefix + 'ZoomVal');
    let zoom = 1;

    function refresh() {
      target.style.transform = `scale(${zoom})`;
      const w = target.offsetWidth || REPORT_WIDTH;
      const h = target.offsetHeight;
      sizer.style.width = Math.ceil(w * zoom) + 'px';
      sizer.style.height = Math.ceil(h * zoom) + 'px';
      if (zoomVal) zoomVal.textContent = Math.round(zoom * 100) + '%';
    }

    function setZoom(z) {
      zoom = Math.max(0.25, Math.min(2.0, z));
      refresh();
    }

    function fit() {
      const avail = outer.clientWidth;
      setZoom(avail > 0 ? (avail - 4) / REPORT_WIDTH : 1);
    }

    $(cfg.prefix + 'ZoomIn').addEventListener('click', () => setZoom(zoom + 0.15));
    $(cfg.prefix + 'ZoomOut').addEventListener('click', () => setZoom(zoom - 0.15));
    $(cfg.prefix + 'ZoomReset').addEventListener('click', fit);

    zoomers[cfg.key] = { refresh, fit };
  });
}

/* Exports must measure/render the report at 100% (transforms distort getBoundingClientRect) */
async function withoutZoom(fn) {
  const targets = [$('dashboardRender'), $('officerRender')];
  const saved = targets.map(t => ({ transform: t.style.transform, transition: t.style.transition }));
  targets.forEach(t => { t.style.transition = 'none'; t.style.transform = 'none'; });
  void document.body.offsetHeight; // force reflow
  try {
    return await fn();
  } finally {
    targets.forEach((t, i) => { t.style.transform = saved[i].transform; t.style.transition = saved[i].transition; });
  }
}

/* ==========================================================================
   PDF / file export
   ========================================================================== */

async function renderNodeToCanvas(node) {
  return await html2canvas(node, {
    scale: 2,
    backgroundColor: '#ffffff',
    useCORS: true,
    scrollX: 0,
    scrollY: 0,
    windowWidth: Math.max(document.documentElement.clientWidth, REPORT_WIDTH + 120)
  });
}

let _downloadsCap;
function getDownloads() {
  if (!_downloadsCap) {
    _downloadsCap = (window.claude && typeof window.claude.use === 'function')
      ? window.claude.use('downloads').catch(() => null)
      : Promise.resolve(null);
  }
  return _downloadsCap;
}

function downloadErrorText(err) {
  const code = err && err.code;
  const map = {
    declined: 'Save was cancelled.',
    too_large: 'This file is too large to save here.',
    rejected_extension: 'This file type is not allowed here.',
    extension_not_enabled: 'File saving for this format is turned off here.',
    rate_limited: 'Please wait a moment and try again.',
    unavailable: 'File saving is not available in this view.',
    not_granted: 'File saving is not available in this view.',
  };
  return (code && map[code]) || (err && err.message) || 'Could not save the file.';
}

async function saveBlob(filename, blob) {
  const downloads = await getDownloads();
  if (downloads) {
    await downloads.save({ filename, data: blob });
    return 'saved';
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  return 'saved';
}

/* Shared button/busy/message wrapper for all four exports */
async function runExport(btnId, busyText, task, okText, errPrefix) {
  if (!computed) { showMsg(genMsg, 'error', 'Generate the report first.'); return; }
  const btn = $(btnId);
  const orig = btn.innerHTML;
  const setLabel = t => { btn.innerHTML = '<span class="spinner"></span>' + t; };
  btn.disabled = true;
  setLabel(busyText);
  try {
    await task(setLabel);
    showMsg(genMsg, 'ok', okText);
  } catch (err) {
    showMsg(genMsg, 'error', errPrefix + ': ' + downloadErrorText(err));
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.innerHTML = orig;
  }
}

function cropCanvas(src, y, h) {
  const c = document.createElement('canvas');
  c.width = src.width;
  c.height = h;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, h);
  ctx.drawImage(src, 0, y, src.width, h, 0, 0, src.width, h);
  return c;
}

/* Places a canvas on the PDF at full usable width. Starts a new page if it will not fit on the
   current one, and slices blocks taller than a page across several pages. */
function placeCanvas(doc, canvas, st) {
  const mmPerPx = st.maxW / canvas.width;
  const hMm = canvas.height * mmPerPx;
  const usable = st.pageH - st.margin * 2;

  if (hMm <= usable) {
    if (st.y > st.margin && st.y + hMm > st.pageH - st.margin) {
      doc.addPage();
      st.y = st.margin;
    }
    doc.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', st.margin, st.y, st.maxW, hMm, undefined, 'FAST');
    st.y += hMm + st.gap;
    return;
  }

  if (st.y > st.margin) {
    doc.addPage();
    st.y = st.margin;
  }
  const slicePx = Math.floor(usable / mmPerPx);
  for (let off = 0; off < canvas.height; off += slicePx) {
    const h = Math.min(slicePx, canvas.height - off);
    if (off > 0) {
      doc.addPage();
      st.y = st.margin;
    }
    const piece = cropCanvas(canvas, off, h);
    const pieceMm = h * mmPerPx;
    doc.addImage(piece.toDataURL('image/jpeg', 0.92), 'JPEG', st.margin, st.y, st.maxW, pieceMm, undefined, 'FAST');
    st.y += pieceMm + st.gap;
  }
}

function newPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 8;
  return { doc, st: { pageW, pageH, margin, maxW: pageW - margin * 2, y: margin, gap: 4 } };
}

$('dashPdfBtn').addEventListener('click', () => runExport('dashPdfBtn', 'Building PDF...', async () => {
  await withoutZoom(async () => {
    const canvas = await renderNodeToCanvas($('dashboardCanvasTarget'));
    const { doc, st } = newPdf();
    const maxH = st.pageH - st.margin * 2;
    const ratio = canvas.height / canvas.width;

    // Fit on one page when that keeps the report reasonably large; otherwise paginate at full width
    let w = st.maxW, h = w * ratio;
    if (h > maxH) { h = maxH; w = h / ratio; }
    if (w >= st.maxW * 0.75) {
      const x = (st.pageW - w) / 2;
      doc.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', x, st.margin, w, h, undefined, 'FAST');
    } else {
      placeCanvas(doc, canvas, st);
    }
    await saveBlob('NDMC_311_Dashboard_' + computed.reportDateStr + '.pdf', doc.output('blob'));
  });
}, 'Dashboard PDF generated and downloaded successfully.', 'PDF'));

$('offPdfBtn').addEventListener('click', () => runExport('offPdfBtn', 'Building PDF...', async (setLabel) => {
  await withoutZoom(async () => {
    const { doc, st } = newPdf();
    // Render each department block on its own: avoids the browser's maximum-canvas-height limit
    const els = [$('officerTitleBlock'), ...document.querySelectorAll('#officerRender [data-dept-block]')];

    for (let i = 0; i < els.length; i++) {
      setLabel(`Building PDF... ${i + 1}/${els.length}`);
      const canvas = await renderNodeToCanvas(els[i]);
      placeCanvas(doc, canvas, st);
      await new Promise(r => setTimeout(r, 0)); // let the UI repaint between blocks
    }
    await saveBlob('NDMC_311_OfficerWise_' + computed.reportDateStr + '.pdf', doc.output('blob'));
  });
}, 'Officer-Wise PDF generated and downloaded successfully.', 'PDF'));

$('excelBtn').addEventListener('click', () => runExport('excelBtn', 'Building Excel...', async () => {
  const blob = await buildStyledExcelBlob();
  await saveBlob('NDMC_311_Report_' + computed.reportDateStr + '.xlsx', blob);
}, 'Excel report ready for download.', 'Excel'));

$('pivotExcelBtn').addEventListener('click', () => runExport('pivotExcelBtn', 'Building Excel...', async () => {
  const blob = await buildOfficerPivotExcelBlob();
  await saveBlob('Officer_Wise_Pending_List_' + computed.reportDateStr + '.xlsx', blob);
}, 'Officer Pending List Excel ready.', 'Excel'));

/* ==========================================================================
   Excel builders
   ========================================================================== */

async function buildStyledExcelBlob() {
  const { deptList, offByDept, grand, reportDateStr } = computed;
  const depts = orderedDepts(deptList); // same order as previews / PDFs
  const wb = new ExcelJS.Workbook();

  const NAVY = 'FF1F4E78', NAVY_DARK = 'FF17365D', LIGHTBLUE = 'FFBDD7EE', GREEN = 'FF92D050', RED = 'FFFF0000';
  const LIGHT_RED = 'FFFEE2E2', DARK_RED = 'FFB91C1C';
  const WHITE_BOLD = { bold: true, color: { argb: 'FFFFFFFF' } };
  const NAVYTXT_BOLD = { bold: true, color: { argb: 'FF12324F' } };
  const GREENTXT_BOLD = { bold: true, color: { argb: 'FF1C3D00' } };
  const REDTXT_BOLD = { bold: true, color: { argb: 'FFFFFFFF' } };
  const THIN = { style: 'thin', color: { argb: 'FF16324D' } };
  const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function fill(cell, argb) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }
  function border(cell) { cell.border = BORDER_ALL; }
  function centre(cell, wrap) { cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: !!wrap }; }

  /* Dashboard sheet */
  const dash = wb.addWorksheet('Dashboard');
  dash.columns = [
    { width: 30 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 12 }
  ];

  dash.mergeCells('A1:I1');
  const dTitle = dash.getCell('A1');
  dTitle.value = '311 APP-PENDING COMPLAINTS – MANAGEMENT DASHBOARD';
  dTitle.font = { ...WHITE_BOLD, size: 15 };
  centre(dTitle); fill(dTitle, NAVY);
  dash.getRow(1).height = 24;

  dash.mergeCells('A2:I2');
  const dAsOf = dash.getCell('A2');
  dAsOf.value = 'As on ' + reportDateStr;
  dAsOf.font = { bold: true };
  dAsOf.alignment = { horizontal: 'right' };

  dash.mergeCells('A4:C4');
  const ovBar = dash.getCell('A4');
  ovBar.value = 'OVERALL'; ovBar.font = WHITE_BOLD; centre(ovBar); fill(ovBar, NAVY_DARK);

  const ovHeadRow = 5, ovValRow = 6;
  const ovCols = [
    { label: 'TOTAL PENDING', bg: LIGHTBLUE, font: NAVYTXT_BOLD, val: grand.total },
    { label: 'Within SLA', bg: GREEN, font: GREENTXT_BOLD, val: `${grand.withinSla} (OUT OF ${grand.total})` },
    { label: 'ESCALATED', bg: RED, font: REDTXT_BOLD, val: `${grand.escalated} (OUT OF ${grand.total})` },
  ];
  ovCols.forEach((col, i) => {
    const hc = dash.getCell(ovHeadRow, i + 1);
    hc.value = col.label; hc.font = col.font; centre(hc); fill(hc, col.bg); border(hc);
    const vc = dash.getCell(ovValRow, i + 1);
    vc.value = col.val; vc.font = { ...col.font, size: 13 }; centre(vc); fill(vc, col.bg); border(vc);
  });

  dash.mergeCells('A8:I8');
  const dwBar = dash.getCell('A8');
  dwBar.value = 'DEPARTMENT-WISE PENDENCY'; dwBar.font = WHITE_BOLD; centre(dwBar); fill(dwBar, NAVY_DARK);

  const headRowNum = 9;
  const headers = ['Department', 'Total Pending', 'Within SLA', 'Escalated', '0–15 Days', '16–30 Days', '31–45 Days', '46–90 Days', 'More than 90 Days'];
  headers.forEach((h, i) => {
    const c = dash.getCell(headRowNum, i + 1);
    c.value = h; c.font = WHITE_BOLD; centre(c, true); fill(c, NAVY_DARK); border(c);
  });

  let r = headRowNum + 1;
  depts.forEach(d => {
    const rowVals = [d.name, d.total, d.withinSla, d.escalated, ...d.buckets];
    rowVals.forEach((v, i) => {
      const c = dash.getCell(r, i + 1);
      c.value = v; border(c);
      if (i === 0) { c.alignment = { horizontal: 'left', vertical: 'middle' }; }
      else centre(c);
      if (i === 1) { fill(c, LIGHTBLUE); c.font = NAVYTXT_BOLD; }
      else if (i === 2) { fill(c, GREEN); c.font = GREENTXT_BOLD; }
      else if (i === 3) { fill(c, RED); c.font = REDTXT_BOLD; }
      else if (i === 8 && d.buckets[4] >= 1) { fill(c, DARK_RED); c.font = REDTXT_BOLD; }
      else if (d.buckets[4] >= 1) { fill(c, LIGHT_RED); c.font = { bold: true, color: { argb: 'FF7F1D1D' } }; }
    });
    r++;
  });
  const totalVals = ['TOTAL', grand.total, grand.withinSla, grand.escalated, ...grand.buckets];
  totalVals.forEach((v, i) => {
    const c = dash.getCell(r, i + 1);
    c.value = v; c.font = WHITE_BOLD; centre(c); fill(c, NAVY); border(c);
  });

  /* Officers Wise sheet */
  const off = wb.addWorksheet('Officers Wise');
  off.columns = [{ width: 7 }, { width: 36 }, { width: 13 }, { width: 11 }, { width: 11 }, { width: 11 }, { width: 11 }, { width: 14 }];

  off.mergeCells('A1:H1');
  const oTitle = off.getCell('A1');
  oTitle.value = 'OFFICER-WISE PENDENCY AS ON ' + reportDateStr.toUpperCase();
  oTitle.font = { ...WHITE_BOLD, size: 15 }; centre(oTitle); fill(oTitle, NAVY);
  off.getRow(1).height = 24;

  let orow = 3;
  const offHeaders = ['S.No.', 'Officer', 'Total Pending', '0–15 Days', '16–30 Days', '31–45 Days', '46–90 Days', 'More than 90 Days'];
  depts.forEach(d => {
    off.mergeCells(`A${orow}:H${orow}`);
    const bar = off.getCell(orow, 1);
    bar.value = d.name; bar.font = WHITE_BOLD; bar.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 }; fill(bar, NAVY); border(bar);
    orow++;

    offHeaders.forEach((h, i) => {
      const c = off.getCell(orow, i + 1);
      c.value = h; c.font = WHITE_BOLD; centre(c, true); fill(c, NAVY_DARK); border(c);
    });
    orow++;

    const officers = offByDept[d.name] || [];
    officers.forEach((o, idx) => {
      const vals = [idx + 1, o.officer, o.total, ...o.buckets];
      vals.forEach((v, i) => {
        const c = off.getCell(orow, i + 1);
        c.value = v; border(c);
        c.alignment = i === 1 ? { horizontal: 'left', vertical: 'middle' } : { horizontal: 'center', vertical: 'middle' };
      });
      orow++;
    });

    const totals = officers.reduce((acc, o) => { acc.total += o.total; for (let i = 0; i < 5; i++) acc.b[i] += o.buckets[i]; return acc; }, { total: 0, b: [0, 0, 0, 0, 0] });
    off.mergeCells(`A${orow}:B${orow}`);
    const totLbl = off.getCell(orow, 1);
    totLbl.value = 'TOTAL'; totLbl.font = WHITE_BOLD; centre(totLbl); fill(totLbl, NAVY); border(totLbl);
    [totals.total, ...totals.b].forEach((v, i) => {
      const c = off.getCell(orow, i + 3);
      c.value = v; c.font = WHITE_BOLD; centre(c); fill(c, NAVY); border(c);
    });
    orow += 2;
  });

  /* Source Data sheet */
  const src = wb.addWorksheet('Source Data');
  if (currentRows.length) {
    const cols = Object.keys(currentRows[0]);
    cols.forEach((c, i) => { src.getColumn(i + 1).width = Math.min(28, Math.max(11, c.length + 2)); });
    const headerRow = src.getRow(1);
    cols.forEach((h, i) => { headerRow.getCell(i + 1).value = h; });
    headerRow.eachCell(c => { c.font = WHITE_BOLD; fill(c, NAVY_DARK); centre(c); });
    currentRows.forEach((row, ri) => {
      const excelRow = src.getRow(ri + 2);
      cols.forEach((h, i) => { excelRow.getCell(i + 1).value = row[h]; });
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function fmtDateTitleUpper(d) {
  const months = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  return d.getDate() + ' ' + months[d.getMonth()] + ', ' + d.getFullYear();
}

/* Officer Pending List (pivot-style, alphabetical by design — mirrors an Excel pivot table) */
async function buildOfficerPivotExcelBlob() {
  const { deptList, offByDept, grand, statusOrder, deptStatusAgg } = computed;
  const wb = new ExcelJS.Workbook();

  const YELLOW = 'FFFFFF00';
  const TITLE_GREEN = 'FF93C47D';
  const HEADER_GREEN = 'FF6AA84F';
  const DEPT_PINK = 'FFEAD1DC';
  const BLACK_BOLD = { bold: true, color: { argb: 'FF000000' } };
  const THIN = { style: 'thin', color: { argb: 'FF808080' } };
  const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  function fill(cell, argb) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }
  function border(cell) { cell.border = BORDER_ALL; }

  const alphaDepts = [...deptList].sort((a, b) => a.name.localeCompare(b.name));

  /* Sheet1: Department x Current Status pivot */
  const s1 = wb.addWorksheet('Sheet1');
  const statusCols = statusOrder && statusOrder.length ? statusOrder : [];
  s1.columns = [{ width: 7 }, { width: 34 }, ...statusCols.map(() => ({ width: 20 })), { width: 14 }];

  const s1LastCol = 2 + statusCols.length + 1;
  s1.mergeCells(1, 1, 1, s1LastCol);
  const s1Title = s1.getCell(1, 1);
  s1Title.value = 'Pending Complaints On NDMC 311 APP As On DATED ' + fmtDateTitleUpper(computed.reportDateObj);
  s1Title.font = { ...BLACK_BOLD, size: 13 };
  fill(s1Title, TITLE_GREEN);
  s1.getRow(1).height = 22;

  const s1HeadRow = 2;
  ['S.NO.', 'Department', ...statusCols, 'Grand Total'].forEach((h, i) => {
    const c = s1.getCell(s1HeadRow, i + 1);
    c.value = h; c.font = BLACK_BOLD; fill(c, YELLOW); border(c);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  let sr = s1HeadRow + 1;
  const statusTotals = statusCols.map(() => 0);
  alphaDepts.forEach((d, idx) => {
    const c0 = s1.getCell(sr, 1); c0.value = idx + 1; c0.font = BLACK_BOLD; c0.alignment = { horizontal: 'center' }; border(c0);
    const c1 = s1.getCell(sr, 2); c1.value = d.name; c1.font = BLACK_BOLD; border(c1);
    const counts = deptStatusAgg[d.name] || {};
    statusCols.forEach((st, i) => {
      const v = counts[st] || 0;
      const c = s1.getCell(sr, 3 + i);
      c.value = v > 0 ? v : ''; c.font = BLACK_BOLD; c.alignment = { horizontal: 'center' }; border(c);
      statusTotals[i] += v;
    });
    const cg = s1.getCell(sr, s1LastCol); cg.value = d.total; cg.font = BLACK_BOLD; cg.alignment = { horizontal: 'center' }; border(cg);
    sr++;
  });
  const g0 = s1.getCell(sr, 1); g0.value = ''; fill(g0, YELLOW); border(g0);
  const g1 = s1.getCell(sr, 2); g1.value = 'Grand Total'; g1.font = BLACK_BOLD; fill(g1, YELLOW); border(g1);
  statusCols.forEach((st, i) => {
    const c = s1.getCell(sr, 3 + i);
    c.value = statusTotals[i] > 0 ? statusTotals[i] : ''; c.font = BLACK_BOLD; c.alignment = { horizontal: 'center' }; fill(c, YELLOW); border(c);
  });
  const gLast = s1.getCell(sr, s1LastCol); gLast.value = grand.total; gLast.font = BLACK_BOLD; gLast.alignment = { horizontal: 'center' }; fill(gLast, YELLOW); border(gLast);

  /* Sheet2: Officer-wise pending list pivot */
  const s2 = wb.addWorksheet('Sheet2');
  s2.columns = [{ width: 34 }, { width: 42 }, { width: 16 }, { width: 22 }];

  s2.mergeCells('A1:D1');
  const s2Title = s2.getCell('A1');
  s2Title.value = 'OFFICER WISE PENDING LIST AS ON DATED ' + fmtDateTitleUpper(computed.reportDateObj);
  s2Title.font = { ...BLACK_BOLD, size: 13 };
  fill(s2Title, YELLOW);
  s2.getRow(1).height = 22;

  const headRow = 2;
  ['Department', 'Currently Assigned To', 'Assigned To Mobile', 'Count of Complaint Number'].forEach((h, i) => {
    const c = s2.getCell(headRow, i + 1);
    c.value = h; c.font = BLACK_BOLD; fill(c, HEADER_GREEN); border(c);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  let r = headRow + 1;
  alphaDepts.forEach(d => {
    const officers = [...(offByDept[d.name] || [])].sort((a, b) => a.officer.localeCompare(b.officer));
    const deptStartRow = r;
    officers.forEach(o => {
      const c0 = s2.getCell(r, 1); c0.value = ''; border(c0);
      const c1 = s2.getCell(r, 2); c1.value = o.officer; c1.font = BLACK_BOLD; border(c1);
      const c2 = s2.getCell(r, 3); c2.value = o.mobile || ''; c2.font = BLACK_BOLD; c2.alignment = { horizontal: 'center' }; border(c2);
      const c3 = s2.getCell(r, 4); c3.value = o.total; c3.font = BLACK_BOLD; c3.alignment = { horizontal: 'center' }; border(c3);
      r++;
    });
    const deptCell = s2.getCell(deptStartRow, 1);
    deptCell.value = d.name;
    deptCell.font = BLACK_BOLD;
    fill(deptCell, DEPT_PINK);
    border(deptCell);

    const dt0 = s2.getCell(r, 1); dt0.value = d.name + ' Total'; dt0.font = BLACK_BOLD; fill(dt0, YELLOW); border(dt0);
    const dt1 = s2.getCell(r, 2); dt1.value = ''; fill(dt1, YELLOW); border(dt1);
    const dt2 = s2.getCell(r, 3); dt2.value = ''; fill(dt2, YELLOW); border(dt2);
    const dt3 = s2.getCell(r, 4); dt3.value = d.total; dt3.font = BLACK_BOLD; dt3.alignment = { horizontal: 'center' }; fill(dt3, YELLOW); border(dt3);
    r++;
  });

  const gt0 = s2.getCell(r, 1); gt0.value = 'Grand Total'; gt0.font = { ...BLACK_BOLD, size: 12 }; fill(gt0, YELLOW); border(gt0);
  const gt1 = s2.getCell(r, 2); gt1.value = ''; fill(gt1, YELLOW); border(gt1);
  const gt2 = s2.getCell(r, 3); gt2.value = ''; fill(gt2, YELLOW); border(gt2);
  const gt3 = s2.getCell(r, 4); gt3.value = grand.total; gt3.font = { ...BLACK_BOLD, size: 12 }; gt3.alignment = { horizontal: 'center' }; fill(gt3, YELLOW); border(gt3);

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ==========================================================================
   Cursor spotlight, card sheen, click sparks
   ========================================================================== */

function initCursorLighting() {
  const coarse = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
  if (REDUCED_MOTION || coarse) return;

  const spotlight = $('cursorSpotlight');
  let tx = window.innerWidth / 2, ty = window.innerHeight / 2;
  let sx = tx, sy = ty, raf = 0;

  function paint() {
    if (spotlight) spotlight.style.transform = `translate(${sx}px, ${sy}px) translate(-50%, -50%)`;
  }
  paint();

  // The loop only runs while the spotlight is still catching up with the cursor
  function loop() {
    sx += (tx - sx) * 0.12;
    sy += (ty - sy) * 0.12;
    paint();
    if (Math.abs(tx - sx) > 0.5 || Math.abs(ty - sy) > 0.5) raf = requestAnimationFrame(loop);
    else raf = 0;
  }

  window.addEventListener('mousemove', e => {
    tx = e.clientX; ty = e.clientY;
    if (!raf) raf = requestAnimationFrame(loop);
  }, { passive: true });

  // Per-card, card-relative sheen coordinates (no page-wide style recalculation)
  const area = document.querySelector('.workflow-container');
  if (area) {
    area.addEventListener('mousemove', e => {
      const card = e.target.closest('.card');
      if (!card) return;
      const rect = card.getBoundingClientRect();
      card.style.setProperty('--mouse-x', (e.clientX - rect.left) + 'px');
      card.style.setProperty('--mouse-y', (e.clientY - rect.top) + 'px');
    }, { passive: true });
  }

  window.addEventListener('click', e => createClickSparks(e.clientX, e.clientY));
}

/* Cards tilt toward the cursor (restored). Skipped on touch screens and reduced-motion. */
function init3DTilt() {
  const coarse = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
  if (REDUCED_MOTION || coarse) return;

  document.querySelectorAll('.tilt-card').forEach(card => {
    let raf = 0, lastEvent = null;

    function apply() {
      raf = 0;
      if (!lastEvent) return;
      const rect = card.getBoundingClientRect();
      const x = lastEvent.clientX - rect.left;
      const y = lastEvent.clientY - rect.top;
      const cx = rect.width / 2, cy = rect.height / 2;
      const rotateX = ((y - cy) / cy) * -3.5;
      const rotateY = ((x - cx) / cx) * 3.5;
      card.style.transform = `perspective(1000px) rotateX(${rotateX}deg) rotateY(${rotateY}deg) translateY(-2px)`;
    }

    card.addEventListener('mousemove', e => {
      lastEvent = e;
      card.classList.add('tilting');
      if (!raf) raf = requestAnimationFrame(apply);
    }, { passive: true });

    card.addEventListener('mouseleave', () => {
      lastEvent = null;
      card.classList.remove('tilting');
      card.style.transform = 'perspective(1000px) rotateX(0deg) rotateY(0deg) translateY(0px)';
    });
  });
}

function createClickSparks(x, y) {
  const sparkColors = ['#3b82f6', '#6366f1', '#06b6d4', '#10b981', '#f43f5e', '#a855f7'];
  for (let i = 0; i < 10; i++) {
    const spark = document.createElement('div');
    spark.className = 'cursor-spark';
    const size = Math.random() * 6 + 3;
    const color = sparkColors[Math.floor(Math.random() * sparkColors.length)];
    spark.style.width = size + 'px';
    spark.style.height = size + 'px';
    spark.style.left = x + 'px';
    spark.style.top = y + 'px';
    spark.style.backgroundColor = color;
    spark.style.boxShadow = `0 0 10px ${color}`;

    const angle = Math.random() * Math.PI * 2;
    const dist = Math.random() * 55 + 25;
    spark.style.setProperty('--dx', Math.cos(angle) * dist + 'px');
    spark.style.setProperty('--dy', Math.sin(angle) * dist + 'px');

    document.body.appendChild(spark);
    setTimeout(() => spark.remove(), 600);
  }
}

/* ==========================================================================
   Theme
   ========================================================================== */

function initThemeToggle() {
  const themeBtn = $('themeToggleBtn');
  if (!themeBtn) return;

  // The inline <head> script already applied the saved theme before first paint
  const current = () => document.documentElement.getAttribute('data-theme') || 'light';
  themeBtn.setAttribute('aria-pressed', current() === 'dark' ? 'true' : 'false');

  themeBtn.addEventListener('click', () => {
    const next = current() === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    themeBtn.setAttribute('aria-pressed', next === 'dark' ? 'true' : 'false');
    safeStorageSet('ndmc_theme', next);
  });
}

/* ==========================================================================
   3D hero (Three.js) — paused off-screen / in background tabs, optional if WebGL fails
   ========================================================================== */

function initHero3D() {
  const container = $('hero3dContainer');
  const canvas = $('hero3dCanvas');
  if (!container || !canvas || typeof THREE === 'undefined') return;
  if (container.clientWidth === 0 || container.clientHeight === 0) return; // hidden on small screens

  let renderer;
  try {
    renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true });
  } catch (e) {
    container.style.display = 'none';
    return;
  }

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 0.1, 1000);
  camera.position.z = 6.4;

  renderer.setSize(container.clientWidth, container.clientHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

  scene.add(new THREE.AmbientLight(0xffffff, 0.9));

  const dirLight1 = new THREE.DirectionalLight(0x6366f1, 1.4);
  dirLight1.position.set(4, 5, 4);
  scene.add(dirLight1);

  const dirLight2 = new THREE.DirectionalLight(0x06b6d4, 1.2);
  dirLight2.position.set(-4, -3, 2);
  scene.add(dirLight2);

  const pointLight = new THREE.PointLight(0xf43f5e, 1.0, 10);
  pointLight.position.set(0, 2, 3);
  scene.add(pointLight);

  const heroGroup = new THREE.Group();
  scene.add(heroGroup);

  const icoMesh = new THREE.Mesh(
    new THREE.IcosahedronGeometry(1.5, 0),
    new THREE.MeshPhysicalMaterial({
      color: 0xffffff, emissive: 0x3b82f6, emissiveIntensity: 0.2,
      roughness: 0.1, metalness: 0.1, transmission: 0.75, ior: 1.55, thickness: 1.6,
      transparent: true, opacity: 0.94
    })
  );
  heroGroup.add(icoMesh);

  const wireMesh = new THREE.Mesh(
    new THREE.IcosahedronGeometry(1.52, 0),
    new THREE.MeshBasicMaterial({ color: 0x6366f1, wireframe: true, transparent: true, opacity: 0.4 })
  );
  heroGroup.add(wireMesh);

  const ring1 = new THREE.Mesh(
    new THREE.TorusGeometry(2.2, 0.035, 16, 100),
    new THREE.MeshStandardMaterial({ color: 0x6366f1, roughness: 0.2, metalness: 0.85, emissive: 0x3b82f6, emissiveIntensity: 0.25 })
  );
  ring1.rotation.x = Math.PI / 3;
  heroGroup.add(ring1);

  const ring2 = new THREE.Mesh(
    new THREE.TorusGeometry(2.5, 0.025, 16, 100),
    new THREE.MeshStandardMaterial({ color: 0x06b6d4, roughness: 0.2, metalness: 0.85, emissive: 0x06b6d4, emissiveIntensity: 0.25 })
  );
  ring2.rotation.y = Math.PI / 4;
  ring2.rotation.x = -Math.PI / 6;
  heroGroup.add(ring2);

  const particleGeo = new THREE.SphereGeometry(0.05, 8, 8);
  const particleMat = new THREE.MeshStandardMaterial({ color: 0x8b5cf6, roughness: 0.3, metalness: 0.7, emissive: 0x6366f1, emissiveIntensity: 0.6 });
  const particles = [];
  for (let i = 0; i < 40; i++) {
    const p = new THREE.Mesh(particleGeo, particleMat);
    const radius = 2.0 + Math.random() * 1.4;
    const theta = Math.random() * Math.PI * 2;
    const phi = Math.acos(Math.random() * 2 - 1);
    p.position.set(radius * Math.sin(phi) * Math.cos(theta), radius * Math.sin(phi) * Math.sin(theta), radius * Math.cos(phi));
    p.userData = { baseY: p.position.y, speed: 0.005 + Math.random() * 0.015, offset: Math.random() * Math.PI * 2 };
    heroGroup.add(p);
    particles.push(p);
  }

  let targetRotX = 0, targetRotY = 0;

  function aim(clientX, clientY) {
    const rect = container.getBoundingClientRect();
    targetRotY = ((clientX - rect.left - rect.width / 2) / (rect.width / 2)) * 0.8;
    targetRotX = ((clientY - rect.top - rect.height / 2) / (rect.height / 2)) * 0.8;
  }

  container.addEventListener('mousemove', e => aim(e.clientX, e.clientY));
  container.addEventListener('mouseleave', () => { targetRotX = 0; targetRotY = 0; });
  container.addEventListener('touchmove', e => {
    if (e.touches.length) aim(e.touches[0].clientX, e.touches[0].clientY);
  }, { passive: true });

  const clock = new THREE.Clock();

  function renderFrame() {
    const t = clock.getElapsedTime();
    heroGroup.position.y = Math.sin(t * 1.2) * 0.12;
    icoMesh.rotation.y += 0.006; icoMesh.rotation.x += 0.004;
    wireMesh.rotation.y += 0.006; wireMesh.rotation.x += 0.004;
    ring1.rotation.z += 0.008;
    ring2.rotation.z -= 0.006;
    particles.forEach(p => {
      p.position.y = p.userData.baseY + Math.sin(t * 2 + p.userData.offset) * 0.15;
      p.rotation.y += p.userData.speed;
    });
    heroGroup.rotation.y += (targetRotY - heroGroup.rotation.y) * 0.05;
    heroGroup.rotation.x += (targetRotX - heroGroup.rotation.x) * 0.05;
    renderer.render(scene, camera);
  }

  // Reduced motion: draw a single static frame
  if (REDUCED_MOTION) { renderFrame(); return; }

  let running = false, rafId = 0, onScreen = true;

  function loop() {
    if (!running) return;
    rafId = requestAnimationFrame(loop);
    renderFrame();
  }
  function start() { if (!running && onScreen && !document.hidden) { running = true; loop(); } }
  function stop() { running = false; cancelAnimationFrame(rafId); }

  if ('IntersectionObserver' in window) {
    new IntersectionObserver(entries => {
      onScreen = entries[0].isIntersecting;
      onScreen ? start() : stop();
    }, { threshold: 0.05 }).observe(container);
  }
  document.addEventListener('visibilitychange', () => { document.hidden ? stop() : start(); });
  start();

  window.addEventListener('resize', () => {
    const w = container.clientWidth, h = container.clientHeight;
    if (!w || !h) return;
    camera.aspect = w / h;
    camera.updateProjectionMatrix();
    renderer.setSize(w, h);
  });
}

/* ==========================================================================
   Init — each step is isolated so one failure cannot break the rest
   ========================================================================== */

function safeInit(fn, label) {
  try { fn(); } catch (err) { console.warn('Init step failed (' + label + '):', err); }
}

function checkLibraries() {
  const missing = [];
  if (typeof XLSX === 'undefined') missing.push('Excel reader (SheetJS)');
  if (typeof ExcelJS === 'undefined') missing.push('Excel writer (ExcelJS)');
  if (!window.jspdf) missing.push('PDF builder (jsPDF)');
  if (typeof html2canvas === 'undefined') missing.push('html2canvas');
  if (missing.length) {
    showMsg(uploadMsg, 'error',
      'Some required libraries did not load (' + missing.join(', ') + '). Check your internet connection or network restrictions, then reload the page.');
  }
}

function setDefaultDate() {
  reportDateInput.value = localDateInputValue(new Date());
}

function applyTouchHint() {
  if (window.matchMedia && window.matchMedia('(pointer: coarse)').matches) {
    const hintLine = dropZone.querySelector('.drop-primary-text');
    if (hintLine) hintLine.textContent = 'Tap here to select your complaint Excel';
  }
}

// Disabled nav links should not navigate
document.addEventListener('click', e => {
  const link = e.target.closest && e.target.closest('.nav-link.is-disabled');
  if (link) e.preventDefault();
});

document.addEventListener('DOMContentLoaded', () => {
  safeInit(setDefaultDate, 'default date');
  safeInit(applyTouchHint, 'touch hint');
  safeInit(checkLibraries, 'library check');
  safeInit(initThemeToggle, 'theme');
  safeInit(initPreviewControls, 'preview controls');
  safeInit(refreshNav, 'nav');
  safeInit(initCursorLighting, 'cursor lighting');
  safeInit(init3DTilt, 'card tilt');
  safeInit(initHero3D, '3D hero');
});
