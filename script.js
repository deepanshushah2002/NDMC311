const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
let workbook = null;
let currentHeaders = [];
let currentRows = [];
let computed = null; // {deptList, offByDept, grand, reportDateStr, reportDateObj, totalRows, unparsedDates, statusOrder, deptStatusAgg}

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
const uploadMsg = document.getElementById('uploadMsg');
const mapCard = document.getElementById('mapCard');
const mapTable = document.getElementById('mapTable');
const sheetSelect = document.getElementById('sheetSelect');
const reportDateInput = document.getElementById('reportDateInput');
const mapMsg = document.getElementById('mapMsg');

/* ---------------- File Drag & Drop ---------------- */

dropZone.addEventListener('click', (e) => {
  if (e.target !== fileInput) {
    e.preventDefault();
    fileInput.click();
  }
});

['dragenter', 'dragover'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.add('has-file');
  });
});

['dragleave', 'dragend'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    if (!fileInput.files.length) dropZone.classList.remove('has-file');
  });
});

dropZone.addEventListener('drop', e => {
  e.preventDefault(); e.stopPropagation();
  const dt = e.dataTransfer;
  if (dt && dt.files && dt.files.length) {
    fileInput.files = dt.files;
    handleFile(dt.files[0]);
  }
});

// Safety net: stop browser from opening dropped files outside drop zone
['dragover', 'drop'].forEach(evt => {
  window.addEventListener(evt, e => { e.preventDefault(); }, false);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length) handleFile(fileInput.files[0]);
});

if (window.matchMedia && window.matchMedia('(pointer: coarse)').matches) {
  const hintLine = dropZone.querySelector('.drop-primary-text');
  if (hintLine) hintLine.textContent = 'Tap here to select your complaint Excel';
}

// Default report date = today
(function () {
  const t = new Date();
  reportDateInput.value = t.toISOString().slice(0, 10);
})();

function showMsg(el, type, text) {
  el.className = 'msg ' + type;
  el.textContent = text;
}

function handleFile(file) {
  fileNameEl.textContent = '📁 ' + file.name;
  dropZone.classList.add('has-file');
  uploadMsg.className = 'msg';
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array', cellDates: true });
      sheetSelect.innerHTML = '';
      workbook.SheetNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name; opt.textContent = name;
        sheetSelect.appendChild(opt);
      });
      // Prefer a sheet that looks like complaint data
      let preferred = workbook.SheetNames.find(n => /complaint/i.test(n)) || workbook.SheetNames[0];
      sheetSelect.value = preferred;
      loadSheet(preferred);
      mapCard.classList.remove('hidden');
      showMsg(uploadMsg, 'ok', 'File loaded successfully: ' + file.name);
      mapCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      showMsg(uploadMsg, 'error', 'Could not read this file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener('change', () => loadSheet(sheetSelect.value));

function loadSheet(name) {
  const ws = workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  if (!json.length) { showMsg(mapMsg, 'error', 'Selected sheet has no data rows.'); return; }
  currentHeaders = Object.keys(json[0]);
  currentRows = json;
  buildMapTable();
}

function normKey(s) { return s.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ''); }

function guessColumn(candidates) {
  const norm = currentHeaders.map(normKey);
  for (const c of candidates) {
    const nc = normKey(c);
    let idx = norm.indexOf(nc);
    if (idx >= 0) return currentHeaders[idx];
  }
  for (const c of candidates) {
    const nc = normKey(c);
    let idx = norm.findIndex(h => h.includes(nc) || nc.includes(h));
    if (idx >= 0) return currentHeaders[idx];
  }
  return currentHeaders[0] || '';
}

const FIELD_DEFS = [
  { key: 'date', label: 'Complaint Date column', candidates: ['Date', 'Complaint Date'] },
  { key: 'dept', label: 'Department column', candidates: ['Department'] },
  { key: 'officer', label: 'Currently Assigned To (Officer) column', candidates: ['Currently Assigned To'] },
  { key: 'esc', label: 'Is Escalated column', candidates: ['Is Escalated'] },
  { key: 'mobile', label: 'Assigned To Mobile column (for pivot list export)', candidates: ['Assigned To Mobile', 'Mobile'] },
  { key: 'status', label: 'Current Status column (for status-wise pivot export)', candidates: ['Current Satus', 'Current Status', 'Status'] },
];
let mapSelections = {};

function buildMapTable() {
  mapTable.innerHTML = '';
  FIELD_DEFS.forEach(f => {
    const guess = guessColumn(f.candidates);
    mapSelections[f.key] = guess;
    const tr = document.createElement('tr');
    const td1 = document.createElement('td');
    td1.textContent = f.label;
    const td2 = document.createElement('td');
    const sel = document.createElement('select');
    currentHeaders.forEach(h => {
      const opt = document.createElement('option');
      opt.value = h; opt.textContent = h;
      if (h === guess) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', () => { mapSelections[f.key] = sel.value; });
    td2.appendChild(sel);
    tr.appendChild(td1); tr.appendChild(td2);
    mapTable.appendChild(tr);
  });
  showMsg(mapMsg, 'info', 'Columns auto-detected. Adjust if needed, verify report date, and click Generate.');
}

function parseDateVal(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return new Date(d.y, d.m - 1, d.d);
    } catch (e) { }
  }
  if (typeof v === 'string') {
    const s = v.trim();
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

document.getElementById('generateBtn').addEventListener('click', () => {
  try {
    if (!reportDateInput.value) { showMsg(mapMsg, 'error', 'Please choose a report date.'); return; }
    const [ry, rm, rd] = reportDateInput.value.split('-').map(x => parseInt(x, 10));
    const reportDate = new Date(ry, rm - 1, rd);
    const reportDateStr = fmtDateDDMMMYYYY(reportDate);

    const colDate = mapSelections.date, colDept = mapSelections.dept,
      colOfficer = mapSelections.officer, colEsc = mapSelections.esc,
      colMobile = mapSelections.mobile, colStatus = mapSelections.status;

    const deptAgg = {}; // dept -> {total,escalated,buckets[5]}
    const offAgg = {};  // dept||officer -> {dept,officer,total,buckets[5],mobile}
    const statusOrder = []; // first-seen order of distinct status values
    const deptStatusAgg = {}; // dept -> {status: count}
    let totalRows = 0, unparsedDates = 0;

    currentRows.forEach(r => {
      let dept = (r[colDept] === undefined || r[colDept] === null) ? '' : r[colDept].toString().trim();
      if (dept === '') dept = 'Unspecified Department';
      let officer = (r[colOfficer] === undefined || r[colOfficer] === null) ? '' : r[colOfficer].toString().trim();
      if (officer === '' || officer === '-') officer = 'Unassigned';
      let mobile = (colMobile && r[colMobile] !== undefined && r[colMobile] !== null) ? r[colMobile].toString().trim() : '';
      if (mobile === '-') mobile = '';
      let status = (colStatus && r[colStatus] !== undefined && r[colStatus] !== null) ? r[colStatus].toString().trim() : '';
      if (status === '') status = 'Unspecified';
      const escRaw = (r[colEsc] === undefined || r[colEsc] === null) ? '' : r[colEsc].toString().trim().toLowerCase();
      const isEsc = (escRaw === 'yes' || escRaw === 'y' || escRaw === 'true' || escRaw === '1');
      const dateVal = parseDateVal(r[colDate]);
      let age = 0;
      if (dateVal) {
        age = Math.floor((reportDate - dateVal) / 86400000);
        if (age < 0) age = 0;
      } else { unparsedDates++; }
      const b = bucketOf(age);

      totalRows++;
      if (!deptAgg[dept]) deptAgg[dept] = { total: 0, escalated: 0, buckets: [0, 0, 0, 0, 0] };
      deptAgg[dept].total++;
      if (isEsc) { deptAgg[dept].escalated++; deptAgg[dept].buckets[b]++; }

      const key = dept + '||' + officer;
      if (!offAgg[key]) offAgg[key] = { dept, officer, total: 0, buckets: [0, 0, 0, 0, 0], mobile: '' };
      offAgg[key].total++;
      offAgg[key].buckets[b]++;
      if (!offAgg[key].mobile && mobile) offAgg[key].mobile = mobile;

      if (!statusOrder.includes(status)) statusOrder.push(status);
      if (!deptStatusAgg[dept]) deptStatusAgg[dept] = {};
      deptStatusAgg[dept][status] = (deptStatusAgg[dept][status] || 0) + 1;
    });

    // department list sorted desc by total pending
    let deptList = Object.keys(deptAgg).map(name => {
      const a = deptAgg[name];
      return { name, total: a.total, escalated: a.escalated, withinSla: a.total - a.escalated, buckets: a.buckets };
    });
    deptList.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));

    // officers grouped by department, ordered same as deptList; sorted desc within dept
    const offByDept = {};
    Object.values(offAgg).forEach(o => {
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

    computed = { deptList, offByDept, grand, reportDateStr, reportDateObj: reportDate, totalRows, unparsedDates, statusOrder: statusOrder.sort((a, b) => a.localeCompare(b)), deptStatusAgg };

    renderSummary();
    renderDashboard();
    renderOfficerWise();
    document.getElementById('summaryCard').classList.remove('hidden');
    document.getElementById('dashPreviewCard').classList.remove('hidden');
    document.getElementById('offPreviewCard').classList.remove('hidden');
    document.getElementById('summaryCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    showMsg(mapMsg, 'error', 'Error while generating: ' + err.message);
    console.error(err);
  }
});

function renderSummary() {
  const g = computed.grand;
  document.getElementById('sTotal').textContent = g.total.toLocaleString();
  document.getElementById('sSla').textContent = g.withinSla.toLocaleString();
  document.getElementById('sEsc').textContent = g.escalated.toLocaleString();
  const bucketSum = g.buckets.reduce((a, b) => a + b, 0);
  const offTotalCheck = Object.values(computed.offByDept).flat().every(o => o.buckets.reduce((a, b) => a + b, 0) === o.total);
  const deptSumCheck = computed.deptList.every(d => d.withinSla + d.escalated === d.total);
  
  let html = '';
  html += `Total records ingested: <b>${computed.totalRows}</b>` + (computed.unparsedDates ? ` &nbsp;(<span class="bad">${computed.unparsedDates} rows had an unreadable date &mdash; aged as 0 days</span>)` : '') + '<br>';
  html += `Integrity Check &middot; Total Pending = Within SLA + Escalated: <span class="${deptSumCheck ? 'ok' : 'bad'}">${deptSumCheck ? 'PASS' : 'FAIL'}</span> &nbsp;&middot;&nbsp; `;
  html += `Dashboard aging buckets (${bucketSum}) = Escalated (${g.escalated}): <span class="${bucketSum === g.escalated ? 'ok' : 'bad'}">${bucketSum === g.escalated ? 'PASS' : 'FAIL'}</span> &nbsp;&middot;&nbsp; `;
  html += `Officer aging consistency: <span class="${offTotalCheck ? 'ok' : 'bad'}">${offTotalCheck ? 'PASS' : 'FAIL'}</span>`;
  document.getElementById('checklist').innerHTML = html;
}

function renderDashboard() {
  const { deptList, grand, reportDateStr } = computed;
  let rows = deptList.map(d => `
    <tr>
      <td class="dept-name-cell">${escapeHtml(d.name)}</td>
      <td class="lightblue">${d.total}</td>
      <td class="green">${d.withinSla}</td>
      <td class="red">${d.escalated}</td>
      <td>${d.buckets[0]}</td><td>${d.buckets[1]}</td><td>${d.buckets[2]}</td><td>${d.buckets[3]}</td><td>${d.buckets[4]}</td>
    </tr>`).join('');

  const html = `
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
  document.getElementById('dashboardRender').innerHTML = html;
}

function renderOfficerWise() {
  const { deptList, offByDept, reportDateStr } = computed;
  let blocks = '';
  deptList.forEach(d => {
    const officers = offByDept[d.name] || [];
    let rows = officers.map((o, i) => `
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

  const html = `
  <div class="rpt-wrap" id="officerCanvasTarget">
    <div class="off-title" id="officerTitleBlock">OFFICER-WISE PENDENCY AS ON ${reportDateStr.toUpperCase()}</div>
    ${blocks}
  </div>`;
  document.getElementById('officerRender').innerHTML = html;
}

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ---------------- PDF / FILE EXPORT ---------------- */

async function renderNodeToCanvas(node) {
  return await html2canvas(node, { scale: 2, backgroundColor: '#ffffff', useCORS: true });
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

document.getElementById('dashPdfBtn').addEventListener('click', async () => {
  const btn = document.getElementById('dashPdfBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building PDF...';
  try {
    const node = document.getElementById('dashboardCanvasTarget');
    const canvas = await renderNodeToCanvas(node);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 8;
    const maxW = pageW - margin * 2;
    const maxH = pageH - margin * 2;
    const ratio = canvas.height / canvas.width;
    let w = maxW, h = w * ratio;
    if (h > maxH) { h = maxH; w = h / ratio; }
    const x = (pageW - w) / 2;
    const y = margin;
    doc.addImage(canvas.toDataURL('image/jpeg', 0.9), 'JPEG', x, y, w, h, undefined, 'MEDIUM');
    const blob = doc.output('blob');
    await saveBlob('NDMC_311_Dashboard_' + computed.reportDateStr + '.pdf', blob);
    showMsg(genMsg, 'ok', 'Dashboard PDF generated and downloaded successfully.');
  } catch (err) {
    showMsg(genMsg, 'error', 'PDF: ' + downloadErrorText(err));
    console.error(err);
  } finally {
    btn.disabled = false; btn.innerHTML = orig;
  }
});

document.getElementById('offPdfBtn').addEventListener('click', async () => {
  const btn = document.getElementById('offPdfBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building PDF...';
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 8;
    const maxW = pageW - margin * 2;

    const container = document.getElementById('officerCanvasTarget');
    const containerRect = container.getBoundingClientRect();
    const segEls = [document.getElementById('officerTitleBlock'), ...document.querySelectorAll('[data-dept-block]')];
    const segments = segEls.map(el => {
      const r = el.getBoundingClientRect();
      return { top: r.top - containerRect.top, height: r.height };
    });

    const fullCanvas = await renderNodeToCanvas(container);
    const scaleFactor = fullCanvas.width / containerRect.width;
    const mmPerPx = maxW / fullCanvas.width;
    const gapMm = 4;

    let curY = margin;
    segments.forEach((seg, idx) => {
      const segTopPx = Math.round(seg.top * scaleFactor);
      const segHeightPx = Math.max(1, Math.round(seg.height * scaleFactor));
      const segHmm = segHeightPx * mmPerPx;

      if (idx > 0 && curY + segHmm > pageH - margin) {
        doc.addPage();
        curY = margin;
      }

      const cropCanvas = document.createElement('canvas');
      cropCanvas.width = fullCanvas.width;
      cropCanvas.height = segHeightPx;
      cropCanvas.getContext('2d').drawImage(
        fullCanvas, 0, segTopPx, fullCanvas.width, segHeightPx, 0, 0, fullCanvas.width, segHeightPx
      );
      doc.addImage(cropCanvas.toDataURL('image/jpeg', 0.9), 'JPEG', margin, curY, maxW, segHmm, undefined, 'MEDIUM');
      curY += segHmm + gapMm;
    });
    const blob = doc.output('blob');
    await saveBlob('NDMC_311_OfficerWise_' + computed.reportDateStr + '.pdf', blob);
    showMsg(genMsg, 'ok', 'Officer-Wise PDF generated and downloaded successfully.');
  } catch (err) {
    showMsg(genMsg, 'error', 'PDF: ' + downloadErrorText(err));
    console.error(err);
  } finally {
    btn.disabled = false; btn.innerHTML = orig;
  }
});

document.getElementById('excelBtn').addEventListener('click', async () => {
  const btn = document.getElementById('excelBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building Excel...';
  try {
    const blob = await buildStyledExcelBlob();
    await saveBlob('NDMC_311_Report_' + computed.reportDateStr + '.xlsx', blob);
    showMsg(genMsg, 'ok', 'Excel report ready for download.');
  } catch (err) {
    showMsg(genMsg, 'error', 'Excel: ' + downloadErrorText(err));
    console.error(err);
  } finally {
    btn.disabled = false; btn.innerHTML = orig;
  }
});

async function buildStyledExcelBlob() {
  const { deptList, offByDept, grand, reportDateStr } = computed;
  const wb = new ExcelJS.Workbook();

  const NAVY = 'FF1F4E78', NAVY_DARK = 'FF17365D', LIGHTBLUE = 'FFBDD7EE', GREEN = 'FF92D050', RED = 'FFFF0000';
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
  dTitle.value = '311 APP-PENDING COMPLAINTS \u2013 MANAGEMENT DASHBOARD';
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
  const headers = ['Department', 'Total Pending', 'Within SLA', 'Escalated', '0\u201315 Days', '16\u201330 Days', '31\u201345 Days', '46\u201390 Days', 'More than 90 Days'];
  headers.forEach((h, i) => {
    const c = dash.getCell(headRowNum, i + 1);
    c.value = h; c.font = WHITE_BOLD; centre(c, true); fill(c, NAVY_DARK); border(c);
  });

  let r = headRowNum + 1;
  deptList.forEach(d => {
    const rowVals = [d.name, d.total, d.withinSla, d.escalated, ...d.buckets];
    rowVals.forEach((v, i) => {
      const c = dash.getCell(r, i + 1);
      c.value = v; border(c);
      if (i === 0) { c.alignment = { horizontal: 'left', vertical: 'middle' }; }
      else centre(c);
      if (i === 1) { fill(c, LIGHTBLUE); c.font = NAVYTXT_BOLD; }
      else if (i === 2) { fill(c, GREEN); c.font = GREENTXT_BOLD; }
      else if (i === 3) { fill(c, RED); c.font = REDTXT_BOLD; }
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
  const offHeaders = ['S.No.', 'Officer', 'Total Pending', '0\u201315 Days', '16\u201330 Days', '31\u201345 Days', '46\u201390 Days', 'More than 90 Days'];
  deptList.forEach(d => {
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

/* Officer Pending List (pivot-style) Excel */
function fmtDateTitleUpper(d) {
  const months = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  return d.getDate() + ' ' + months[d.getMonth()] + ', ' + d.getFullYear();
}

document.getElementById('pivotExcelBtn').addEventListener('click', async () => {
  const btn = document.getElementById('pivotExcelBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building Excel...';
  try {
    const blob = await buildOfficerPivotExcelBlob();
    await saveBlob('Officer_Wise_Pending_List_' + computed.reportDateStr + '.xlsx', blob);
    showMsg(genMsg, 'ok', 'Officer Pending List Excel ready.');
  } catch (err) {
    showMsg(genMsg, 'error', 'Excel: ' + downloadErrorText(err));
    console.error(err);
  } finally {
    btn.disabled = false; btn.innerHTML = orig;
  }
});

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
  // Grand Total row
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
   Interactive 3D Hero Scene (Three.js)
   ========================================================================== */

function initHero3D() {
  const container = document.getElementById('hero3dContainer');
  const canvas = document.getElementById('hero3dCanvas');
  if (!container || !canvas || typeof THREE === 'undefined') return;

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 0.1, 1000);
  camera.position.z = 6;

  const renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true });
  renderer.setSize(container.clientWidth, container.clientHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

  // Lighting
  const ambientLight = new THREE.AmbientLight(0xffffff, 0.85);
  scene.add(ambientLight);

  const dirLight1 = new THREE.DirectionalLight(0x6366f1, 1.2);
  dirLight1.position.set(4, 5, 4);
  scene.add(dirLight1);

  const dirLight2 = new THREE.DirectionalLight(0x06b6d4, 1.0);
  dirLight2.position.set(-4, -3, 2);
  scene.add(dirLight2);

  const pointLight = new THREE.PointLight(0xf43f5e, 0.8, 10);
  pointLight.position.set(0, 2, 3);
  scene.add(pointLight);

  // Group for full 3D object
  const heroGroup = new THREE.Group();
  scene.add(heroGroup);

  // Central Crystal Icosahedron
  const icoGeo = new THREE.IcosahedronGeometry(1.6, 0);
  const icoMat = new THREE.MeshPhysicalMaterial({
    color: 0xffffff,
    emissive: 0x3b82f6,
    emissiveIntensity: 0.15,
    roughness: 0.1,
    metalness: 0.1,
    transmission: 0.7,
    ior: 1.5,
    thickness: 1.5,
    transparent: true,
    opacity: 0.92,
    wireframe: false
  });
  const icoMesh = new THREE.Mesh(icoGeo, icoMat);
  heroGroup.add(icoMesh);

  // Wireframe Cage overlay
  const wireMat = new THREE.MeshBasicMaterial({
    color: 0x6366f1,
    wireframe: true,
    transparent: true,
    opacity: 0.35
  });
  const wireMesh = new THREE.Mesh(new THREE.IcosahedronGeometry(1.62, 0), wireMat);
  heroGroup.add(wireMesh);

  // Orbital Rings
  const ringGeo1 = new THREE.TorusGeometry(2.3, 0.035, 16, 100);
  const ringMat1 = new THREE.MeshStandardMaterial({
    color: 0x6366f1,
    roughness: 0.2,
    metalness: 0.85,
    emissive: 0x3b82f6,
    emissiveIntensity: 0.2
  });
  const ring1 = new THREE.Mesh(ringGeo1, ringMat1);
  ring1.rotation.x = Math.PI / 3;
  heroGroup.add(ring1);

  const ringGeo2 = new THREE.TorusGeometry(2.6, 0.025, 16, 100);
  const ringMat2 = new THREE.MeshStandardMaterial({
    color: 0x06b6d4,
    roughness: 0.2,
    metalness: 0.85,
    emissive: 0x06b6d4,
    emissiveIntensity: 0.2
  });
  const ring2 = new THREE.Mesh(ringGeo2, ringMat2);
  ring2.rotation.y = Math.PI / 4;
  ring2.rotation.x = -Math.PI / 6;
  heroGroup.add(ring2);

  // Floating particles / data nodes
  const particleCount = 45;
  const particleGeo = new THREE.SphereGeometry(0.05, 8, 8);
  const particleMat = new THREE.MeshStandardMaterial({
    color: 0x8b5cf6,
    roughness: 0.3,
    metalness: 0.7,
    emissive: 0x6366f1,
    emissiveIntensity: 0.5
  });
  
  const particles = [];
  for (let i = 0; i < particleCount; i++) {
    const pMesh = new THREE.Mesh(particleGeo, particleMat);
    const radius = 2.0 + Math.random() * 1.5;
    const theta = Math.random() * Math.PI * 2;
    const phi = Math.acos(Math.random() * 2 - 1);
    
    pMesh.position.x = radius * Math.sin(phi) * Math.cos(theta);
    pMesh.position.y = radius * Math.sin(phi) * Math.sin(theta);
    pMesh.position.z = radius * Math.cos(phi);
    
    pMesh.userData = {
      basePos: pMesh.position.clone(),
      speed: 0.005 + Math.random() * 0.015,
      offset: Math.random() * Math.PI * 2
    };
    heroGroup.add(pMesh);
    particles.push(pMesh);
  }

  // Interactive Mouse Parallax
  let targetRotX = 0;
  let targetRotY = 0;
  let mouseX = 0;
  let mouseY = 0;

  function onMouseMove(e) {
    const rect = container.getBoundingClientRect();
    const x = e.clientX - rect.left - rect.width / 2;
    const y = e.clientY - rect.top - rect.height / 2;
    mouseX = (x / (rect.width / 2));
    mouseY = (y / (rect.height / 2));
    targetRotY = mouseX * 0.8;
    targetRotX = mouseY * 0.8;
  }

  container.addEventListener('mousemove', onMouseMove);
  container.addEventListener('mouseleave', () => {
    targetRotX = 0;
    targetRotY = 0;
  });

  // Touch Support
  container.addEventListener('touchmove', (e) => {
    if (e.touches.length > 0) {
      const rect = container.getBoundingClientRect();
      const x = e.touches[0].clientX - rect.left - rect.width / 2;
      const y = e.touches[0].clientY - rect.top - rect.height / 2;
      targetRotY = (x / (rect.width / 2)) * 0.8;
      targetRotX = (y / (rect.height / 2)) * 0.8;
    }
  }, { passive: true });

  // Render loop
  let clock = new THREE.Clock();
  let animId;

  function animate() {
    animId = requestAnimationFrame(animate);
    const elapsedTime = clock.getElapsedTime();

    // Constant smooth floating & self-rotation
    heroGroup.position.y = Math.sin(elapsedTime * 1.2) * 0.12;
    icoMesh.rotation.y += 0.006;
    icoMesh.rotation.x += 0.004;
    wireMesh.rotation.y += 0.006;
    wireMesh.rotation.x += 0.004;

    ring1.rotation.z += 0.008;
    ring2.rotation.z -= 0.006;

    // Orbit particles
    particles.forEach(p => {
      p.position.y = p.userData.basePos.y + Math.sin(elapsedTime * 2 + p.userData.offset) * 0.15;
      p.rotation.y += p.userData.speed;
    });

    // Mouse Damping Interpolation
    heroGroup.rotation.y += (targetRotY - heroGroup.rotation.y) * 0.05;
    heroGroup.rotation.x += (targetRotX - heroGroup.rotation.x) * 0.05;

    renderer.render(scene, camera);
  }
  animate();

  // Resize Handler
  function handleResize() {
    if (!container) return;
    const width = container.clientWidth;
    const height = container.clientHeight;
    camera.aspect = width / height;
    camera.updateProjectionMatrix();
    renderer.setSize(width, height);
  }
  window.addEventListener('resize', handleResize);
}

/* ==========================================================================
   3D Card Tilt Micro-Interactions
   ========================================================================== */

function init3DTilt() {
  const cards = document.querySelectorAll('.tilt-card');
  cards.forEach(card => {
    card.addEventListener('mousemove', e => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      const centerX = rect.width / 2;
      const centerY = rect.height / 2;
      const rotateX = ((y - centerY) / centerY) * -3.5;
      const rotateY = ((x - centerX) / centerX) * 3.5;

      card.style.transform = `perspective(1000px) rotateX(${rotateX}deg) rotateY(${rotateY}deg) translateY(-2px)`;
    });

    card.addEventListener('mouseleave', () => {
      card.style.transform = 'perspective(1000px) rotateX(0deg) rotateY(0deg) translateY(0px)';
    });
  });
}

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  initHero3D();
  init3DTilt();
});
