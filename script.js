const MONTHS = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
let workbook = null;
let currentHeaders = [];
let currentRows = [];
let computed = null; // {deptList, offByDept, grand, reportDateStr}

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
const uploadMsg = document.getElementById('uploadMsg');
const mapCard = document.getElementById('mapCard');
const mapTable = document.getElementById('mapTable');
const sheetSelect = document.getElementById('sheetSelect');
const reportDateInput = document.getElementById('reportDateInput');
const mapMsg = document.getElementById('mapMsg');

// Some browsers (notably iOS Safari / many mobile browsers) open the
// picker via the label's native click-forwarding to the nested <input>
// already. Also calling fileInput.click() ourselves on top of that can
// fire the dialog twice and make it look like clicking "does nothing" on
// some devices, so we only add our own click-forward when the click did
// NOT originate on the input itself (covers clicks that land on the
// label's padding/icon/text, without double-triggering the native path).
dropZone.addEventListener('click', (e)=>{
  if(e.target !== fileInput){
    e.preventDefault();
    fileInput.click();
  }
});

['dragenter','dragover'].forEach(evt=>{
  dropZone.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.add('has-file');
  });
});
['dragleave','dragend'].forEach(evt=>{
  dropZone.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation();
    if(!fileInput.files.length) dropZone.classList.remove('has-file');
  });
});
dropZone.addEventListener('drop', e=>{
  e.preventDefault(); e.stopPropagation();
  const dt = e.dataTransfer;
  if(dt && dt.files && dt.files.length){ fileInput.files = dt.files; handleFile(dt.files[0]); }
});
// Safety net: stop the browser from navigating away to the raw file if a
// drop ever lands outside the drop zone itself.
['dragover','drop'].forEach(evt=>{
  window.addEventListener(evt, e=>{ e.preventDefault(); }, false);
});
fileInput.addEventListener('change', ()=>{ if(fileInput.files.length) handleFile(fileInput.files[0]); });

// Drag-and-drop of files between apps generally isn't supported on phone
// browsers (Android/iOS) -- only desktop browsers support dragging a file
// out of a file-manager window. Adjust the hint text on touch devices so
// it doesn't promise something that won't work there.
if(window.matchMedia && window.matchMedia('(pointer: coarse)').matches){
  const hintLine = dropZone.querySelector('div:nth-of-type(2)');
  if(hintLine) hintLine.textContent = 'Tap here to choose a file';
}

// default report date = today
(function(){
  const t = new Date();
  reportDateInput.value = t.toISOString().slice(0,10);
})();

function showMsg(el, type, text){
  el.className = 'msg ' + type;
  el.textContent = text;
}

function handleFile(file){
  fileNameEl.textContent = file.name;
  dropZone.classList.add('has-file');
  uploadMsg.className = 'msg';
  const reader = new FileReader();
  reader.onload = function(e){
    try{
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, {type:'array', cellDates:true});
      sheetSelect.innerHTML = '';
      workbook.SheetNames.forEach(name=>{
        const opt = document.createElement('option');
        opt.value = name; opt.textContent = name;
        sheetSelect.appendChild(opt);
      });
      // prefer a sheet that looks like the complaint data
      let preferred = workbook.SheetNames.find(n=>/complaint/i.test(n)) || workbook.SheetNames[0];
      sheetSelect.value = preferred;
      loadSheet(preferred);
      mapCard.classList.remove('hidden');
      showMsg(uploadMsg,'ok','File loaded: ' + file.name);
    }catch(err){
      showMsg(uploadMsg,'error','Could not read this file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener('change', ()=>loadSheet(sheetSelect.value));

function loadSheet(name){
  const ws = workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, {defval:'', raw:true});
  if(!json.length){ showMsg(mapMsg,'error','Selected sheet has no data rows.'); return; }
  currentHeaders = Object.keys(json[0]);
  currentRows = json;
  buildMapTable();
}

function normKey(s){ return s.toString().trim().toLowerCase().replace(/[^a-z0-9]/g,''); }

function guessColumn(candidates){
  const norm = currentHeaders.map(normKey);
  for(const c of candidates){
    const nc = normKey(c);
    let idx = norm.indexOf(nc);
    if(idx>=0) return currentHeaders[idx];
  }
  for(const c of candidates){
    const nc = normKey(c);
    let idx = norm.findIndex(h=>h.includes(nc) || nc.includes(h));
    if(idx>=0) return currentHeaders[idx];
  }
  return currentHeaders[0] || '';
}

const FIELD_DEFS = [
  {key:'date', label:'Complaint Date column', candidates:['Date','Complaint Date']},
  {key:'dept', label:'Department column', candidates:['Department']},
  {key:'officer', label:'Currently Assigned To (Officer) column', candidates:['Currently Assigned To']},
  {key:'esc', label:'Is Escalated column', candidates:['Is Escalated']},
];
let mapSelections = {};

function buildMapTable(){
  mapTable.innerHTML = '';
  FIELD_DEFS.forEach(f=>{
    const guess = guessColumn(f.candidates);
    mapSelections[f.key] = guess;
    const tr = document.createElement('tr');
    const td1 = document.createElement('td');
    td1.style.fontWeight = '600'; td1.style.width='260px'; td1.style.color='#3a4a5c'; td1.style.fontSize='13px';
    td1.textContent = f.label;
    const td2 = document.createElement('td');
    const sel = document.createElement('select');
    sel.style.cssText = 'padding:7px 8px;border:1px solid #c6d0da;border-radius:6px;font-size:13px;min-width:260px;';
    currentHeaders.forEach(h=>{
      const opt = document.createElement('option');
      opt.value = h; opt.textContent = h;
      if(h===guess) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', ()=>{ mapSelections[f.key] = sel.value; });
    td2.appendChild(sel);
    tr.appendChild(td1); tr.appendChild(td2);
    mapTable.appendChild(tr);
  });
  showMsg(mapMsg,'info','Columns auto-detected. Adjust if wrong, then set report date and click Generate.');
}

function parseDateVal(v){
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if(typeof v === 'number'){
    try{
      const d = XLSX.SSF.parse_date_code(v);
      if(d) return new Date(d.y, d.m-1, d.d);
    }catch(e){}
  }
  if(typeof v === 'string'){
    const s = v.trim();
    let m = s.match(/^(\d{1,2})[-\/]([A-Za-z]{3,})[-\/](\d{4})/);
    if(m){
      const day = parseInt(m[1],10);
      const mon = MONTHS[m[2].toLowerCase().slice(0,3)];
      const year = parseInt(m[3],10);
      if(mon!==undefined) return new Date(year, mon, day);
    }
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if(m){
      return new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
    }
    const d2 = new Date(s);
    if(!isNaN(d2)) return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  }
  return null;
}

function bucketOf(days){
  if(days<=15) return 0;
  if(days<=30) return 1;
  if(days<=45) return 2;
  if(days<=90) return 3;
  return 4;
}
const BUCKET_LABELS = ['0\u201315 Days','16\u201330 Days','31\u201345 Days','46\u201390 Days','More than 90 Days'];

function fmtDateDDMMMYYYY(d){
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return String(d.getDate()).padStart(2,'0')+'-'+months[d.getMonth()]+'-'+d.getFullYear();
}

document.getElementById('generateBtn').addEventListener('click', ()=>{
  try{
    if(!reportDateInput.value){ showMsg(mapMsg,'error','Please choose a report date.'); return; }
    const [ry,rm,rd] = reportDateInput.value.split('-').map(x=>parseInt(x,10));
    const reportDate = new Date(ry, rm-1, rd);
    const reportDateStr = fmtDateDDMMMYYYY(reportDate);

    const colDate = mapSelections.date, colDept = mapSelections.dept,
          colOfficer = mapSelections.officer, colEsc = mapSelections.esc;

    const deptAgg = {}; // dept -> {total,escalated,buckets[5]}
    const offAgg = {};  // dept||officer -> {dept,officer,total,buckets[5]}
    let totalRows = 0, unparsedDates = 0;

    currentRows.forEach(r=>{
      let dept = (r[colDept]===undefined||r[colDept]===null) ? '' : r[colDept].toString().trim();
      if(dept==='') dept = 'Unspecified Department';
      let officer = (r[colOfficer]===undefined||r[colOfficer]===null) ? '' : r[colOfficer].toString().trim();
      if(officer===''||officer==='-') officer = 'Unassigned';
      const escRaw = (r[colEsc]===undefined||r[colEsc]===null) ? '' : r[colEsc].toString().trim().toLowerCase();
      const isEsc = (escRaw==='yes' || escRaw==='y' || escRaw==='true' || escRaw==='1');
      const dateVal = parseDateVal(r[colDate]);
      let age = 0;
      if(dateVal){
        age = Math.floor((reportDate - dateVal)/86400000);
        if(age<0) age=0;
      } else { unparsedDates++; }
      const b = bucketOf(age);

      totalRows++;
      if(!deptAgg[dept]) deptAgg[dept] = {total:0, escalated:0, buckets:[0,0,0,0,0]};
      deptAgg[dept].total++;
      if(isEsc){ deptAgg[dept].escalated++; deptAgg[dept].buckets[b]++; }

      const key = dept+'||'+officer;
      if(!offAgg[key]) offAgg[key] = {dept, officer, total:0, buckets:[0,0,0,0,0]};
      offAgg[key].total++;
      offAgg[key].buckets[b]++;
    });

    // department list sorted desc by total pending
    let deptList = Object.keys(deptAgg).map(name=>{
      const a = deptAgg[name];
      return {name, total:a.total, escalated:a.escalated, withinSla:a.total-a.escalated, buckets:a.buckets};
    });
    deptList.sort((a,b)=> b.total - a.total || a.name.localeCompare(b.name));

    // officers grouped by department, ordered same as deptList; sorted desc within dept
    const offByDept = {};
    Object.values(offAgg).forEach(o=>{
      if(!offByDept[o.dept]) offByDept[o.dept]=[];
      offByDept[o.dept].push(o);
    });
    Object.keys(offByDept).forEach(d=>{
      offByDept[d].sort((a,b)=> b.total-a.total || a.officer.localeCompare(b.officer));
    });

    const grand = deptList.reduce((acc,d)=>{
      acc.total+=d.total; acc.escalated+=d.escalated; acc.withinSla+=d.withinSla;
      for(let i=0;i<5;i++) acc.buckets[i]+=d.buckets[i];
      return acc;
    }, {total:0,escalated:0,withinSla:0,buckets:[0,0,0,0,0]});

    computed = {deptList, offByDept, grand, reportDateStr, totalRows, unparsedDates};

    renderSummary();
    renderDashboard();
    renderOfficerWise();
    document.getElementById('summaryCard').classList.remove('hidden');
    document.getElementById('dashPreviewCard').classList.remove('hidden');
    document.getElementById('offPreviewCard').classList.remove('hidden');
    document.getElementById('summaryCard').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    showMsg(mapMsg,'error','Error while generating: '+err.message);
    console.error(err);
  }
});

function renderSummary(){
  const g = computed.grand;
  document.getElementById('sTotal').textContent = g.total;
  document.getElementById('sSla').textContent = g.withinSla;
  document.getElementById('sEsc').textContent = g.escalated;
  const bucketSum = g.buckets.reduce((a,b)=>a+b,0);
  const offTotalCheck = Object.values(computed.offByDept).flat().every(o=> o.buckets.reduce((a,b)=>a+b,0)===o.total);
  const deptSumCheck = computed.deptList.every(d=> d.withinSla+d.escalated===d.total);
  let html = '';
  html += `Total rows read: <b>${computed.totalRows}</b>` + (computed.unparsedDates? ` &nbsp;(<span class="bad">${computed.unparsedDates} rows had an unreadable date &mdash; aged as 0 days</span>)`:'') + '<br>';
  html += `Check &middot; Total Pending = Within SLA + Escalated: <span class="${deptSumCheck?'ok':'bad'}">${deptSumCheck?'PASS':'FAIL'}</span> &nbsp;&middot;&nbsp; `;
  html += `Dashboard aging buckets total (${bucketSum}) = Overall Escalated (${g.escalated}): <span class="${bucketSum===g.escalated?'ok':'bad'}">${bucketSum===g.escalated?'PASS':'FAIL'}</span> &nbsp;&middot;&nbsp; `;
  html += `Officer aging = Officer total (every officer): <span class="${offTotalCheck?'ok':'bad'}">${offTotalCheck?'PASS':'FAIL'}</span>`;
  document.getElementById('checklist').innerHTML = html;
}

function renderDashboard(){
  const {deptList, grand, reportDateStr} = computed;
  let rows = deptList.map(d=>`
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

function renderOfficerWise(){
  const {deptList, offByDept, reportDateStr} = computed;
  let blocks = '';
  deptList.forEach(d=>{
    const officers = offByDept[d.name] || [];
    let rows = officers.map((o,i)=>`
      <tr>
        <td>${i+1}</td>
        <td class="off-name-cell">${escapeHtml(o.officer)}</td>
        <td>${o.total}</td>
        <td>${o.buckets[0]}</td><td>${o.buckets[1]}</td><td>${o.buckets[2]}</td><td>${o.buckets[3]}</td><td>${o.buckets[4]}</td>
      </tr>`).join('');
    const totals = officers.reduce((acc,o)=>{acc.total+=o.total; for(let i=0;i<5;i++) acc.b[i]+=o.buckets[i]; return acc;}, {total:0,b:[0,0,0,0,0]});
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

function escapeHtml(s){
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/* ---------------- PDF / FILE EXPORT ---------------- */

async function renderNodeToCanvas(node){
  return await html2canvas(node, {scale:2, backgroundColor:'#ffffff', useCORS:true});
}

let _downloadsCap; // cached promise
function getDownloads(){
  if(!_downloadsCap){
    _downloadsCap = (window.claude && typeof window.claude.use === 'function')
      ? window.claude.use('downloads').catch(()=>null)
      : Promise.resolve(null);
  }
  return _downloadsCap;
}

function downloadErrorText(err){
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

// Saves a Blob under `filename`, using the artifact downloads capability
// when this page is running as a published artifact, falling back to a
// plain browser download link otherwise (e.g. inside the chat preview).
async function saveBlob(filename, blob){
  const downloads = await getDownloads();
  if(downloads){
    await downloads.save({filename, data: blob});
    return 'saved';
  }
  // fallback: classic anchor download (works in the inline chat preview)
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
  return 'saved';
}

document.getElementById('dashPdfBtn').addEventListener('click', async ()=>{
  const btn = document.getElementById('dashPdfBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building PDF...';
  try{
    const node = document.getElementById('dashboardCanvasTarget');
    const canvas = await renderNodeToCanvas(node);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({orientation:'landscape', unit:'mm', format:'a4', compress:true});
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 8;
    const maxW = pageW - margin*2;
    const maxH = pageH - margin*2;
    const ratio = canvas.height / canvas.width;
    let w = maxW, h = w*ratio;
    if(h > maxH){ h = maxH; w = h/ratio; }
    const x = (pageW - w)/2;
    const y = margin;
    doc.addImage(canvas.toDataURL('image/jpeg', 0.9), 'JPEG', x, y, w, h, undefined, 'MEDIUM');
    const blob = doc.output('blob');
    await saveBlob('NDMC_311_Dashboard_' + computed.reportDateStr + '.pdf', blob);
    showMsg(genMsg,'ok','Dashboard PDF ready.');
  }catch(err){
    showMsg(genMsg,'error','PDF: '+downloadErrorText(err));
    console.error(err);
  }finally{
    btn.disabled = false; btn.innerHTML = orig;
  }
});

document.getElementById('offPdfBtn').addEventListener('click', async ()=>{
  const btn = document.getElementById('offPdfBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building PDF...';
  try{
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({orientation:'landscape', unit:'mm', format:'a4', compress:true});
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 8;
    const maxW = pageW - margin*2;

    // Measure each segment's exact box (title + each department block) in
    // CSS px BEFORE capturing, so slices line up perfectly with no
    // accumulated rounding gaps between them.
    const container = document.getElementById('officerCanvasTarget');
    const containerRect = container.getBoundingClientRect();
    const segEls = [document.getElementById('officerTitleBlock'), ...document.querySelectorAll('[data-dept-block]')];
    const segments = segEls.map(el=>{
      const r = el.getBoundingClientRect();
      return { top: r.top - containerRect.top, height: r.height };
    });

    // One single capture of the whole content, then slice it per segment.
    const fullCanvas = await renderNodeToCanvas(container);
    const scaleFactor = fullCanvas.width / containerRect.width;
    const mmPerPx = maxW / fullCanvas.width;
    const gapMm = 4;

    let curY = margin;
    segments.forEach((seg, idx)=>{
      const segTopPx = Math.round(seg.top * scaleFactor);
      const segHeightPx = Math.max(1, Math.round(seg.height * scaleFactor));
      const segHmm = segHeightPx * mmPerPx;

      if(idx > 0 && curY + segHmm > pageH - margin){
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
    showMsg(genMsg,'ok','Officer-Wise PDF ready.');
  }catch(err){
    showMsg(genMsg,'error','PDF: '+downloadErrorText(err));
    console.error(err);
  }finally{
    btn.disabled = false; btn.innerHTML = orig;
  }
});

document.getElementById('excelBtn').addEventListener('click', async ()=>{
  const btn = document.getElementById('excelBtn');
  const genMsg = document.getElementById('genMsg');
  btn.disabled = true; const orig = btn.innerHTML; btn.innerHTML = '<span class="spinner"></span>Building Excel...';
  try{
    const blob = await buildStyledExcelBlob();
    await saveBlob('NDMC_311_Report_' + computed.reportDateStr + '.xlsx', blob);
    showMsg(genMsg,'ok','Excel file ready.');
  }catch(err){
    showMsg(genMsg,'error','Excel: '+downloadErrorText(err));
    console.error(err);
  }finally{
    btn.disabled = false; btn.innerHTML = orig;
  }
});

// Builds the 3-sheet workbook with the same navy/light-blue/green/red
// colour scheme as the Dashboard & Officer-Wise report, using ExcelJS
// (the plain XLSX writer used elsewhere in this file has no styling API).
async function buildStyledExcelBlob(){
  const {deptList, offByDept, grand, reportDateStr} = computed;
  const wb = new ExcelJS.Workbook();

  const NAVY = 'FF1F4E78', NAVY_DARK = 'FF17365D', LIGHTBLUE = 'FFBDD7EE', GREEN = 'FF92D050', RED = 'FFFF0000';
  const WHITE_BOLD = {bold:true, color:{argb:'FFFFFFFF'}};
  const NAVYTXT_BOLD = {bold:true, color:{argb:'FF12324F'}};
  const GREENTXT_BOLD = {bold:true, color:{argb:'FF1C3D00'}};
  const REDTXT_BOLD = {bold:true, color:{argb:'FFFFFFFF'}};
  const THIN = {style:'thin', color:{argb:'FF16324D'}};
  const BORDER_ALL = {top:THIN,left:THIN,bottom:THIN,right:THIN};

  function fill(cell, argb){ cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb}}; }
  function border(cell){ cell.border = BORDER_ALL; }
  function centre(cell, wrap){ cell.alignment = {horizontal:'center', vertical:'middle', wrapText: !!wrap}; }

  /* ---------------- Dashboard sheet ---------------- */
  const dash = wb.addWorksheet('Dashboard');
  dash.columns = [
    {width:30},{width:12},{width:12},{width:12},{width:10},{width:10},{width:10},{width:10},{width:12}
  ];

  dash.mergeCells('A1:I1');
  const dTitle = dash.getCell('A1');
  dTitle.value = '311 APP-PENDING COMPLAINTS \u2013 MANAGEMENT DASHBOARD';
  dTitle.font = {...WHITE_BOLD, size:15};
  centre(dTitle); fill(dTitle, NAVY);
  dash.getRow(1).height = 24;

  dash.mergeCells('A2:I2');
  const dAsOf = dash.getCell('A2');
  dAsOf.value = 'As on ' + reportDateStr;
  dAsOf.font = {bold:true};
  dAsOf.alignment = {horizontal:'right'};

  dash.mergeCells('A4:C4');
  const ovBar = dash.getCell('A4');
  ovBar.value = 'OVERALL'; ovBar.font = WHITE_BOLD; centre(ovBar); fill(ovBar, NAVY_DARK);

  const ovHeadRow = 5, ovValRow = 6;
  const ovCols = [
    {label:'TOTAL PENDING', bg:LIGHTBLUE, font:NAVYTXT_BOLD, val: grand.total},
    {label:'Within SLA', bg:GREEN, font:GREENTXT_BOLD, val: `${grand.withinSla} (OUT OF ${grand.total})`},
    {label:'ESCALATED', bg:RED, font:REDTXT_BOLD, val: `${grand.escalated} (OUT OF ${grand.total})`},
  ];
  ovCols.forEach((col,i)=>{
    const hc = dash.getCell(ovHeadRow, i+1);
    hc.value = col.label; hc.font = col.font; centre(hc); fill(hc, col.bg); border(hc);
    const vc = dash.getCell(ovValRow, i+1);
    vc.value = col.val; vc.font = {...col.font, size:13}; centre(vc); fill(vc, col.bg); border(vc);
  });

  dash.mergeCells('A8:I8');
  const dwBar = dash.getCell('A8');
  dwBar.value = 'DEPARTMENT-WISE PENDENCY'; dwBar.font = WHITE_BOLD; centre(dwBar); fill(dwBar, NAVY_DARK);

  const headRowNum = 9;
  const headers = ['Department','Total Pending','Within SLA','Escalated','0\u201315 Days','16\u201330 Days','31\u201345 Days','46\u201390 Days','More than 90 Days'];
  headers.forEach((h,i)=>{
    const c = dash.getCell(headRowNum, i+1);
    c.value = h; c.font = WHITE_BOLD; centre(c, true); fill(c, NAVY_DARK); border(c);
  });

  let r = headRowNum + 1;
  deptList.forEach(d=>{
    const rowVals = [d.name, d.total, d.withinSla, d.escalated, ...d.buckets];
    rowVals.forEach((v,i)=>{
      const c = dash.getCell(r, i+1);
      c.value = v; border(c);
      if(i===0){ c.alignment = {horizontal:'left', vertical:'middle'}; }
      else centre(c);
      if(i===1){ fill(c, LIGHTBLUE); c.font = NAVYTXT_BOLD; }
      else if(i===2){ fill(c, GREEN); c.font = GREENTXT_BOLD; }
      else if(i===3){ fill(c, RED); c.font = REDTXT_BOLD; }
    });
    r++;
  });
  const totalVals = ['TOTAL', grand.total, grand.withinSla, grand.escalated, ...grand.buckets];
  totalVals.forEach((v,i)=>{
    const c = dash.getCell(r, i+1);
    c.value = v; c.font = WHITE_BOLD; centre(c); fill(c, NAVY); border(c);
  });

  /* ---------------- Officers Wise sheet ---------------- */
  const off = wb.addWorksheet('Officers Wise');
  off.columns = [{width:7},{width:36},{width:13},{width:11},{width:11},{width:11},{width:11},{width:14}];

  off.mergeCells('A1:H1');
  const oTitle = off.getCell('A1');
  oTitle.value = 'OFFICER-WISE PENDENCY AS ON ' + reportDateStr.toUpperCase();
  oTitle.font = {...WHITE_BOLD, size:15}; centre(oTitle); fill(oTitle, NAVY);
  off.getRow(1).height = 24;

  let orow = 3;
  const offHeaders = ['S.No.','Officer','Total Pending','0\u201315 Days','16\u201330 Days','31\u201345 Days','46\u201390 Days','More than 90 Days'];
  deptList.forEach(d=>{
    off.mergeCells(`A${orow}:H${orow}`);
    const bar = off.getCell(orow,1);
    bar.value = d.name; bar.font = WHITE_BOLD; bar.alignment = {horizontal:'left', vertical:'middle', indent:1}; fill(bar, NAVY); border(bar);
    orow++;

    offHeaders.forEach((h,i)=>{
      const c = off.getCell(orow, i+1);
      c.value = h; c.font = WHITE_BOLD; centre(c, true); fill(c, NAVY_DARK); border(c);
    });
    orow++;

    const officers = offByDept[d.name] || [];
    officers.forEach((o, idx)=>{
      const vals = [idx+1, o.officer, o.total, ...o.buckets];
      vals.forEach((v,i)=>{
        const c = off.getCell(orow, i+1);
        c.value = v; border(c);
        c.alignment = i===1 ? {horizontal:'left', vertical:'middle'} : {horizontal:'center', vertical:'middle'};
      });
      orow++;
    });

    const totals = officers.reduce((acc,o)=>{acc.total+=o.total; for(let i=0;i<5;i++) acc.b[i]+=o.buckets[i]; return acc;}, {total:0,b:[0,0,0,0,0]});
    off.mergeCells(`A${orow}:B${orow}`);
    const totLbl = off.getCell(orow,1);
    totLbl.value = 'TOTAL'; totLbl.font = WHITE_BOLD; centre(totLbl); fill(totLbl, NAVY); border(totLbl);
    [totals.total, ...totals.b].forEach((v,i)=>{
      const c = off.getCell(orow, i+3);
      c.value = v; c.font = WHITE_BOLD; centre(c); fill(c, NAVY); border(c);
    });
    orow += 2; // blank spacer row before next department
  });

  /* ---------------- Source Data sheet (raw, unchanged) ---------------- */
  const src = wb.addWorksheet('Source Data');
  if(currentRows.length){
    const cols = Object.keys(currentRows[0]);
    cols.forEach((c,i)=>{ src.getColumn(i+1).width = Math.min(28, Math.max(11, c.length+2)); });
    const headerRow = src.getRow(1);
    cols.forEach((h,i)=>{ headerRow.getCell(i+1).value = h; });
    headerRow.eachCell(c=>{ c.font = WHITE_BOLD; fill(c, NAVY_DARK); centre(c); });
    currentRows.forEach((row,ri)=>{
      const excelRow = src.getRow(ri+2);
      cols.forEach((h,i)=>{ excelRow.getCell(i+1).value = row[h]; });
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
