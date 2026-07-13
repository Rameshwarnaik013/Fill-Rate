from flask import Flask, request, jsonify, send_file, Response
import pandas as pd
import numpy as np
from io import BytesIO

app = Flask(__name__)

INSTAMART_CUSTOMERS = {
    "PJTJ Technologies Private Limited",
    "Cloudstore Retail Private Limited",
    "Moksh Enterprises Private Limited",
    "Jupiter Kart Private Limited",
    "Cloudkart Ventures Private Limited",
}

CUSTOMER_RENAMES = {
    "FK-Grocery-VS48860":   "FK-Grocery",
    "FK-Hyperlocal-VS46867": "FK-Hyperlocal",
    "FK-Alpha-VS46867":     "FK-Alpha",
}


# ── HTML (single-page app) ─────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fill Rate Dashboard</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size:13px; }
  .tab-btn { border-bottom: 3px solid transparent; transition: all .15s; }
  .tab-btn.active { border-bottom-color: #2563eb; color: #2563eb; }
  .tab-btn:hover:not(.active) { color: #374151; border-bottom-color: #d1d5db; }
  .drop-zone { border: 2px dashed #cbd5e1; transition: all .2s; }
  .drop-zone.drag-over { border-color: #2563eb; background: #eff6ff; }
  tr.grand-total td { background: #cbd5e1 !important; font-weight: 800; color: #0f172a; }
  tbody tr:not(.grand-total):not(.child-row):hover td { background: #f8fafc; }
  thead th { position: sticky; top: 0; z-index: 1; }
  tr.expandable-row { cursor: pointer; }
  tr.expandable-row:hover td { background: #EFF6FF !important; }
  tr.child-row td { background: #F0F9FF; }
  tr.child-row:hover td { background: #DBEAFE !important; }
  tr[data-cg-customer] { cursor: pointer; }
  tr[data-cg-customer]:hover td { background: #BAE6FD !important; }
  tr.grandchild-row td { background: #E0F2FE; }
  tr.grandchild-row:hover td { background: #BAE6FD !important; }
  /* compact table */
  table td, table th { white-space: nowrap; }
</style>
</head>
<body class="bg-gray-50 min-h-screen">

<!-- Header -->
<header class="bg-white border-b shadow-sm">
  <div class="max-w-screen-xl mx-auto px-4 py-3 flex items-center gap-3">
    <span class="text-2xl">&#128230;</span>
    <h1 class="text-xl font-semibold text-gray-800">Fill Rate Dashboard</h1>
    <span class="ml-auto text-xs text-gray-400" id="file-badge"></span>
  </div>
</header>

<div class="max-w-screen-xl mx-auto px-4 py-3 space-y-3">

  <!-- Hidden file input (OUTSIDE any clickable container to avoid click-loop bugs) -->
  <input type="file" id="file-input" accept=".xlsx,.xls"
         style="position:absolute;width:1px;height:1px;opacity:0;overflow:hidden;pointer-events:none;">

  <!-- Drop zone: a <label> natively opens the file picker on click — no JS needed -->
  <label for="file-input" id="drop-zone"
         style="display:block;background:#fff;border:2px dashed #cbd5e1;border-radius:16px;
                padding:48px 24px;text-align:center;cursor:pointer;transition:all .2s;user-select:none;">
    <div style="font-size:40px;margin-bottom:10px;">&#128193;</div>
    <p style="color:#6b7280;font-size:14px;margin:0;">Click to upload or drag &amp; drop your Fill Rate Excel file</p>
    <p style="color:#9ca3af;font-size:12px;margin:6px 0 0 0;">.xlsx / .xls</p>
  </label>

  <!-- File-ready bar — shown after a file is chosen -->
  <div id="file-ready-bar" style="display:none;background:#fff;border:1px solid #bfdbfe;
       border-radius:16px;padding:16px 20px;box-shadow:0 1px 4px rgba(0,0,0,.06);">
    <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
      <span style="font-size:26px;">&#128196;</span>
      <div style="flex:1;min-width:0;">
        <p id="selected-file-name"
           style="margin:0;font-size:14px;font-weight:700;color:#1e293b;
                  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"></p>
        <p id="upload-msg" style="margin:4px 0 0 0;font-size:12px;color:#6b7280;"></p>
      </div>
      <button onclick="resetFileSelection()"
        style="font-size:12px;color:#3b82f6;text-decoration:underline;background:none;
               border:none;cursor:pointer;white-space:nowrap;padding:4px 8px;">
        &#x21BA; Change file
      </button>
      <button id="generate-btn" onclick="triggerGenerate()"
        style="padding:11px 32px;background:#2563eb;color:#fff;font-size:14px;font-weight:700;
               border:none;border-radius:12px;cursor:pointer;display:inline-flex;align-items:center;
               gap:8px;box-shadow:0 3px 10px rgba(37,99,235,.35);white-space:nowrap;letter-spacing:.01em;">
        &#9654;&nbsp; Generate Fill Rate
      </button>
    </div>
  </div>

  <!-- Dashboard (hidden until data loaded) -->
  <div id="dashboard" class="hidden space-y-5">

    <!-- Filters -->
    <div class="bg-white rounded-2xl shadow-sm px-5 py-4 flex flex-wrap gap-4 items-end">
      <div>
        <label class="block text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-1">Origin</label>
        <select id="f-origin" onchange="applyFilters()"
          class="border border-gray-200 rounded-lg px-3 py-1.5 text-sm min-w-[130px] focus:outline-none focus:ring-2 focus:ring-blue-400"></select>
      </div>
      <div>
        <label class="block text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-1">Date From</label>
        <input type="date" id="f-date-from" onchange="applyFilters()"
          class="border border-gray-200 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
      </div>
      <div>
        <label class="block text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-1">Date To</label>
        <input type="date" id="f-date-to" onchange="applyFilters()"
          class="border border-gray-200 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
      </div>
      <div>
        <label class="block text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-1">Customer Group</label>
        <select id="f-cg" onchange="applyFilters()"
          class="border border-gray-200 rounded-lg px-3 py-1.5 text-sm min-w-[160px] focus:outline-none focus:ring-2 focus:ring-blue-400"></select>
      </div>
      <button onclick="resetFilters()"
        class="ml-auto px-4 py-1.5 text-sm rounded-lg bg-gray-100 hover:bg-gray-200 text-gray-600 transition-colors">
        &#x21BA; Reset
      </button>
    </div>

    <!-- KPI cards -->
    <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-4" id="kpi-cards"></div>

    <!-- Table card -->
    <div class="bg-white rounded-2xl shadow-sm overflow-hidden">
      <!-- Tabs -->
      <div class="flex border-b border-gray-200 overflow-x-auto px-2 pt-1">
        <button class="tab-btn active px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('item-group',this)">&#128193; Item Group</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('date',this)">&#128197; By Date</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('customer-group',this)">&#128101; Customer Group</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('customer',this)">&#127978; Customer</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('origin',this)">&#127758; Origin</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('closed-date',this)">&#128274; Closed Kgs</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('closed-remark',this)">&#128221; Item Close Remark</button>
        <button class="tab-btn px-5 py-3 text-sm font-medium text-gray-500 whitespace-nowrap"
                onclick="switchTab('closed-parent',this)">&#128202; Parent wise - Closed Kgs</button>
      </div>

      <!-- Table -->
      <div class="overflow-x-auto">
        <table class="w-full border-collapse" style="font-size:13px;font-weight:500;">
          <thead id="table-head">
            <tr class="bg-slate-100 border-b-2 border-slate-300">
              <th class="px-2 py-2 text-left text-xs font-bold text-slate-800 uppercase tracking-wide" id="col-label">Item Group</th>
              <th class="px-2 py-2 text-right text-xs font-bold text-slate-800 uppercase tracking-wide">Stock (KGS)</th>
              <th class="px-2 py-2 text-right text-xs font-bold text-slate-800 uppercase tracking-wide">Delivered (KGS)</th>
              <th class="px-2 py-2 text-right text-xs font-bold text-slate-800 uppercase tracking-wide">Closed KGS</th>
              <th class="px-2 py-2 text-right text-xs font-bold text-slate-800 uppercase tracking-wide">Pending KGS</th>
              <th class="px-2 py-2 text-center text-xs font-bold text-slate-800 uppercase tracking-wide">Fill Rate %</th>
              <th class="px-2 py-2 text-center text-xs font-bold text-slate-800 uppercase tracking-wide">Closed %</th>
              <th class="px-2 py-2 text-center text-xs font-bold text-slate-800 uppercase tracking-wide">Pen Dis %</th>
            </tr>
          </thead>
          <tbody id="table-body"></tbody>
        </table>
      </div>
    </div>

    <!-- Charts -->
    <div id="charts-card" class="bg-white rounded-2xl shadow-sm p-4 grid grid-cols-1 lg:grid-cols-2 gap-6">
      <div>
        <p class="text-xs font-bold text-slate-600 uppercase tracking-wide mb-2">Fill Rate % by Group</p>
        <div style="position:relative;height:280px"><canvas id="chart-fillrate"></canvas></div>
      </div>
      <div>
        <p class="text-xs font-bold text-slate-600 uppercase tracking-wide mb-2">Volume Breakdown (KGS)</p>
        <div style="position:relative;height:280px"><canvas id="chart-volume"></canvas></div>
      </div>
    </div>

    <!-- Bottom bar -->
    <div class="flex flex-wrap items-center gap-4">
      <button onclick="downloadExcel()" id="dl-btn"
        class="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-xl transition-colors flex items-center gap-2">
        &#128229; Download Excel Report
      </button>

      <!-- Legend -->
      <div class="ml-auto bg-white rounded-xl shadow-sm px-4 py-2 flex flex-wrap gap-5 text-xs text-gray-600">
        <span><b>Fill Rate:</b>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#70AD47">&#8805;90%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FFD966">70-89%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FF6B6B">&lt;70%</span>
        </span>
        <span><b>Closed %:</b>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#70AD47">0-5%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FFD966">6-10%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FF6B6B">&gt;10%</span>
        </span>
        <span><b>Pen Dis %:</b>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#70AD47">0-5%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FFD966">6-15%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FF6B6B">&gt;15%</span>
        </span>
      </div>
    </div>

  </div><!-- /dashboard -->
</div><!-- /container -->

<script>
// ── State ─────────────────────────────────────────────────────────────────────
let allRows = [];
let selectedFile = null;
let parsedCsvData = null;   // CSV string produced by SheetJS in-browser
let currentTab = 'item-group';
let expandedGroups      = new Set();  // Item Group → Parent Item
let expandedCGRows      = new Set();  // Customer Group → Customer
let expandedCGCustomers = new Set();  // Customer Group > Amazon/Flipkart → Client Type
let expandedCustomers   = new Set();  // Customer tab: Amazon / Flipkart → sub-channels

const TABS = {
  'item-group':      { key: 'NEW MIS ITEM GROUP', label: 'Item Group' },
  'date':            { key: 'Sales Order Date',   label: 'SO Date'    },
  'customer-group':  { key: 'Customer Group',     label: 'Customer Group' },
  'customer':        { key: 'Customer',           label: 'Customer'   },
  'origin':          { key: 'Origin',             label: 'Origin'     },
};

// Pivot tabs: sum of Closed Kgs, separate table layout (standard tabs untouched)
const PIVOT_TABS = {
  'closed-date':   { row: 'Sales Order Date',  label: 'SO Date' },
  'closed-remark': { row: 'Item Close Remark', label: 'Closed Remarks' },
  'closed-parent': { row: 'Parent Item',       label: 'Parent Name', col: 'Item Close Remark' },
};

// ── Data-processing constants (mirrors Python server logic exactly) ────────────
const INSTAMART_CUSTOMERS_JS = new Set([
  "PJTJ Technologies Private Limited","Cloudstore Retail Private Limited",
  "Moksh Enterprises Private Limited","Jupiter Kart Private Limited",
  "Cloudkart Ventures Private Limited"
]);
const CUSTOMER_RENAMES_JS = {
  "FK-Grocery-VS48860":"FK-Grocery",
  "FK-Hyperlocal-VS46867":"FK-Hyperlocal",
  "FK-Alpha-VS46867":"FK-Alpha"
};
const REQUIRED_COLS = ["Stock Qty in KGS","Delivered Qty (Kgs)","Closed Kgs",
                       "Pending Dispatch Kgs","NEW MIS ITEM GROUP"];
const KEEP_COLS = ["NEW MIS ITEM GROUP","Parent Item","Sales Order Date",
                   "Customer Group","Customer","Client Type","Origin","Item Close Remark",
                   "Stock Qty in KGS","Delivered Qty (Kgs)","Closed Kgs","Pending Dispatch Kgs"];

// ── Upload / Drag & Drop ──────────────────────────────────────────────────────
// The drop zone is a <label for="file-input"> so clicking it natively opens the
// file picker without any JS. We only need JS for drag-and-drop and change events.
const dz = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

// Drag & drop on the label/drop zone
dz.addEventListener('dragover', e => {
  e.preventDefault();
  dz.style.borderColor = '#2563eb';
  dz.style.background  = '#eff6ff';
});
dz.addEventListener('dragleave', () => {
  dz.style.borderColor = '#cbd5e1';
  dz.style.background  = '#fff';
});
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.style.borderColor = '#cbd5e1';
  dz.style.background  = '#fff';
  const file = e.dataTransfer.files[0];
  if (file) previewFile(file);
});

// File chosen via native dialog
fileInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) previewFile(file);
});

async function previewFile(file) {
  selectedFile = file;
  parsedCsvData = null;
  allRows = [];
  const mb = (file.size / 1024 / 1024).toFixed(1);
  document.getElementById('selected-file-name').textContent = file.name + '  (' + mb + ' MB)';
  setMsg('⏳ Reading Excel…');
  dz.style.display = 'none';
  document.getElementById('file-ready-bar').style.display = 'block';
  // Disable Generate button while processing
  const btn = document.getElementById('generate-btn');
  btn.disabled = true; btn.style.opacity = '0.55'; btn.innerHTML = '⏳ Processing…';

  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws, { raw: false, dateNF: 'yyyy-mm-dd', defval: '' });

    if (!rawRows.length) { setMsg('⚠ File appears empty', true); resetBtn(); return; }

    // Build a canonical→actual header map (case/space-insensitive)
    const fileHeaders = Object.keys(rawRows[0]).map(h => h.trim());
    const hMap = {};
    KEEP_COLS.forEach(canon => {
      hMap[canon] =
        fileHeaders.find(h => h === canon) ||
        fileHeaders.find(h => h.toLowerCase() === canon.toLowerCase()) ||
        fileHeaders.find(h => h.toLowerCase().replace(/\\s+/g,'') === canon.toLowerCase().replace(/\\s+/g,'')) ||
        null;
    });

    // Validate required columns
    const missing = REQUIRED_COLS.filter(r => !hMap[r]);
    if (missing.length) {
      setMsg('❌ Missing columns: ' + missing.join(', ') +
             '<br><small style="color:#9ca3af">Found: ' + fileHeaders.join(', ') + '</small>', true);
      resetBtn(); return;
    }

    const numCols = ['Stock Qty in KGS','Delivered Qty (Kgs)','Closed Kgs','Pending Dispatch Kgs'];

    // Process every row entirely in the browser
    allRows = rawRows
      .map(raw => {
        const r = {};
        for (const canon of KEEP_COLS) {
          r[canon] = hMap[canon] ? (raw[hMap[canon]] !== undefined ? raw[hMap[canon]] : '') : '';
        }
        // Numeric coerce
        for (const c of numCols) r[c] = parseFloat(r[c]) || 0;
        // Instamart override
        const cust = String(r['Customer'] || '').trim();
        if (INSTAMART_CUSTOMERS_JS.has(cust)) r['Customer Group'] = 'Instamart';
        // Customer rename
        r['Customer'] = CUSTOMER_RENAMES_JS[cust] || cust;
        // Origin taken directly from the file's "Origin" column
        r['Origin'] = String(r['Origin'] || '').trim();
        return r;
      })
      .filter(r => String(r['NEW MIS ITEM GROUP'] || '').trim() !== '');

    setMsg('✓ ' + allRows.length.toLocaleString() + ' rows ready — click Generate Fill Rate');

  } catch (e) {
    console.error('SheetJS error:', e);
    allRows = [];
    setMsg('❌ Could not read file: ' + e.message, true);
  }
  resetBtn();
}

function resetFileSelection() {
  selectedFile = null;
  parsedCsvData = null;
  fileInput.value = '';
  document.getElementById('file-ready-bar').style.display = 'none';
  dz.style.display = 'block';
  setMsg('');
}

function setMsg(html, err = false) {
  const el = document.getElementById('upload-msg');
  el.innerHTML = html;
  el.style.color = err ? '#ef4444' : '#6b7280';
}

function resetBtn() {
  const btn = document.getElementById('generate-btn');
  btn.disabled = false;
  btn.style.background = '#2563eb';
  btn.style.opacity = '1';
  btn.style.cursor = 'pointer';
  btn.innerHTML = '&#9654;&nbsp; Generate Fill Rate';
}

function triggerGenerate() {
  if (!selectedFile) return;

  if (allRows.length > 0) {
    // ✅ Data already processed in browser — show dashboard instantly, no upload
    document.getElementById('file-badge').textContent =
      selectedFile.name + ' · ' + allRows.length.toLocaleString() + ' rows';
    document.getElementById('file-ready-bar').style.display = 'none';
    document.getElementById('dashboard').classList.remove('hidden');
    initFilters();
    applyFilters();
    return;
  }

  // Fallback: if SheetJS failed, upload file to server
  const btn = document.getElementById('generate-btn');
  btn.disabled = true; btn.style.background = '#1e40af';
  btn.style.opacity = '0.8'; btn.style.cursor = 'not-allowed';
  btn.innerHTML = '&#9203;&nbsp; Uploading…';
  uploadFile(selectedFile).finally(resetBtn);
}

async function uploadFile(file) {
  const mb = (file.size / 1024 / 1024).toFixed(1);
  const fd = new FormData();

  if (parsedCsvData) {
    // Upload the slim CSV extracted in-browser — far smaller than the raw Excel
    const kb = Math.round(new Blob([parsedCsvData]).size / 1024);
    setMsg('Uploading ' + kb + ' KB CSV (from ' + mb + ' MB Excel)…');
    fd.append('csv_data', parsedCsvData);
  } else {
    // Fallback: upload the original file
    setMsg('Uploading ' + file.name + ' (' + mb + ' MB)…');
    fd.append('file', file);
  }
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: fd });

    // Guard: server may return HTML (413 / 500) instead of JSON
    const ct = res.headers.get('content-type') || '';
    if (!ct.includes('application/json')) {
      const text = await res.text();
      if (res.status === 413)
        setMsg('❌ File too large (' + mb + ' MB). Vercel limit is 4.5 MB — try a smaller date range.', true);
      else
        setMsg('❌ Server error HTTP ' + res.status + ': ' + text.slice(0, 150), true);
      return;
    }

    const data = await res.json();
    if (!res.ok) {
      // Format column errors with line breaks
      const msg = (data.error || 'Upload failed').replace(/\\n/g, '<br>');
      const extra = data.columns_in_file
        ? '<br><span style="font-size:11px;color:#6b7280">Columns found in file: ' +
          data.columns_in_file.join(', ') + '</span>'
        : '';
      setMsg('❌ ' + msg + extra, true);
      return;
    }

    allRows = data.rows;
    document.getElementById('file-badge').textContent = file.name + ' · ' + allRows.length.toLocaleString() + ' rows';
    // Hide the ready bar, show dashboard FIRST so canvases have dimensions before Chart.js renders
    document.getElementById('file-ready-bar').style.display = 'none';
    document.getElementById('dashboard').classList.remove('hidden');
    initFilters();
    applyFilters();
    setMsg('✓ Loaded ' + allRows.length.toLocaleString() + ' rows from ' + file.name);
  } catch (err) {
    setMsg('❌ ' + err.message, true);
    console.error('Upload error:', err);
  }
}

// ── Filters ───────────────────────────────────────────────────────────────────
function initFilters() {
  const uniq = col => [...new Set(allRows.map(r => r[col]).filter(Boolean))].sort();
  populateSel('f-origin', uniq('Origin'));
  populateSel('f-cg', uniq('Customer Group'));
  const dates = allRows.map(r => r['Sales Order Date']).filter(Boolean).sort();
  if (dates.length) {
    document.getElementById('f-date-from').value = dates[0];
    document.getElementById('f-date-to').value = dates[dates.length - 1];
  }
}

function populateSel(id, opts) {
  document.getElementById(id).innerHTML =
    '<option value="">All</option>' + opts.map(o => `<option>${o}</option>`).join('');
}

function resetFilters() {
  document.getElementById('f-origin').value = '';
  document.getElementById('f-cg').value = '';
  const dates = allRows.map(r => r['Sales Order Date']).filter(Boolean).sort();
  if (dates.length) {
    document.getElementById('f-date-from').value = dates[0];
    document.getElementById('f-date-to').value = dates[dates.length - 1];
  }
  applyFilters();
}

function filteredRows() {
  const origin = document.getElementById('f-origin').value;
  const cg     = document.getElementById('f-cg').value;
  const dFrom  = document.getElementById('f-date-from').value;
  const dTo    = document.getElementById('f-date-to').value;
  return allRows.filter(r => {
    if (origin && r['Origin'] !== origin) return false;
    if (cg && r['Customer Group'] !== cg) return false;
    if (dFrom && r['Sales Order Date'] && r['Sales Order Date'] < dFrom) return false;
    if (dTo   && r['Sales Order Date'] && r['Sales Order Date'] > dTo)   return false;
    return true;
  });
}

function applyFilters() {
  const rows = filteredRows();
  renderKPIs(rows);
  renderTable(currentTab, rows);
  try { renderCharts(currentTab, rows); } catch(e) { console.error('Chart render error:', e); }
}

// ── KPI Cards ─────────────────────────────────────────────────────────────────
function renderKPIs(rows) {
  const ts = sumCol(rows, 'Stock Qty in KGS');
  const td = sumCol(rows, 'Delivered Qty (Kgs)');
  const tc = sumCol(rows, 'Closed Kgs');
  const tp = sumCol(rows, 'Pending Dispatch Kgs');
  const fr = ts ? td / ts : 0;
  const cp = ts ? tc / ts : 0;
  const pp = ts ? tp / ts : 0;

  const cards = [
    { label: 'Total Stock (KGS)',        val: fmtN(ts), bg: '#EFF6FF', fg: '#1e40af', lb: '#1e3a8a' },
    { label: 'Total Delivered (KGS)',    val: fmtN(td), bg: '#ECFDF5', fg: '#065f46', lb: '#064e3b' },
    { label: 'Closed KGS',              val: fmtN(tc), bg: '#FFF7ED', fg: '#9a3412', lb: '#7c2d12' },
    { label: 'Pending Dispatch KGS',    val: fmtN(tp), bg: '#FAF5FF', fg: '#7e22ce', lb: '#6b21a8' },
    { label: 'Fill Rate',               val: pct(fr),  bg: fillClrAbs(fr), fg: '#000', lb: '#1f2937' },
    { label: 'Closed %',                val: pct(cp),  bg: closedClr(cp), fg: '#000', lb: '#1f2937' },
    { label: 'Pen Dis %',               val: pct(pp),  bg: penClr(pp),    fg: '#000', lb: '#1f2937' },
  ];
  document.getElementById('kpi-cards').innerHTML = cards.map(c => `
    <div class="rounded-2xl shadow-sm border border-gray-200 p-4" style="background:${c.bg}">
      <p class="text-xs font-bold uppercase tracking-wider mb-1" style="color:${c.lb}">${c.label}</p>
      <p class="text-3xl font-extrabold" style="color:${c.fg}">${c.val}</p>
    </div>`).join('');
}

// ── Tab switch ────────────────────────────────────────────────────────────────
function switchTab(tab, btn) {
  currentTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const rows = filteredRows();
  renderTable(tab, rows);   // restores/replaces thead, so set col-label after
  const colLabel = document.getElementById('col-label');
  if (colLabel && TABS[tab]) colLabel.textContent = TABS[tab].label;
  renderCharts(tab, rows);
}

// ── Aggregation ───────────────────────────────────────────────────────────────
function aggregate(rows, key) {
  const g = {};
  rows.forEach(r => {
    const k = r[key];
    // skip rows where the group key is blank, null, undefined, or NaN
    if (k == null || k === '' || k !== k) return;
    const ks = String(k).trim();
    if (!ks) return;
    if (!g[ks]) g[ks] = { stock: 0, delivered: 0, closed: 0, pending: 0 };
    g[ks].stock     += r['Stock Qty in KGS']        || 0;
    g[ks].delivered += r['Delivered Qty (Kgs)']     || 0;
    g[ks].closed    += r['Closed Kgs']              || 0;
    g[ks].pending   += r['Pending Dispatch Kgs']    || 0;
  });

  let result = Object.entries(g).map(([label, v]) => ({
    label,
    stock: v.stock, delivered: v.delivered, closed: v.closed, pending: v.pending,
    fillRate:   v.stock ? v.delivered / v.stock : 0,
    closedPct:  v.stock ? v.closed    / v.stock : 0,
    penDis:     v.stock ? v.pending   / v.stock : 0,
  }));

  // date tab → chrono; others → stock descending
  if (key === 'Sales Order Date') {
    result.sort((a, b) => a.label < b.label ? -1 : 1);
  } else {
    result.sort((a, b) => b.stock - a.stock);
  }

  const tS = result.reduce((s, r) => s + r.stock, 0);
  const tD = result.reduce((s, r) => s + r.delivered, 0);
  const tC = result.reduce((s, r) => s + r.closed, 0);
  const tP = result.reduce((s, r) => s + r.pending, 0);
  result.push({
    label: 'Grand Total', isTotal: true,
    stock: tS, delivered: tD, closed: tC, pending: tP,
    fillRate:  tS ? tD / tS : 0,
    closedPct: tS ? tC / tS : 0,
    penDis:    tS ? tP / tS : 0,
  });
  return result;
}

// ── Table rendering ───────────────────────────────────────────────────────────
function metricCells(r, size) {
  const p = size === 'sm' ? 'px-2 py-0.5' : 'px-2 py-1';
  return `
    <td class="${p} text-right text-gray-900 font-semibold">${fmtN(r.stock)}</td>
    <td class="${p} text-right text-gray-900 font-semibold">${fmtN(r.delivered)}</td>
    <td class="${p} text-right text-gray-900 font-semibold">${r.closed  > 0 ? fmtN(r.closed)  : '-'}</td>
    <td class="${p} text-right text-gray-900 font-semibold">${r.pending > 0 ? fmtN(r.pending) : '-'}</td>
    <td class="${p} text-center font-extrabold" style="background:${fillClr(r.fillRate)};color:#1a1a1a">${pct(r.fillRate)}</td>
    <td class="${p} text-center font-extrabold" style="background:${closedClr(r.closedPct)};color:#1a1a1a">${pct(r.closedPct)}</td>
    <td class="${p} text-center font-extrabold" style="background:${penClr(r.penDis)};color:#1a1a1a">${pct(r.penDis)}</td>`;
}

// ── Customer tab: group Amazon & Flipkart sub-channels under parent ───────────
function getCustomerParent(name) {
  if (!name) return null;
  if (name.startsWith('Amazon Retail India'))              return 'Amazon Retail India';
  if (name.startsWith('Flipkart India Private Limited'))   return 'Flipkart India Private Limited';
  return null;
}

function aggregateCustomersTab(rows) {
  const g = {};
  const toAcc = () => ({ stock: 0, delivered: 0, closed: 0, pending: 0 });

  rows.forEach(r => {
    const customer = (r['Customer'] || '').trim();
    if (!customer) return;
    const parent = getCustomerParent(customer);
    const key = parent || customer;

    if (!g[key]) g[key] = { ...toAcc(), isParent: !!parent, children: {} };
    g[key].stock     += r['Stock Qty in KGS']     || 0;
    g[key].delivered += r['Delivered Qty (Kgs)']  || 0;
    g[key].closed    += r['Closed Kgs']           || 0;
    g[key].pending   += r['Pending Dispatch Kgs'] || 0;

    if (parent) {
      if (!g[key].children[customer]) g[key].children[customer] = toAcc();
      g[key].children[customer].stock     += r['Stock Qty in KGS']     || 0;
      g[key].children[customer].delivered += r['Delivered Qty (Kgs)']  || 0;
      g[key].children[customer].closed    += r['Closed Kgs']           || 0;
      g[key].children[customer].pending   += r['Pending Dispatch Kgs'] || 0;
    }
  });

  const toRow = (label, v) => ({
    label,
    stock: v.stock, delivered: v.delivered, closed: v.closed, pending: v.pending,
    fillRate:  v.stock ? v.delivered / v.stock : 0,
    closedPct: v.stock ? v.closed    / v.stock : 0,
    penDis:    v.stock ? v.pending   / v.stock : 0,
  });

  let result = Object.entries(g).map(([label, v]) => ({
    ...toRow(label, v),
    isParent: v.isParent,
    children: Object.entries(v.children || {})
      .map(([cl, cv]) => toRow(cl, cv))
      .sort((a, b) => b.stock - a.stock),
  })).sort((a, b) => b.stock - a.stock);

  const tS = result.reduce((s, r) => s + r.stock, 0);
  const tD = result.reduce((s, r) => s + r.delivered, 0);
  const tC = result.reduce((s, r) => s + r.closed, 0);
  const tP = result.reduce((s, r) => s + r.pending, 0);
  result.push({
    ...toRow('Grand Total', { stock: tS, delivered: tD, closed: tC, pending: tP }),
    isTotal: true, isParent: false, children: [],
  });
  return result;
}

// ── Expandable row helpers ────────────────────────────────────────────────────
function expandableRow(attr, label, r, expanded) {
  return `<tr class="expandable-row border-t border-gray-200" ${attr}="${esc(label)}">
    <td class="px-2 py-1 text-gray-900 font-semibold">
      <span class="inline-block w-4 text-blue-500 text-[10px] select-none">${expanded ? '&#9660;' : '&#9654;'}</span>${label}
    </td>${metricCells(r, 'md')}</tr>`;
}

function childRow(label, r) {
  return `<tr class="child-row border-t border-blue-200">
    <td class="px-2 py-0.5 text-blue-900 pl-7 font-semibold">&#8627; ${label}</td>
    ${metricCells(r, 'sm')}</tr>`;
}

function regularRow(r) {
  return `<tr class="${r.isTotal ? 'grand-total' : ''} border-t border-gray-200">
    <td class="px-2 py-1 text-gray-900 font-semibold">${r.label}</td>
    ${metricCells(r, 'md')}</tr>`;
}

// ── Pivot tabs (Closed Kgs) ───────────────────────────────────────────────────
let DEFAULT_THEAD = null;   // captured on first swap so standard tabs restore exactly

const fmtD = n => n.toLocaleString('en-IN', { maximumFractionDigits: 3 });
const thP  = (txt, align) =>
  `<th class="px-2 py-2 text-${align} text-xs font-bold text-slate-800 uppercase tracking-wide">${txt}</th>`;

function renderPivotTable(tab, rows) {
  const cfg = PIVOT_TABS[tab];
  const thead = document.getElementById('table-head');
  if (DEFAULT_THEAD === null) DEFAULT_THEAD = thead.innerHTML;

  // Only rows that actually have closed quantity
  const closedRows = rows.filter(r => (r['Closed Kgs'] || 0) > 0);
  const keyOf = r => String(r[cfg.row] || '').trim() || '(Blank)';

  let headHtml, bodyHtml = '';

  if (cfg.col) {
    // ── Matrix pivot: rows × remark columns, cells = sum of Closed Kgs ────────
    const colOf = r => String(r[cfg.col] || '').trim() || '(Blank)';
    const g = {}, colTotals = {};
    let grand = 0;
    closedRows.forEach(r => {
      const rk = keyOf(r), ck = colOf(r), v = r['Closed Kgs'] || 0;
      if (!g[rk]) g[rk] = { total: 0, cells: {} };
      g[rk].cells[ck] = (g[rk].cells[ck] || 0) + v;
      g[rk].total += v;
      colTotals[ck] = (colTotals[ck] || 0) + v;
      grand += v;
    });
    const cols = Object.keys(colTotals).sort((a, b) => colTotals[b] - colTotals[a]);
    const rws  = Object.entries(g).sort((a, b) => b[1].total - a[1].total);

    headHtml = `<tr class="bg-slate-100 border-b-2 border-slate-300">
      ${thP(cfg.label, 'left')}${cols.map(c => thP(esc(c), 'right')).join('')}${thP('Grand Total', 'right')}</tr>`;

    const cell = v => `<td class="px-2 py-1 text-right text-gray-900 font-semibold">${v ? fmtD(v) : ''}</td>`;
    rws.forEach(([label, { total, cells }]) => {
      bodyHtml += `<tr class="border-t border-gray-200">
        <td class="px-2 py-1 text-gray-900 font-semibold">${esc(label)}</td>
        ${cols.map(c => cell(cells[c] || 0)).join('')}
        <td class="px-2 py-1 text-right text-gray-900 font-extrabold">${fmtD(total)}</td></tr>`;
    });
    bodyHtml += `<tr class="grand-total border-t border-gray-200">
      <td class="px-2 py-1">Grand Total</td>
      ${cols.map(c => `<td class="px-2 py-1 text-right">${fmtD(colTotals[c])}</td>`).join('')}
      <td class="px-2 py-1 text-right">${fmtD(grand)}</td></tr>`;

  } else {
    // ── Simple pivot: key → sum of Closed Kgs ─────────────────────────────────
    const g = {};
    let grand = 0;
    closedRows.forEach(r => {
      const k = keyOf(r), v = r['Closed Kgs'] || 0;
      g[k] = (g[k] || 0) + v;
      grand += v;
    });
    let entries = Object.entries(g);
    entries = tab === 'closed-date'
      ? entries.sort((a, b) => a[0] < b[0] ? -1 : 1)   // dates ascending
      : entries.sort((a, b) => b[1] - a[1]);           // remarks by volume desc

    headHtml = `<tr class="bg-slate-100 border-b-2 border-slate-300">
      ${thP(cfg.label, 'left')}${thP('Sum of Closed Kgs', 'right')}</tr>`;

    entries.forEach(([label, v]) => {
      bodyHtml += `<tr class="border-t border-gray-200">
        <td class="px-2 py-1 text-gray-900 font-semibold">${esc(label)}</td>
        <td class="px-2 py-1 text-right text-gray-900 font-semibold">${fmtD(v)}</td></tr>`;
    });
    bodyHtml += `<tr class="grand-total border-t border-gray-200">
      <td class="px-2 py-1">Grand Total</td>
      <td class="px-2 py-1 text-right">${fmtD(grand)}</td></tr>`;
  }

  thead.innerHTML = headHtml;
  document.getElementById('table-body').innerHTML = bodyHtml;
}

// ── Table rendering ───────────────────────────────────────────────────────────
function renderTable(tab, rows) {
  if (PIVOT_TABS[tab]) { renderPivotTable(tab, rows); return; }
  // Standard tab: restore the original 8-column header if a pivot replaced it
  if (DEFAULT_THEAD !== null) {
    document.getElementById('table-head').innerHTML = DEFAULT_THEAD;
    DEFAULT_THEAD = null;
  }
  let html = '';

  if (tab === 'customer') {
    const custData = aggregateCustomersTab(rows);
    setFillScale(custData);
    for (const r of custData) {
      if (r.isTotal) { html += regularRow(r); continue; }
      if (r.isParent) {
        const exp = expandedCustomers.has(r.label);
        html += expandableRow('data-customer', r.label, r, exp);
        if (exp) r.children.forEach(c => { html += childRow(c.label, c); });
      } else {
        html += regularRow(r);
      }
    }
  } else {
    const data = aggregate(rows, TABS[tab].key);
    setFillScale(data);
    for (const r of data) {
      if (tab === 'item-group' && !r.isTotal) {
        const exp = expandedGroups.has(r.label);
        html += expandableRow('data-group', r.label, r, exp);
        if (exp) {
          aggregate(rows.filter(x => x['NEW MIS ITEM GROUP'] === r.label), 'Parent Item')
            .filter(s => !s.isTotal)
            .forEach(s => { html += childRow(s.label, s); });
        }
      } else if (tab === 'customer-group' && !r.isTotal) {
        const exp = expandedCGRows.has(r.label);
        html += expandableRow('data-cg', r.label, r, exp);
        if (exp) {
          const cgRows = rows.filter(x => x['Customer Group'] === r.label);
          aggregate(cgRows, 'Customer')
            .filter(s => !s.isTotal)
            .forEach(s => {
              const needsClientType = s.label.startsWith('Amazon Retail India') ||
                                      s.label.startsWith('Flipkart India Private Limited');
              if (needsClientType) {
                const ck = r.label + '||' + s.label;
                const cExp = expandedCGCustomers.has(ck);
                // Render as expandable child row
                html += `<tr class="child-row expandable-row border-t border-blue-200" data-cg-customer="${esc(ck)}">
                  <td class="px-2 py-0.5 text-blue-900 pl-7 font-semibold">
                    <span class="inline-block w-4 text-blue-500 text-[10px] select-none">${cExp ? '&#9660;' : '&#9654;'}</span>&#8627; ${s.label}
                  </td>${metricCells(s, 'sm')}</tr>`;
                if (cExp) {
                  aggregate(cgRows.filter(x => x['Customer'] === s.label), 'Client Type')
                    .filter(ct => !ct.isTotal)
                    .forEach(ct => {
                      html += `<tr class="grandchild-row border-t border-sky-200">
                        <td class="px-2 py-0.5 text-sky-900 pl-12 font-semibold">&#8627; ${ct.label || '(No Client Type)'}</td>
                        ${metricCells(ct, 'sm')}</tr>`;
                    });
                }
              } else {
                html += childRow(s.label, s);
              }
            });
        }
      } else {
        html += regularRow(r);
      }
    }
  }

  document.getElementById('table-body').innerHTML = html;
}

// ── Unified click handler for all expandable tabs ────────────────────────────
document.getElementById('table-body').addEventListener('click', e => {
  // Second-level toggle inside Customer Group: Amazon / Flipkart → Client Type
  if (currentTab === 'customer-group') {
    const trCC = e.target.closest('tr[data-cg-customer]');
    if (trCC) {
      const ck = trCC.dataset.cgCustomer;   // data-cg-customer → dataset.cgCustomer
      expandedCGCustomers.has(ck) ? expandedCGCustomers.delete(ck) : expandedCGCustomers.add(ck);
      renderTable(currentTab, filteredRows());
      return;
    }
  }
  const handlers = {
    'item-group':     ['tr[data-group]',    'group',    expandedGroups],
    'customer-group': ['tr[data-cg]',       'cg',       expandedCGRows],
    'customer':       ['tr[data-customer]', 'customer', expandedCustomers],
  };
  const h = handlers[currentTab];
  if (!h) return;
  const tr = e.target.closest(h[0]);
  if (!tr) return;
  const label = tr.dataset[h[1]];
  h[2].has(label) ? h[2].delete(label) : h[2].add(label);
  renderTable(currentTab, filteredRows());
});

// ── Download ──────────────────────────────────────────────────────────────────
async function downloadExcel() {
  const btn = document.getElementById('dl-btn');
  btn.textContent = 'Generating…'; btn.disabled = true;
  const rows = filteredRows();
  const sheets = {};
  for (const [, { key, label }] of Object.entries(TABS)) {
    sheets[label] = aggregate(rows, key).map(r => ({
      label: r.label, stock: r.stock, delivered: r.delivered,
      closed: r.closed, pending: r.pending,
      fillRate: r.fillRate, closedPct: r.closedPct, penDis: r.penDis,
    }));
  }
  try {
    const res = await fetch('/api/download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheets }),
    });
    if (!res.ok) { alert('Download failed'); return; }
    const blob = await res.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'fill_rate_dashboard.xlsx';
    a.click();
  } finally {
    btn.innerHTML = '&#128229; Download Excel Report'; btn.disabled = false;
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
const sumCol = (rows, col) => rows.reduce((s, r) => s + (r[col] || 0), 0);
const fmtN   = n => n.toLocaleString('en-IN', { maximumFractionDigits: 0 });
const pct    = v => Math.round(v * 100) + '%';
const esc    = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

// Absolute thresholds — used for the single overall KPI card
function fillClrAbs(v) { return v >= 0.90 ? '#70AD47' : v >= 0.70 ? '#FFD966' : '#FF6B6B'; }
function closedClr(v)  { return v <= 0.05 ? '#70AD47' : v <= 0.10 ? '#FFD966' : '#FF6B6B'; }
function penClr(v)     { return v <= 0.05 ? '#70AD47' : v <= 0.15 ? '#FFD966' : '#FF6B6B'; }

// Row-wise relative scale: highest fill rate in view → greenest, lowest → reddest
let _frScale = { min: 0, max: 1 };
function setFillScale(list) {
  const vals = list.filter(r => !r.isTotal).map(r => r.fillRate);
  _frScale = vals.length
    ? { min: Math.min(...vals), max: Math.max(...vals) }
    : { min: 0, max: 1 };
}
function _hexBlend(a, b, t) {
  const pa = [1, 3, 5].map(i => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map(i => parseInt(b.slice(i, i + 2), 16));
  return '#' + pa.map((v, i) =>
    Math.round(v + (pb[i] - v) * t).toString(16).padStart(2, '0')).join('');
}
function fillClr(v) {
  const { min, max } = _frScale;
  const t = max > min ? Math.max(0, Math.min(1, (v - min) / (max - min))) : 1;
  return t < 0.5 ? _hexBlend('#FF6B6B', '#FFD966', t * 2)
                 : _hexBlend('#FFD966', '#70AD47', (t - 0.5) * 2);
}

// ── Charts ────────────────────────────────────────────────────────────────────
let _chartFR = null, _chartVol = null;

function renderCharts(tab, rows) {
  // Pivot tabs have no fill-rate series — hide the charts card entirely
  const chartsCard = document.getElementById('charts-card');
  if (PIVOT_TABS[tab]) { chartsCard.style.display = 'none'; return; }
  chartsCard.style.display = '';
  const raw = tab === 'customer' ? aggregateCustomersTab(rows) : aggregate(rows, TABS[tab].key);
  const data = raw.filter(r => !r.isTotal).slice(0, 15);  // top 15 rows, skip grand total
  setFillScale(data);
  const labels = data.map(r => r.label.length > 20 ? r.label.slice(0,18)+'…' : r.label);

  // ── Fill Rate % horizontal bar ─────────────────────────────────────────────
  const frCtx = document.getElementById('chart-fillrate').getContext('2d');
  if (_chartFR) _chartFR.destroy();
  _chartFR = new Chart(frCtx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Fill Rate %',  data: data.map(r => +(r.fillRate*100).toFixed(1)),
          backgroundColor: data.map(r => fillClr(r.fillRate)), borderRadius: 3 },
        { label: 'Closed %',     data: data.map(r => +(r.closedPct*100).toFixed(1)),
          backgroundColor: data.map(r => closedClr(r.closedPct) + 'cc'), borderRadius: 3 },
        { label: 'Pen Dis %',    data: data.map(r => +(r.penDis*100).toFixed(1)),
          backgroundColor: data.map(r => penClr(r.penDis) + 'cc'), borderRadius: 3 },
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { size: 10 } } },
                 tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.x}%` } } },
      scales: {
        x: { max: 100, ticks: { callback: v => v+'%', font: { size: 10 } }, grid: { color: '#f1f5f9' } },
        y: { ticks: { font: { size: 10 } } }
      }
    }
  });

  // ── Volume stacked bar ─────────────────────────────────────────────────────
  const volCtx = document.getElementById('chart-volume').getContext('2d');
  if (_chartVol) _chartVol.destroy();
  _chartVol = new Chart(volCtx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Delivered', data: data.map(r => r.delivered), backgroundColor: '#70AD47cc', borderRadius: 2 },
        { label: 'Closed',    data: data.map(r => r.closed),    backgroundColor: '#FFD966cc', borderRadius: 2 },
        { label: 'Pending',   data: data.map(r => r.pending),   backgroundColor: '#FF6B6Bcc', borderRadius: 2 },
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { size: 10 } } },
                 tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${fmtN(ctx.parsed.x)} KGS` } } },
      scales: {
        x: { stacked: true, ticks: { font: { size: 10 } }, grid: { color: '#f1f5f9' } },
        y: { stacked: true, ticks: { font: { size: 10 } } }
      }
    }
  });
}
</script>
</body>
</html>"""


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return Response(HTML, mimetype="text/html")


@app.route("/api/upload", methods=["POST"])
def upload():
    import json as _json
    try:
        csv_data = request.form.get("csv_data")
        if csv_data:
            # Browser converted Excel → CSV via SheetJS before uploading
            from io import StringIO
            try:
                df = pd.read_csv(StringIO(csv_data))
            except Exception as e:
                return jsonify({"error": f"Could not parse CSV data: {e}"}), 400
        elif "file" in request.files:
            file = request.files["file"]
            try:
                df = pd.read_excel(file)
            except Exception as e:
                return jsonify({"error": f"Could not read Excel file: {e}"}), 400
        else:
            return jsonify({"error": "No file provided"}), 400

        # ── Normalise column names (strip leading/trailing whitespace) ──────────
        df.columns = [str(c).strip() for c in df.columns]
        available = list(df.columns)

        # ── Smart column validation ─────────────────────────────────────────────
        required = ["Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs",
                    "Pending Dispatch Kgs", "NEW MIS ITEM GROUP"]
        problems = []
        for req in required:
            if req in df.columns:
                continue
            # Case-insensitive exact match
            ci = next((c for c in available if c.lower() == req.lower()), None)
            if ci:
                problems.append(f"'{req}' → file has '{ci}' (case mismatch)")
                continue
            # Whitespace-stripped case-insensitive match
            ws = next((c for c in available
                       if c.lower().replace(" ", "") == req.lower().replace(" ", "")), None)
            if ws:
                problems.append(f"'{req}' → file has '{ws}' (spacing/case mismatch)")
                continue
            # Partial match (contains key word)
            part = next((c for c in available
                         if req.lower().split()[0] in c.lower()), None)
            hint = f" — closest: '{part}'" if part else ""
            problems.append(f"'{req}' — NOT FOUND in file{hint}")

        if problems:
            return jsonify({
                "error": "Column issues detected:\n• " + "\n• ".join(problems),
                "columns_in_file": available
            }), 400

        # ── Processing ──────────────────────────────────────────────────────────
        num_cols = ["Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs",
                    "Pending Dispatch Kgs"]
        for c in num_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        if "Sales Order Date" in df.columns:
            df["Sales Order Date"] = (
                pd.to_datetime(df["Sales Order Date"], errors="coerce")
                .dt.strftime("%Y-%m-%d")
            )

        # Origin comes straight from the file's own "Origin" column
        if "Origin" in df.columns:
            df["Origin"] = df["Origin"].apply(
                lambda v: str(v).strip() if pd.notna(v) else ""
            )
        else:
            df["Origin"] = ""

        # Override Customer Group for Instamart customers
        if "Customer" in df.columns:
            df["Customer Group"] = df.apply(
                lambda r: "Instamart"
                if str(r.get("Customer") or "").strip() in INSTAMART_CUSTOMERS
                else r.get("Customer Group"),
                axis=1,
            )
            # Rename specific customer names
            df["Customer"] = df["Customer"].apply(
                lambda v: CUSTOMER_RENAMES.get(str(v).strip(), v) if pd.notna(v) else v
            )

        keep = ["NEW MIS ITEM GROUP", "Parent Item", "Sales Order Date",
                "Customer Group", "Customer", "Client Type", "Origin",
                "Item Close Remark"] + num_cols
        keep = [c for c in keep if c in df.columns]

        rows = _json.loads(df[keep].to_json(orient="records"))
        return jsonify({"rows": rows})

    except Exception as e:
        import traceback
        return jsonify({"error": f"Processing error: {str(e)}",
                        "detail": traceback.format_exc()[-500:]}), 500


@app.route("/api/download", methods=["POST"])
def download():
    data = request.get_json(silent=True)
    if not data or "sheets" not in data:
        return jsonify({"error": "No data"}), 400

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, records in data["sheets"].items():
            if not records:
                continue
            col_label = sheet_name  # first column header = sheet name
            rows = []
            for r in records:
                rows.append({
                    col_label: r["label"],
                    "Stock Qty in KGS":       f"{r['stock']:,.0f}",
                    "Delivered Qty (Kgs)":    f"{r['delivered']:,.0f}",
                    "Closed Kgs":             f"{r['closed']:,.0f}" if r["closed"] > 0 else "-",
                    "Pending Dispatch Kgs":   f"{r['pending']:,.0f}" if r["pending"] > 0 else "-",
                    "Fill Rate %":  f"{int(round(r['fillRate']  * 100))}%",
                    "Closed %":     f"{int(round(r['closedPct'] * 100))}%",
                    "Pen Dis %":    f"{int(round(r['penDis']    * 100))}%",
                })
            frame = pd.DataFrame(rows)
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]
            for col_cells in ws.columns:
                max_len = max(
                    (len(str(c.value)) for c in col_cells if c.value is not None),
                    default=8,
                )
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 45)

    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="fill_rate_dashboard.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True)
