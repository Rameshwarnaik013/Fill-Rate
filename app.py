from flask import Flask, request, jsonify, send_file, Response
import pandas as pd
import numpy as np
from io import BytesIO

app = Flask(__name__)

# ── Origin lookup ──────────────────────────────────────────────────────────────
ITEM_GROUP_ORIGIN = {
    "Almonds": "Indore", "Roasted Almonds": "Indore", "Apricots": "Indore",
    "Roasted & Flavoured Cashew": "Indore", "Dates": "Indore",
    "Healthy Desserts": "Indore", "Cranberries": "Indore",
    "Trail Mixes": "Indore", "Seed Mix": "Indore", "Berry Mix": "Indore",
    "Savoury Mixes": "Indore", "Figs": "Indore", "Pistachio": "Indore",
    "Raisins": "Indore", "Other Seeds": "Indore", "Chia Seeds": "Indore",
    "Walnut": "Indore", "Roasted & Flavoured Makhana": "Purnia",
    "Makhana Raw": "Purnia",
}
MUNCHIES_REBELA = ["makha shaka imli waves", "makha shaka cheese waves"]


def get_origin(row):
    branch = str(row.get("Branch") or "").strip()
    item_group = str(row.get("NEW MIS ITEM GROUP") or "").strip()
    item_name = str(row.get("Item Name") or "").strip().lower()
    customer_group = str(row.get("Customer Group") or "").strip()
    if branch == "Haryana":
        return "CFA(Gurgaon)"
    if item_group == "Cashew":
        return "Indore" if customer_group == "CPC KPKB" else "Udupi"
    if item_group == "Munchies":
        if any(k in item_name for k in MUNCHIES_REBELA):
            return "Rebela"
        return "UD Foods"
    return ITEM_GROUP_ORIGIN.get(item_group, "Indore")


# ── HTML (single-page app) ─────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fill Rate Dashboard</title>
<script src="https://cdn.tailwindcss.com"></script>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
  .tab-btn { border-bottom: 3px solid transparent; transition: all .15s; }
  .tab-btn.active { border-bottom-color: #2563eb; color: #2563eb; }
  .tab-btn:hover:not(.active) { color: #374151; border-bottom-color: #d1d5db; }
  .drop-zone { border: 2px dashed #cbd5e1; transition: all .2s; }
  .drop-zone.drag-over { border-color: #2563eb; background: #eff6ff; }
  tr.grand-total td { background: #e2e8f0 !important; font-weight: 700; }
  tbody tr:not(.grand-total):hover td { background: #f8fafc; }
  thead th { position: sticky; top: 0; z-index: 1; }
</style>
</head>
<body class="bg-gray-50 min-h-screen">

<!-- Header -->
<header class="bg-white border-b shadow-sm">
  <div class="max-w-screen-2xl mx-auto px-6 py-3 flex items-center gap-3">
    <span class="text-2xl">&#128230;</span>
    <h1 class="text-xl font-semibold text-gray-800">Fill Rate Dashboard</h1>
    <span class="ml-auto text-xs text-gray-400" id="file-badge"></span>
  </div>
</header>

<div class="max-w-screen-2xl mx-auto px-6 py-6 space-y-5">

  <!-- Upload -->
  <div class="drop-zone bg-white rounded-2xl p-10 text-center cursor-pointer select-none" id="drop-zone"
       onclick="document.getElementById('file-input').click()">
    <div class="text-4xl mb-2">&#128193;</div>
    <p class="text-gray-500 text-sm">Click to upload or drag &amp; drop your Fill Rate Excel file</p>
    <p class="text-gray-400 text-xs mt-1">.xlsx / .xls</p>
    <input type="file" id="file-input" accept=".xlsx,.xls" class="hidden">
    <p id="upload-msg" class="mt-3 text-sm text-gray-400"></p>
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
    <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4" id="kpi-cards"></div>

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
      </div>

      <!-- Table -->
      <div class="overflow-x-auto">
        <table class="w-full text-sm border-collapse">
          <thead>
            <tr class="bg-slate-50 border-b border-gray-200">
              <th class="px-4 py-3 text-left text-xs font-semibold text-gray-500 uppercase tracking-wide w-56" id="col-label">Item Group</th>
              <th class="px-4 py-3 text-right text-xs font-semibold text-gray-500 uppercase tracking-wide">Stock Qty (KGS)</th>
              <th class="px-4 py-3 text-right text-xs font-semibold text-gray-500 uppercase tracking-wide">Delivered Qty (KGS)</th>
              <th class="px-4 py-3 text-right text-xs font-semibold text-gray-500 uppercase tracking-wide">Closed KGS</th>
              <th class="px-4 py-3 text-right text-xs font-semibold text-gray-500 uppercase tracking-wide">Pending Dispatch KGS</th>
              <th class="px-4 py-3 text-center text-xs font-semibold text-gray-500 uppercase tracking-wide">Fill Rate %</th>
              <th class="px-4 py-3 text-center text-xs font-semibold text-gray-500 uppercase tracking-wide">Closed %</th>
              <th class="px-4 py-3 text-center text-xs font-semibold text-gray-500 uppercase tracking-wide">Pen Dis %</th>
            </tr>
          </thead>
          <tbody id="table-body"></tbody>
        </table>
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
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#70AD47">&#8805;85%</span>
          <span class="inline-block px-2 py-0.5 rounded font-bold ml-1" style="background:#FFD966">70-84%</span>
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
let currentTab = 'item-group';

const TABS = {
  'item-group':      { key: 'NEW MIS ITEM GROUP', label: 'Item Group' },
  'date':            { key: 'Sales Order Date',   label: 'SO Date'    },
  'customer-group':  { key: 'Customer Group',     label: 'Customer Group' },
  'customer':        { key: 'Customer',           label: 'Customer'   },
  'origin':          { key: 'Origin',             label: 'Origin'     },
};

// ── Upload / Drag & Drop ──────────────────────────────────────────────────────
const dz = document.getElementById('drop-zone');
document.getElementById('file-input').addEventListener('change', e => {
  if (e.target.files[0]) uploadFile(e.target.files[0]);
});
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0]);
});

function setMsg(txt, err = false) {
  const el = document.getElementById('upload-msg');
  el.textContent = txt;
  el.className = 'mt-3 text-sm ' + (err ? 'text-red-500' : 'text-gray-400');
}

async function uploadFile(file) {
  setMsg('Uploading ' + file.name + '…');
  const fd = new FormData();
  fd.append('file', file);
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) { setMsg(data.error || 'Upload failed', true); return; }
    allRows = data.rows;
    document.getElementById('file-badge').textContent = file.name + ' · ' + allRows.length.toLocaleString() + ' rows';
    initFilters();
    applyFilters();
    document.getElementById('dashboard').classList.remove('hidden');
    setMsg('✓ Loaded ' + allRows.length.toLocaleString() + ' rows');
  } catch (err) {
    setMsg('Error: ' + err.message, true);
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
    { label: 'Total Stock (KGS)',     val: fmtN(ts), bg: '#fff',    fg: '#1e40af' },
    { label: 'Total Delivered (KGS)', val: fmtN(td), bg: '#fff',    fg: '#065f46' },
    { label: 'Fill Rate',             val: pct(fr),  bg: fillClr(fr), fg: '#000' },
    { label: 'Closed %',              val: pct(cp),  bg: closedClr(cp), fg: '#000' },
    { label: 'Pen Dis %',             val: pct(pp),  bg: penClr(pp),  fg: '#000' },
  ];
  document.getElementById('kpi-cards').innerHTML = cards.map(c => `
    <div class="rounded-2xl shadow-sm border border-gray-100 p-4" style="background:${c.bg}">
      <p class="text-[11px] font-semibold text-gray-400 uppercase tracking-wider">${c.label}</p>
      <p class="text-2xl font-bold mt-1" style="color:${c.fg}">${c.val}</p>
    </div>`).join('');
}

// ── Tab switch ────────────────────────────────────────────────────────────────
function switchTab(tab, btn) {
  currentTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('col-label').textContent = TABS[tab].label;
  renderTable(tab, filteredRows());
}

// ── Aggregation ───────────────────────────────────────────────────────────────
function aggregate(rows, key) {
  const g = {};
  rows.forEach(r => {
    const k = (r[key] != null && r[key] !== '') ? r[key] : '(blank)';
    if (!g[k]) g[k] = { stock: 0, delivered: 0, closed: 0, pending: 0 };
    g[k].stock     += r['Stock Qty in KGS']        || 0;
    g[k].delivered += r['Delivered Qty (Kgs)']     || 0;
    g[k].closed    += r['Closed Kgs']              || 0;
    g[k].pending   += r['Pending Dispatch Kgs']    || 0;
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
function renderTable(tab, rows) {
  const data = aggregate(rows, TABS[tab].key);
  document.getElementById('table-body').innerHTML = data.map(r => `
    <tr class="${r.isTotal ? 'grand-total' : ''} border-t border-gray-100">
      <td class="px-4 py-2.5 text-gray-800">${r.label}</td>
      <td class="px-4 py-2.5 text-right text-gray-700">${fmtN(r.stock)}</td>
      <td class="px-4 py-2.5 text-right text-gray-700">${fmtN(r.delivered)}</td>
      <td class="px-4 py-2.5 text-right text-gray-700">${r.closed > 0 ? fmtN(r.closed) : '-'}</td>
      <td class="px-4 py-2.5 text-right text-gray-700">${r.pending > 0 ? fmtN(r.pending) : '-'}</td>
      <td class="px-4 py-2.5 text-center font-bold" style="background:${fillClr(r.fillRate)}">${pct(r.fillRate)}</td>
      <td class="px-4 py-2.5 text-center font-bold" style="background:${closedClr(r.closedPct)}">${pct(r.closedPct)}</td>
      <td class="px-4 py-2.5 text-center font-bold" style="background:${penClr(r.penDis)}">${pct(r.penDis)}</td>
    </tr>`).join('');
}

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

function fillClr(v)   { return v >= 0.85 ? '#70AD47' : v >= 0.70 ? '#FFD966' : '#FF6B6B'; }
function closedClr(v) { return v <= 0.05 ? '#70AD47' : v <= 0.10 ? '#FFD966' : '#FF6B6B'; }
function penClr(v)    { return v <= 0.05 ? '#70AD47' : v <= 0.15 ? '#FFD966' : '#FF6B6B'; }
</script>
</body>
</html>"""


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return Response(HTML, mimetype="text/html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    required = ["Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs",
                "Pending Dispatch Kgs", "NEW MIS ITEM GROUP"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return jsonify({"error": f"Missing columns: {', '.join(missing)}"}), 400

    num_cols = ["Stock Qty in KGS", "Delivered Qty (Kgs)", "Closed Kgs", "Pending Dispatch Kgs"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "Sales Order Date" in df.columns:
        df["Sales Order Date"] = (
            pd.to_datetime(df["Sales Order Date"], errors="coerce")
            .dt.strftime("%Y-%m-%d")
        )

    df["Origin"] = df.apply(get_origin, axis=1)

    keep = ["NEW MIS ITEM GROUP", "Sales Order Date", "Customer Group",
            "Customer", "Origin"] + num_cols
    keep = [c for c in keep if c in df.columns]

    rows = df[keep].where(pd.notnull(df[keep]), None).to_dict("records")
    return jsonify({"rows": rows})


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
