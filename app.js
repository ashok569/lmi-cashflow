/* ===========================================================
   LMI Cashflow Manager — application logic
   VERSION 2.4.10 — new LMI South Asia header image; Nature OTHER manual entry on PI+TI; email templates (PI: Nirali standard, TI: Nirali standard) with live template switcher; PROGRAM dropdown with auto-fill from product list; Add Freight button; max 3 program rows — adds: Product master (ADD PRODUCT, PRODUCT LIST, CSV import, OFFLINE/ONLINE/OTHER categories, MOVE button), multi-line invoice items (Add another program), dynamic Word doc (landscape, LMI South Asia header, QTY column, no freight row, multi-item rows), delete receivable PI ripple, date of supply pre-fill, cancel invoice retains delete button. — fix: Word download uses Packer.toBlob (browser-compatible) instead of Packer.toBuffer (Node-only); fix dataset.invWord reference; invBuildWordDoc returns Document not Buffer. — edit PI/TI syncs cashflow receivable amount; cancel vs permanent delete modal; invUpdateReceivableAmount() — header updated to match actual LMI India letterhead (LMI INDIA branding, Apeejay House address, CIN, email/web/tel, logo placeholder), footer updated.
   doc output matching exact template layout (15-col table,
   all fields, borders, amounts in words), next number preview,
   invoicing module auto-opens from dashboard buttons.
   VERSION 2.3.1 — adds: Test Mode (snapshot/restore), PDF
   generation (jsPDF), improved Email modal with Gmail link,
   delete client confirmed wired, all client fields verified.
   VERSION 2.3.0 — adds: Invoicing module (PI/TI generation,
   client master, Excel export, cashflow integration, reports).
   VERSION 2.2.1 — adds: Pending Actions board (shared task
   list with sections per team member, done/delete, realtime),
   opening balance Reset-to-auto button.
   VERSION 2.1.3 — fix: opening balance Reset to auto button.
   VERSION 2.1.2 — fix: provisioning stamp prevents deleted
   payments (Salaries, Reimbursements etc.) from reappearing
   on navigation. sync.js updated to stamp existing cloud months.
   VERSION 2.1.1 — adds: Scenario Planner (what-if cashflow
   modelling, 3-month projection, save/load/delete scenarios),
   receipt delete GST deduction modal, receipt render fix.
   VERSION 2.1 — adds: Ad Hoc Receivables button (no PI/GST),
   carry-forward exclusion from future-month totals, full-
   payment receipt deletion restores receivable.
   VERSION 2 — includes: FD/BG/Receivables edit+delete, payment
   status click-toggle, Schedule-to-month, opening balance fix,
   recurring-forward-only, part-payment <-> receivable
   reconciliation, self-healing month provisioning.
   Single-file, localStorage-backed cashflow tracker.
   =========================================================== */

const STORAGE_KEY = 'lmi_cashflow_v1';
const STANDARD_VENDORS_KEY_DEFAULT = [
  { name: 'AT HDFC payment', amount: 0, tds: false },
  { name: 'SN HDFC payment', amount: 0, tds: false },
  { name: 'ISHA', amount: 0, tds: false },
  { name: 'VISHAL', amount: 0, tds: false },
  { name: 'Calendly', amount: 0, tds: false },
  { name: 'Zoom', amount: 0, tds: false },
  { name: 'Fathom', amount: 0, tds: false },
  { name: 'Stationery', amount: 0, tds: false },
  { name: 'Professional Couriers', amount: 0, tds: false },
  { name: 'Lic Ad budget', amount: 15000, tds: false },
];

const TDS_ESTIMATE = 150000;
const WACO_FEE = 250000;
const QUARTERLY_MONTHS = [4, 7, 10, 1]; // Apr, Jul, Oct, Jan

/* ---------- Financial year helpers ---------- */
// FY label like "26-27" -> starts April of 2026 (20YY).
function fyStartYear(fyLabel) {
  const [a] = fyLabel.split('-');
  return 2000 + parseInt(a, 10);
}
function fyLabelForStartYear(y) {
  const yy = y % 100;
  const yy2 = (y + 1) % 100;
  return `${String(yy).padStart(2,'0')}-${String(yy2).padStart(2,'0')}`;
}
function monthsOfFY(fyLabel) {
  const startYear = fyStartYear(fyLabel);
  const out = [];
  for (let i = 0; i < 12; i++) {
    let m = 4 + i;
    let y = startYear;
    if (m > 12) { m -= 12; y += 1; }
    out.push(monthKey(y, m));
  }
  return out;
}
function monthKey(y, m) { return `${y}-${String(m).padStart(2,'0')}`; }
function monthLabel(mk) {
  const [y, m] = mk.split('-').map(Number);
  const names = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${names[m]} ${String(y).slice(2)}`;
}
function nextMonthKey(mk) {
  let [y, m] = mk.split('-').map(Number);
  m += 1; if (m > 12) { m = 1; y += 1; }
  return monthKey(y, m);
}
function prevMonthKey(mk) {
  let [y, m] = mk.split('-').map(Number);
  m -= 1; if (m < 1) { m = 12; y -= 1; }
  return monthKey(y, m);
}
function monthNum(mk) { return parseInt(mk.split('-')[1], 10); }
function fyLabelForMonth(mk) {
  const [y, m] = mk.split('-').map(Number);
  if (m >= 4) return fyLabelForStartYear(y);
  return fyLabelForStartYear(y - 1);
}
function todayMonthKey() {
  const d = new Date();
  return monthKey(d.getFullYear(), d.getMonth() + 1);
}

/* ---------- State ---------- */
let DB = null;

function emptyMonth() {
  return { opening: 0, hdfc: 0, yesbank: 0, openingManual: false, receipts: [], payments: [], receivables: [] };
}

function loadDB() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { /* fall through to seed */ }
  }
  return buildSeedDB();
}

function buildSeedDB() {
  const seedEl = document.getElementById('seedData');
  const seed = JSON.parse(seedEl.textContent);
  const db = {
    months: {},
    fixedDeposits: seed.fixedDeposits.map(f => ({ id: uid(), type: f.type, amount: f.amount, term: f.term || '' })),
    bankGuarantees: seed.bankGuarantees.map(b => ({ id: uid(), name: b.name, amount: b.amount })),
    standardVendors: STANDARD_VENDORS_KEY_DEFAULT.map(v => ({ id: uid(), ...v })),
    recurringTemplate: [], // [{name, amount, tds}] — recurring payments going forward from current point
    currentFY: '26-27',
    selectedMonth: null,
  };
  for (const [mk, m] of Object.entries(seed.months)) {
    db.months[mk] = {
      opening: m.opening || 0,
      hdfc: m.hdfc || 0,
      yesbank: m.yesbank || 0,
      openingManual: true,
      _provisioned: true, // seeded from spreadsheet — treat as already provisioned
      receipts: m.receipts.map(r => ({ id: uid(), name: r.name, amount: r.amount, status: r.status || 'RECD' })),
      payments: m.payments.map(p => ({ id: uid(), name: p.name, amount: p.amount, status: p.status || 'planned', recurring: !!p.recurring, tds: !!p.tds })),
      receivables: m.receivables.map(r => ({ id: uid(), name: r.name, amount: r.amount })),
    };
  }
  // Seed recurring template from whichever seeded month has the most starred (*) items
  let bestMk = null, bestCount = -1;
  for (const [mk, m] of Object.entries(db.months)) {
    const c = m.payments.filter(p => p.recurring).length;
    if (c > bestCount) { bestCount = c; bestMk = mk; }
  }
  const templateMonth = bestMk ? db.months[bestMk] : null;
  db.recurringTemplate = templateMonth
    ? templateMonth.payments
        .filter(p => p.recurring)
        .map(p => ({ name: p.name.replace(/\s*(for|incl)\s+\w*\d*.*$/i, '').trim() || p.name, amount: p.amount, tds: false }))
    : [];

  const lastMk = Object.keys(seed.months).sort().pop();
  db.selectedMonth = lastMk;
  return db;
}

function uid() { return 'x' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }

function saveDB(extraMonthKeys) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
  if (window.Cloud && Cloud.cloudConfigured() && Cloud.currentUser) {
    Cloud.setSyncStatus('saving');
    Cloud.queueCloudSave(null); // workspace-level lists (FDs, BGs, vendors, recurring template, FY)
    const keys = new Set([DB.selectedMonth, ...(extraMonthKeys || [])].filter(Boolean));
    keys.forEach(mk => Cloud.queueCloudSave(mk));
  }
}

function getMonth(mk) {
  if (!DB.months[mk]) DB.months[mk] = emptyMonth();
  return DB.months[mk];
}

function ensureMonthExists(mk) {
  if (DB.months[mk]) return DB.months[mk];
  const m = emptyMonth();
  // opening dynamically follows previous month's closing unless manually set
  m.opening = computeClosing(prevMonthKey(mk));
  m.openingManual = false;
  DB.months[mk] = m;
  return m;
}

/* ---------- Computation ---------- */
function monthTotals(mk) {
  const m = DB.months[mk] || emptyMonth();
  const receiptsTotal = m.receipts.reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const paymentsTotal = m.payments.reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const paymentsOnHold = m.payments.filter(p => p.status === 'on hold').reduce((s, p) => s + (Number(p.amount) || 0), 0);
  // Carried receivables only count toward the total once the month becomes the current month.
  // While it's still a future month they are shown for visibility but excluded from the total.
  const isFutureMonth = mk > todayMonthKey();
  const receivablesTotal = m.receivables
    .filter(r => !isFutureMonth || !r._carriedFrom)
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);
  return { receiptsTotal, paymentsTotal, paymentsOnHold, receivablesTotal };
}

function getOpening(mk) {
  const m = DB.months[mk];
  if (m && m.openingManual) return Number(m.opening) || 0;
  // dynamic: previous month's closing
  if (DB.months[prevMonthKey(mk)]) {
    return computeClosing(prevMonthKey(mk));
  }
  return m ? (Number(m.opening) || 0) : 0;
}

function computeClosing(mk) {
  if (!DB.months[mk]) return 0;
  const { receiptsTotal, paymentsTotal } = monthTotals(mk);
  return getOpening(mk) + receiptsTotal - paymentsTotal;
}

function anticipatedStatus(mk) {
  const { paymentsOnHold } = monthTotals(mk);
  const closing = computeClosing(mk);
  return closing - paymentsOnHold;
}

/* ===========================================================
   RECEIVABLES CARRY-FORWARD
   Source: always today's calendar month.
   Target: always the following calendar month.
   Receivables with balance > 0 in today's month are mirrored
   in next month as _carriedFrom entries. Next month's own
   independently-added receivables are never touched.
   Call syncCarriedReceivables() after any change to today's
   month's receivables list.
   =========================================================== */
function syncCarriedReceivables() {
  const todayMk = todayMonthKey();
  const nextMk = nextMonthKey(todayMk);
  const todayMonth = DB.months[todayMk];
  if (!todayMonth) return; // today's month doesn't exist yet, nothing to carry

  const nextMonth = ensureMonthExists(nextMk);

  // Remove all existing carried entries from next month (they'll be rebuilt from scratch)
  nextMonth.receivables = nextMonth.receivables.filter(r => !r._carriedFrom);

  // Re-insert carried entries for every today-month receivable with balance > 0
  todayMonth.receivables
    .filter(r => !r._carriedFrom && (Number(r.amount) || 0) > 0)
    .forEach(r => {
      nextMonth.receivables.push({
        id: 'carried_' + r.id,   // stable ID derived from source so we can find it
        name: r.name,
        amount: r.amount,
        _carriedFrom: todayMk,   // marks this as a mirror, not an original
        _sourceId: r.id,
      });
    });
}

// When a month rolls over (e.g. today is now July but July still has _carriedFrom
// entries from June), promote those carried entries to real editable receivables
// so they can be actioned (payment received, edit, delete) in the current month.
function promoteCarriedReceivables() {
  const todayMk = todayMonthKey();
  const todayMonth = DB.months[todayMk];
  if (!todayMonth) return;

  let changed = false;
  todayMonth.receivables = todayMonth.receivables.map(r => {
    if (!r._carriedFrom) return r; // already a real entry
    // This entry was carried from a prior month — now that we're IN this month,
    // promote it to a proper editable receivable by removing the carried flags.
    changed = true;
    const { _carriedFrom, _sourceId, ...rest } = r;
    // Give it a real ID if it still has the 'carried_' prefix
    return { ...rest, id: rest.id.startsWith('carried_') ? uid() : rest.id };
  });

  if (changed) {
    saveDB([todayMk]);
  }
}

/* ===========================================================
   RENDERING
   =========================================================== */

function fmtMoney(n) {
  n = Number(n) || 0;
  const neg = n < 0;
  n = Math.abs(n);
  const s = n.toLocaleString('en-IN', { maximumFractionDigits: 0 });
  return (neg ? '-' : '') + '\u20B9' + s;
}

function renderAll() {
  renderFYSelect();
  renderMonthTabs();
  renderDashboard();
  renderRecurringList();
  renderFDList();
  renderBGList();
  renderLedgers();
  renderReceivables();
  renderImports();
}

function renderFYSelect() {
  const sel = document.getElementById('fySelect');
  const startYears = new Set();
  Object.keys(DB.months).forEach(mk => startYears.add(fyStartYear(fyLabelForMonth(mk))));
  startYears.add(fyStartYear(DB.currentFY));
  const base = Math.min(...startYears);
  const opts = [];
  for (let i = 0; i < 11; i++) {
    const y = base + i;
    opts.push(fyLabelForStartYear(y));
  }
  sel.innerHTML = opts.map(fy => `<option value="${fy}" ${fy === DB.currentFY ? 'selected' : ''}>FY ${fy}</option>`).join('');
  sel.onchange = () => {
    const prevFY = DB.currentFY;
    DB.currentFY = sel.value;
    const months = monthsOfFY(DB.currentFY);
    let touched = [];
    if (!months.includes(DB.selectedMonth)) {
      touched = selectMonth(months[0]);
    }
    // If switching to a FY that has no data yet, seed all 12 months with recurring template
    const isNewFY = months.every(mk => !DB.months[mk] || !DB.months[mk]._provisioned);
    if (isNewFY && DB.recurringTemplate && DB.recurringTemplate.length > 0) {
      const firstMonth = months[0];
      const seeded = applyRecurringForward(firstMonth);
      // Also run ensureMonthlyProvisions for TDS/WACO on each month
      seeded.forEach(mk => ensureMonthlyProvisions(mk));
      touched = [...new Set([...touched, ...seeded])];
      toast(`Recurring payments seeded into FY ${DB.currentFY}`);
    }
    saveDB(touched);
    renderAll();
  };
}

// Switches the active month, creating it and applying standing provisions (TDS estimate,
// quarterly Waco fee, recurring payments) if it's new OR if it already exists but is missing
// them (e.g. months created before this auto-provisioning existed). Returns touched month keys.
function monthIsMissingProvisions(mk) {
  const m = DB.months[mk];
  if (!m) return false;
  if (m._provisioned) return false;
  if (mk <= todayMonthKey()) return false;
  // If recurringTemplate has entries and none are in this month's payments, needs provisioning
  const hasAnyRecurringTemplate = DB.recurringTemplate && DB.recurringTemplate.length > 0;
  const hasRecurringPayments = m.payments.some(p => p.recurring);
  return hasAnyRecurringTemplate && !hasRecurringPayments;
}

function selectMonth(mk) {
  const isNew = !DB.months[mk];
  const needsHealing = !isNew && monthIsMissingProvisions(mk);
  DB.selectedMonth = mk;
  ensureMonthExists(mk);
  const touched = [mk];
  if (isNew || needsHealing) {
    ensureMonthlyProvisions(mk);
  }
  if (isNew) {
    const prevMk = prevMonthKey(mk);
    recalcTDSRollup(prevMk);
    if (DB.months[prevMk]) touched.push(prevMk);
  }
  return touched;
}

function renderMonthTabs() {
  const wrap = document.getElementById('monthTabs');
  const months = monthsOfFY(DB.currentFY);
  const tmk = todayMonthKey();
  const isPendingOpen = document.getElementById('pendingActionsOverlay') &&
    document.getElementById('pendingActionsOverlay').style.display !== 'none';
  const isInvoicingOpen = document.getElementById('invoicingOverlay') &&
    document.getElementById('invoicingOverlay').style.display !== 'none';
  wrap.innerHTML = months.map(mk => {
    const active = (mk === DB.selectedMonth && !isPendingOpen && !isInvoicingOpen) ? 'active' : '';
    const isCurrent = mk === tmk ? 'is-current' : '';
    const hasData = !!DB.months[mk];
    return `<button class="month-tab ${active} ${isCurrent} ${hasData ? '' : 'future'}" data-mk="${mk}">${monthLabel(mk)}${isCurrent ? '<span class="dot"></span>' : ''}</button>`;
  }).join('') +
  `<button class="month-tab ${isPendingOpen ? 'active' : ''}" id="pendingActionsTab" style="border-left:2px solid rgba(255,255,255,.15); margin-left:8px;">&#9654; Pending actions</button>` +
  `<button class="month-tab ${isInvoicingOpen ? 'active' : ''}" id="invoicingTab" style="border-left:1px solid rgba(255,255,255,.1); margin-left:4px;">&#128196; Invoicing</button>`;

  wrap.querySelectorAll('.month-tab[data-mk]').forEach(btn => {
    btn.onclick = () => {
      closePendingActions();
      closeInvoicingModule();
      const touched = selectMonth(btn.dataset.mk);
      saveDB(touched);
      renderAll();
    };
  });
  const paTab = document.getElementById('pendingActionsTab');
  if (paTab) paTab.onclick = openPendingActions;
  const invTab = document.getElementById('invoicingTab');
  if (invTab) invTab.onclick = openInvoicingModule;
}

function renderDashboard() {
  const mk = DB.selectedMonth;
  ensureMonthExists(mk);
  const opening = getOpening(mk);
  const closing = computeClosing(mk);
  const anticipated = anticipatedStatus(mk);
  const { receiptsTotal, paymentsTotal, paymentsOnHold, receivablesTotal } = monthTotals(mk);

  document.getElementById('obMonthLabel').textContent = monthLabel(mk);
  document.getElementById('openingAmt').textContent = fmtMoney(opening);
  document.getElementById('openingAmt').classList.toggle('neg', opening < 0);
  document.getElementById('closingAmt').textContent = fmtMoney(closing);
  document.getElementById('closingAmt').classList.toggle('neg', closing < 0);
  document.getElementById('anticipatedAmt').textContent = fmtMoney(anticipated);

  document.getElementById('statReceipts').textContent = fmtMoney(receiptsTotal);
  document.getElementById('statPayments').textContent = fmtMoney(paymentsTotal);
  document.getElementById('statReceivables').textContent = fmtMoney(receivablesTotal);
  document.getElementById('statHold').textContent = fmtMoney(paymentsOnHold);
}

function renderRecurringList() {
  const wrap = document.getElementById('recurList');
  const list = DB.recurringTemplate || [];
  if (!list.length) {
    wrap.innerHTML = '<div class="empty-note">No recurring payments set up yet.</div>';
    return;
  }

  const monthly   = list.filter(r => !r.frequency || r.frequency === 'monthly');
  const quarterly = list.filter(r => r.frequency === 'quarterly');
  const annual    = list.filter(r => r.frequency === 'annual');

  function recRow(r) {
    return `<div class="recur-item">
      <span class="nm"><span class="star">&#9733;</span>${escapeHtml(r.name)}</span>
      <span class="amt">${fmtMoney(r.amount)}</span>
    </div>`;
  }

  function recSep(label, color) {
    return `<div style="font-size:10px;font-weight:700;color:${color};text-transform:uppercase;
      letter-spacing:.07em;padding:5px 0 2px;border-bottom:1px solid ${color}44;
      margin-top:6px;">${label}</div>`;
  }

  let html = '';
  if (monthly.length)   html += recSep('Monthly', '#1e4f8a')   + monthly.map(recRow).join('');
  if (quarterly.length) html += recSep('Quarterly', '#1f7a4d') + quarterly.map(recRow).join('');
  if (annual.length)    html += recSep('Annual', '#9a6b14')    + annual.map(recRow).join('');
  wrap.innerHTML = html;
}

function renderFDList() {
  const wrap = document.getElementById('fdList');
  const list = DB.fixedDeposits || [];
  if (!list.length) { wrap.innerHTML = '<div class="empty-note">No fixed deposits recorded.</div>'; return; }
  const total = list.reduce((s, f) => s + (Number(f.amount) || 0), 0);
  wrap.innerHTML = list.map(f => `
    <div class="recur-item">
      <span class="nm">${escapeHtml(f.type)}${f.term ? ' &middot; ' + escapeHtml(f.term) : ''}</span>
      <span class="amt">${fmtMoney(f.amount)}</span>
      <button data-edit-fd="${f.id}" title="Edit">&#9998;</button>
    </div>`).join('') + `<div class="recur-item" style="border-top:1px solid var(--line); margin-top:4px; padding-top:9px;"><span class="nm" style="font-weight:700;">Total</span><span class="amt" style="font-weight:700;">${fmtMoney(total)}</span></div>`;
  wrap.querySelectorAll('[data-edit-fd]').forEach(b => b.onclick = () => openEditFd(b.dataset.editFd));
}

function renderBGList() {
  const wrap = document.getElementById('bgList');
  const list = DB.bankGuarantees || [];
  if (!list.length) { wrap.innerHTML = '<div class="empty-note">No bank guarantees recorded.</div>'; return; }
  wrap.innerHTML = list.map(b => `
    <div class="recur-item">
      <span class="nm">${escapeHtml(b.name)}</span>
      <span class="amt">${fmtMoney(b.amount)}</span>
      <button data-edit-bg="${b.id}" title="Edit">&#9998;</button>
    </div>`).join('');
  wrap.querySelectorAll('[data-edit-bg]').forEach(b => b.onclick = () => openEditBg(b.dataset.editBg));
}

function statusBadge(status) {
  const cls = (status || '').toLowerCase().replace(/\s+/g, '-') === 'on-hold' ? 'hold' : (status || '').toLowerCase();
  const label = status || '';
  return `<span class="badge ${cls}">${escapeHtml(label)}</span>`;
}

function renderLedgers() {
  const mk = DB.selectedMonth;
  const m = getMonth(mk);

  const rTotal = m.receipts.reduce((s, r) => s + (Number(r.amount) || 0), 0);
  document.getElementById('receiptsTotal').textContent = fmtMoney(rTotal);
  const receiptsRows = document.getElementById('receiptsRows');
  if (!m.receipts.length) {
    receiptsRows.innerHTML = '<div class="empty-note" style="padding:14px;">No receipts recorded for this month.</div>';
  } else {
    receiptsRows.innerHTML = m.receipts.map(r => `
      <div class="lrow">
        <span class="nm">${escapeHtml(r.name)}</span>
        <span class="amt">${fmtMoney(r.amount)}</span>
        ${statusBadge(r.status)}
        <span class="row-actions">
          <button data-edit-receipt="${r.id}" title="Edit">&#9998;</button>
          <button data-del-receipt="${r.id}" title="Delete">&#10005;</button>
        </span>
      </div>`).join('');
  }

  const pTotal = m.payments.reduce((s, p) => s + (Number(p.amount) || 0), 0);
  document.getElementById('paymentsTotal').textContent = fmtMoney(pTotal);
  const paymentsRows = document.getElementById('paymentsRows');
  if (!m.payments.length) {
    paymentsRows.innerHTML = '<div class="empty-note" style="padding:14px;">No payments recorded for this month.</div>';
  } else {
    paymentsRows.innerHTML = m.payments.map(p => `
      <div class="lrow">
        <span class="nm">${p.recurring ? '<span class="star">&#9733;</span>' : ''}${escapeHtml(p.name)}${p.tds ? ' <span style="color:var(--ink-soft); font-size:10.5px;">(TDS)</span>' : ''}</span>
        <span class="amt">${p.amount ? fmtMoney(p.amount) : '<span style="color:var(--ink-soft);">pending</span>'}</span>
        <span data-toggle-status="${p.id}" style="cursor:pointer;" title="Click to change status">${statusBadge(p.status)}</span>
        <span class="row-actions">
          <button data-edit-payment="${p.id}" title="Edit">&#9998;</button>
          <button data-del-payment="${p.id}" title="Delete">&#10005;</button>
        </span>
      </div>`).join('');
  }

  // wire row actions
  receiptsRows.querySelectorAll('[data-edit-receipt]').forEach(b => b.onclick = () => openEditReceipt(b.dataset.editReceipt));
  receiptsRows.querySelectorAll('[data-del-receipt]').forEach(b => b.onclick = () => deleteReceipt(b.dataset.delReceipt));
  paymentsRows.querySelectorAll('[data-edit-payment]').forEach(b => b.onclick = () => openEditPayment(b.dataset.editPayment));
  paymentsRows.querySelectorAll('[data-del-payment]').forEach(b => b.onclick = () => deletePayment(b.dataset.delPayment));
  paymentsRows.querySelectorAll('[data-toggle-status]').forEach(el => el.onclick = () => togglePaymentStatus(el.dataset.toggleStatus));
}

// Click-to-cycle: planned -> paid -> on hold -> planned (item 4)
function togglePaymentStatus(id) {
  const m = getMonth(DB.selectedMonth);
  const p = m.payments.find(p => p.id === id);
  if (!p) return;
  const order = ['planned', 'paid', 'on hold'];
  const idx = order.indexOf((p.status || 'planned').toLowerCase());
  p.status = order[(idx + 1) % order.length];
  saveDB(); renderAll();
}

function renderReceivables() {
  const mk = DB.selectedMonth;
  const m = getMonth(mk);
  const tbody = document.querySelector('#receivablesTable tbody');
  const isFutureMonth = mk > todayMonthKey();
  // Carried items excluded from total until the month becomes current
  const total = m.receivables
    .filter(r => !isFutureMonth || !r._carriedFrom)
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);
  document.getElementById('receivablesTotal').textContent = fmtMoney(total);
  if (!m.receivables.length) {
    tbody.innerHTML = '<tr><td colspan="3" class="empty-note">No outstanding receivables.</td></tr>';
    return;
  }
  tbody.innerHTML = m.receivables.map(r => {
    const isCarried = !!r._carriedFrom;
    const nameCell = isCarried
      ? `${escapeHtml(r.name)} <span style="font-size:10px; color:var(--ink-soft); font-style:italic;">↩ b/f from ${monthLabel(r._carriedFrom)}</span>`
      : escapeHtml(r.name);
    const actions = isCarried
      ? `<span style="color:var(--ink-soft); font-size:11px;" title="Edit this in ${monthLabel(r._carriedFrom)}">&#128274;</span>`
      : `<button data-receive="${r.id}" title="Record payment received">&#10003;</button>
         <button data-edit-receivable="${r.id}" title="Edit">&#9998;</button>
         <button data-del-receivable="${r.id}" title="Delete">&#10005;</button>`;
    return `<tr>
      <td>${nameCell}</td>
      <td class="amt">${fmtMoney(r.amount)}</td>
      <td class="row-actions">${actions}</td>
    </tr>`;
  }).join('');
  tbody.querySelectorAll('[data-receive]').forEach(b => b.onclick = () => openPaymentReceivedFor(b.dataset.receive));
  tbody.querySelectorAll('[data-edit-receivable]').forEach(b => b.onclick = () => openEditReceivable(b.dataset.editReceivable));
  tbody.querySelectorAll('[data-del-receivable]').forEach(b => b.onclick = () => deleteReceivable(b.dataset.delReceivable));
}

function renderImports() {
  const mk = DB.selectedMonth;
  const m = getMonth(mk);
  const tbody = document.querySelector('#importsTable tbody');
  const orders = (m.imports || []);
  const total = orders.reduce((s, o) => s + (Number(o.amount) || 0) + (Number(o.kerry) || 0) + (Number(o.ceva) || 0), 0);
  document.getElementById('importsTotal').textContent = fmtMoney(total);
  if (!orders.length) {
    tbody.innerHTML = '<tr><td colspan="3" class="empty-note">No import orders this month.</td></tr>';
    return;
  }
  tbody.innerHTML = orders.map((o, i) => `
    <tr>
      <td>Order ${i + 1} &middot; Qty ${escapeHtml(String(o.qty))}</td>
      <td class="amt">${fmtMoney(o.amount)}</td>
      <td class="row-actions">
        <button data-edit-import="${o.id}" title="Edit">&#9998;</button>
        <button data-del-import="${o.id}" title="Delete">&#10005;</button>
      </td>
    </tr>`).join('');
  tbody.querySelectorAll('[data-edit-import]').forEach(b => b.onclick = () => openEditImport(b.dataset.editImport));
  tbody.querySelectorAll('[data-del-import]').forEach(b => b.onclick = () => deleteImport(b.dataset.delImport));
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast._tm);
  toast._tm = setTimeout(() => t.classList.remove('show'), 2400);
}

/* ===========================================================
   MODAL SYSTEM
   =========================================================== */

function openModal(title, bodyHtml, footHtml) {
  const root = document.getElementById('modalRoot');
  root.innerHTML = `
    <div class="modal-backdrop open" id="activeModal">
      <div class="modal">
        <div class="modal-head">
          <h3>${title}</h3>
          <button id="modalCloseBtn">&times;</button>
        </div>
        <div class="modal-body">${bodyHtml}</div>
        <div class="modal-foot">${footHtml}</div>
      </div>
    </div>`;
  document.getElementById('modalCloseBtn').onclick = closeModal;
  document.getElementById('activeModal').addEventListener('mousedown', e => {
    if (e.target.id === 'activeModal') closeModal();
  });
}
function closeModal() {
  document.getElementById('modalRoot').innerHTML = '';
}

/* ===========================================================
   ACTION: Update Opening Balance
   =========================================================== */
function openingHintText(mk) {
  const isManual = DB.months[mk] && DB.months[mk].openingManual;
  const basis = isManual ? 'a manually set amount' : "auto-calculated from the previous month's closing balance";
  return `The current figure shown is ${basis}. Enter a new amount and save to replace it — this becomes the new opening balance for that month, and later months keep auto-calculating forward from whatever you set here.`;
}

function openUpdateOpening() {
  const months = monthsOfFY(DB.currentFY);
  const defaultMk = months.includes(DB.selectedMonth) ? DB.selectedMonth : months[0];
  const body = `
    <div class="field">
      <label>Month</label>
      <select id="ob-month">${months.map(mk => `<option value="${mk}" ${mk===defaultMk?'selected':''}>${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="field">
      <label>Opening balance amount</label>
      <input type="number" id="ob-amount" placeholder="e.g. 2500000" value="${getOpening(defaultMk)||''}">
    </div>
    <div class="hint" id="ob-hint">${escapeHtml(openingHintText(defaultMk))}</div>`;
  const foot = `<button class="btn" id="ob-cancel">Cancel</button><button class="btn btn-sm" id="ob-reset" title="Remove manual override — opening balance will auto-calculate from previous month closing">↺ Reset to auto</button><button class="btn btn-primary" id="ob-save">Save</button>`;
  openModal('Update opening balance', body, foot);
  document.getElementById('ob-cancel').onclick = closeModal;

  document.getElementById('ob-month').addEventListener('change', e => {
    const mk = e.target.value;
    document.getElementById('ob-amount').value = getOpening(mk) || '';
    document.getElementById('ob-hint').textContent = openingHintText(mk);
  });

  document.getElementById('ob-reset').onclick = () => {
    const mk = document.getElementById('ob-month').value;
    ensureMonthExists(mk);
    DB.months[mk].openingManual = false;
    DB.months[mk].opening = 0;
    saveDB([mk]); renderAll(); closeModal();
    toast(`${monthLabel(mk)} opening balance now auto-calculates from ${monthLabel(prevMonthKey(mk))} closing`);
  };

  document.getElementById('ob-save').onclick = () => {
    const mk = document.getElementById('ob-month').value;
    const amt = parseFloat(document.getElementById('ob-amount').value) || 0;
    ensureMonthExists(mk);
    DB.months[mk].opening = amt;
    DB.months[mk].openingManual = true;
    saveDB([mk]); renderAll(); closeModal();
    toast(`Opening balance for ${monthLabel(mk)} replaced with ${fmtMoney(amt)}`);
  };
}

/* ===========================================================
   ACTION: Add Pr. Invoice  (-> Receivables)
   =========================================================== */
function openAddInvoice() {
  const body = `
    <div class="field"><label>Licensee</label><input type="text" id="inv-lic" placeholder="e.g. Anand"></div>
    <div class="field-row">
      <div class="field"><label>Units</label><input type="text" id="inv-units" placeholder="e.g. 5"></div>
      <div class="field"><label>PI number</label><input type="text" id="inv-pi" placeholder="e.g. PI 17"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Base amount</label><input type="number" id="inv-base" placeholder="0"></div>
      <div class="field"><label>GST</label><input type="number" id="inv-gst" placeholder="0"></div>
    </div>
    <div class="calc-line"><span>Base amount</span><span id="inv-c-base">&#8377;0</span></div>
    <div class="calc-line"><span>Less 10%</span><span id="inv-c-less">&#8377;0</span></div>
    <div class="calc-line"><span>Plus GST</span><span id="inv-c-gst">&#8377;0</span></div>
    <div class="calc-line total"><span>Total receivable</span><span id="inv-c-total">&#8377;0</span></div>`;
  const foot = `<button class="btn" id="inv-cancel">Cancel</button><button class="btn btn-primary" id="inv-save">Save</button>`;
  openModal('Add Pr. Invoice', body, foot);

  function recalc() {
    const base = parseFloat(document.getElementById('inv-base').value) || 0;
    const gst = parseFloat(document.getElementById('inv-gst').value) || 0;
    const less = base * 0.10;
    const total = (base - less) + gst;
    document.getElementById('inv-c-base').textContent = fmtMoney(base);
    document.getElementById('inv-c-less').textContent = '-' + fmtMoney(less);
    document.getElementById('inv-c-gst').textContent = '+' + fmtMoney(gst);
    document.getElementById('inv-c-total').textContent = fmtMoney(total);
  }
  ['inv-base', 'inv-gst'].forEach(id => document.getElementById(id).addEventListener('input', recalc));
  recalc();

  document.getElementById('inv-cancel').onclick = closeModal;
  document.getElementById('inv-save').onclick = () => {
    const lic = document.getElementById('inv-lic').value.trim();
    const units = document.getElementById('inv-units').value.trim();
    const pi = document.getElementById('inv-pi').value.trim();
    const base = parseFloat(document.getElementById('inv-base').value) || 0;
    const gst = parseFloat(document.getElementById('inv-gst').value) || 0;
    if (!lic) { toast('Licensee name is required'); return; }
    const total = (base - base * 0.10) + gst;
    const text = [lic, units, pi].filter(Boolean).join(' ');
    const m = getMonth(DB.selectedMonth);
    m.receivables.push({ id: uid(), name: text, amount: Math.round(total * 100) / 100, _base: base, _gst: gst });
    if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
    saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
    toast(`Added ${text} to receivables (${fmtMoney(total)})`);
  };
}

/* ===========================================================
   ACTION: Payment Received  (full / part) -> Receipts
   =========================================================== */
function openPaymentReceived() {
  const m = getMonth(DB.selectedMonth);
  const payable = m.receivables.filter(r => !r._carriedFrom);
  if (!payable.length) {
    openModal('Payment received', '<div class="empty-note">No receivables on file for this month. Add a Pr. Invoice first, or record the payment in the month where the receivable was originally entered.</div>', `<button class="btn btn-primary" id="pr-ok">Close</button>`);
    document.getElementById('pr-ok').onclick = closeModal;
    return;
  }
  const body = `
    <div class="field"><label>Select receivable</label></div>
    <div class="sub-list" id="pr-list">
      ${payable.map(r => `<div class="sub-list-item" data-id="${r.id}"><span>${escapeHtml(r.name)}</span><span class="amt">${fmtMoney(r.amount)}</span></div>`).join('')}
    </div>`;
  openModal('Payment received', body, `<button class="btn" id="pr-cancel">Cancel</button>`);
  document.getElementById('pr-cancel').onclick = closeModal;
  document.querySelectorAll('#pr-list .sub-list-item').forEach(el => {
    el.onclick = () => openPaymentReceivedFor(el.dataset.id);
  });
}

function openPaymentReceivedFor(receivableId) {
  const m = getMonth(DB.selectedMonth);
  const rec = m.receivables.find(r => r.id === receivableId);
  if (!rec) { closeModal(); return; }
  const body = `
    <div class="field"><label>Receivable</label><div style="padding:9px 0; font-weight:600;">${escapeHtml(rec.name)} &mdash; ${fmtMoney(rec.amount)}</div></div>
    <div class="radio-pills">
      <div class="radio-pill"><input type="radio" name="pr-type" id="pr-full" value="full" checked><label for="pr-full">Full</label></div>
      <div class="radio-pill"><input type="radio" name="pr-type" id="pr-part" value="part"><label for="pr-part">Part</label></div>
    </div>
    <div class="field" id="pr-amount-field" style="display:none;">
      <label>Amount received</label>
      <input type="number" id="pr-amount" placeholder="0">
    </div>
    <div class="checkrow">
      <input type="checkbox" id="pr-gst-next">
      <label for="pr-gst-next">This includes GST that should auto-increment next month's GST payment</label>
    </div>
    <div class="field" id="pr-gst-field" style="display:none;">
      <label>GST amount</label>
      <input type="number" id="pr-gst-amount" placeholder="0">
    </div>`;
  openModal('Payment received', body, `<button class="btn" id="pr2-cancel">Cancel</button><button class="btn btn-primary" id="pr2-save">Save</button>`);
  document.getElementById('pr2-cancel').onclick = closeModal;

  function toggleType() {
    const isPart = document.getElementById('pr-part').checked;
    document.getElementById('pr-amount-field').style.display = isPart ? 'block' : 'none';
  }
  document.querySelectorAll('input[name="pr-type"]').forEach(r => r.addEventListener('change', toggleType));
  document.getElementById('pr-gst-next').addEventListener('change', e => {
    document.getElementById('pr-gst-field').style.display = e.target.checked ? 'block' : 'none';
  });

  document.getElementById('pr2-save').onclick = () => {
    const isPart = document.getElementById('pr-part').checked;
    const gstNext = document.getElementById('pr-gst-next').checked;
    const gstAmt = parseFloat(document.getElementById('pr-gst-amount').value) || 0;

    if (!isPart) {
      // Store _receivableId so deleting this receipt can restore the receivable (item 1)
      m.receipts.push({ id: uid(), name: rec.name, amount: rec.amount, status: 'RECD', _receivableId: rec.id, _receivableMonth: DB.selectedMonth, _fullPayment: true });
      m.receivables = m.receivables.filter(r => r.id !== rec.id);
      toast(`Recorded full payment of ${fmtMoney(rec.amount)} from ${rec.name}`);
    } else {
      const amt = parseFloat(document.getElementById('pr-amount').value) || 0;
      if (amt <= 0) { toast('Enter the amount received'); return; }
      m.receipts.push({ id: uid(), name: rec.name, amount: amt, status: 'RECD', _receivableId: rec.id, _receivableMonth: DB.selectedMonth });
      rec.amount = Math.max(0, rec.amount - amt);
      if (rec.amount === 0) m.receivables = m.receivables.filter(r => r.id !== rec.id);
      toast(`Recorded part payment of ${fmtMoney(amt)} from ${rec.name}`);
    }

    let extraMk = null;
    if (gstNext && gstAmt > 0) {
      extraMk = nextMonthKey(DB.selectedMonth);
      const nm = ensureMonthExists(extraMk);
      const existingGst = nm.payments.find(p => /^GST for/i.test(p.name));
      if (existingGst) {
        existingGst.amount = (Number(existingGst.amount) || 0) + gstAmt;
      } else {
        nm.payments.push({ id: uid(), name: `GST for ${monthLabel(DB.selectedMonth)}`, amount: gstAmt, status: 'planned', recurring: false, tds: false });
      }
    }

    if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
    const nextMk = nextMonthKey(todayMonthKey());
    saveDB(extraMk ? [extraMk, nextMk] : [nextMk]);
    closeModal();
    setTimeout(() => renderAll(), 50);
  };
  toggleType();
}


/* ===========================================================
   ACTION: Edit Recurring Payments
   =========================================================== */
function openEditRecurring() {
  renderRecurringEditor();
}

const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const FREQ_LABELS = { monthly:'Monthly', quarterly:'Quarterly', annual:'Annual' };

function renderRecurringEditor() {
  const list = DB.recurringTemplate;

  function freqBadge(freq) {
    const colors = { monthly:'#1e4f8a', quarterly:'#1f7a4d', annual:'#9a6b14' };
    const c = colors[freq||'monthly'] || colors.monthly;
    return `<span style="font-size:10px;font-weight:700;color:#fff;background:${c};border-radius:3px;padding:1px 5px;margin-right:6px;">${(FREQ_LABELS[freq]||'Monthly').slice(0,1)}</span>`;
  }

  const monthly  = list.filter(r => !r.frequency || r.frequency === 'monthly');
  const quarterly = list.filter(r => r.frequency === 'quarterly');
  const annual   = list.filter(r => r.frequency === 'annual');

  function itemRow(r, i) {
    return `
      <div class="sub-list-item" style="cursor:default; display:flex; align-items:center; gap:6px;">
        ${freqBadge(r.frequency)}
        <span style="flex:1; font-size:13px;">${escapeHtml(r.name)}${r.startMonth ? ` <span style="font-size:11px;color:var(--ink-soft);">(from ${r.startMonth})</span>` : ''}</span>
        <input type="number" data-rec-idx="${i}" value="${r.amount}"
          style="width:100px;padding:5px 8px;border:1px solid var(--line);border-radius:4px;font-family:var(--mono);">
        <button data-rec-del="${i}" style="border:none;background:none;color:#aab2bd;cursor:pointer;">&#10005;</button>
      </div>`;
  }

  function section(label, color, items, startIdx) {
    if (!items.length) return '';
    return `
      <div style="font-size:11px;font-weight:700;color:${color};text-transform:uppercase;
                  letter-spacing:.07em;padding:8px 0 4px;border-bottom:2px solid ${color}22;
                  margin-top:10px;">${label}</div>
      ${items.map((r, i) => itemRow(r, startIdx + i)).join('')}`;
  }

  const body = `
    <div class="sub-list" id="rec-edit-list" style="max-height:340px;overflow-y:auto;">
      ${!list.length ? '<div class="empty-note" style="padding:12px;">No recurring items yet.</div>' : ''}
      ${section('Monthly', '#1e4f8a', monthly, 0)}
      ${section('Quarterly', '#1f7a4d', quarterly, monthly.length)}
      ${section('Annual', '#9a6b14', annual, monthly.length + quarterly.length)}
    </div>
    <div class="hint" style="margin:10px 0 6px;">Saving applies these to ${monthLabel(DB.selectedMonth)} and forward this FY. Quarterly and Annual items only appear in their applicable months.</div>
    <button class="btn btn-sm" id="rec-add-new">+ Add recurring item</button>`;

  openModal('Edit recurring payments', body,
    `<button class="btn" id="rec-cancel">Cancel</button>
     <button class="btn btn-primary" id="rec-save">Save changes</button>`);

  document.getElementById('rec-cancel').onclick = closeModal;

  document.getElementById('rec-add-new').onclick = () => {
    // Show inline add form
    const addForm = `
      <div id="rec-add-form" style="border:1px solid var(--line);border-radius:6px;padding:12px;margin-top:10px;background:var(--paper);">
        <div class="field-row" style="margin-bottom:8px;">
          <div class="field" style="flex:2;margin:0;">
            <label style="font-size:11px;">Name</label>
            <input id="rec-new-name" placeholder="e.g. Office Rent" style="width:100%;">
          </div>
          <div class="field" style="flex:1;margin:0;">
            <label style="font-size:11px;">Amount (₹)</label>
            <input type="number" id="rec-new-amount" value="0" style="width:100%;">
          </div>
        </div>
        <div class="field-row">
          <div class="field" style="flex:1;margin:0;">
            <label style="font-size:11px;">Frequency</label>
            <select id="rec-new-freq" style="width:100%;">
              <option value="monthly">Monthly</option>
              <option value="quarterly">Quarterly</option>
              <option value="annual">Annual</option>
            </select>
          </div>
          <div class="field" style="flex:1;margin:0;">
            <label style="font-size:11px;">Start month</label>
            <select id="rec-new-startmonth" style="width:100%;">
              ${MONTHS_SHORT.map((m,i) => `<option value="${m}" ${i===0?'selected':''}>${m}</option>`).join('')}
            </select>
          </div>
          <div class="field" style="flex:0 0 auto;margin:0;display:flex;align-items:flex-end;gap:6px;">
            <button class="btn btn-sm btn-primary" id="rec-new-confirm">Add</button>
            <button class="btn btn-sm" id="rec-new-cancel">✕</button>
          </div>
        </div>
      </div>`;
    const listEl = document.getElementById('rec-edit-list');
    if (listEl) listEl.insertAdjacentHTML('afterend', addForm);
    document.getElementById('rec-add-new').style.display = 'none';
    document.getElementById('rec-new-cancel').onclick = () => {
      document.getElementById('rec-add-form').remove();
      document.getElementById('rec-add-new').style.display = '';
    };
    document.getElementById('rec-new-confirm').onclick = () => {
      const name = document.getElementById('rec-new-name').value.trim();
      if (!name) { toast('Enter a name'); return; }
      const amount = parseFloat(document.getElementById('rec-new-amount').value) || 0;
      const frequency = document.getElementById('rec-new-freq').value;
      const startMonth = document.getElementById('rec-new-startmonth').value;
      DB.recurringTemplate.push({ name, amount, frequency, startMonth, tds: false });
      renderRecurringEditor();
    };
  };

  document.querySelectorAll('[data-rec-del]').forEach(b => {
    b.onclick = () => {
      DB.recurringTemplate.splice(parseInt(b.dataset.recDel, 10), 1);
      renderRecurringEditor();
    };
  });

  document.getElementById('rec-save').onclick = () => {
    document.querySelectorAll('[data-rec-idx]').forEach(inp => {
      const idx = parseInt(inp.dataset.recIdx, 10);
      DB.recurringTemplate[idx].amount = parseFloat(inp.value) || 0;
    });
    const touched = applyRecurringForward(DB.selectedMonth);
    touched.forEach(mk => ensureMonthlyProvisions(mk));
    saveDB(touched); renderAll(); closeModal();
    toast(`Recurring payments updated from ${monthLabel(DB.selectedMonth)} onward`);
  };
}

// Pushes current recurringTemplate into the selected month and all later months in the
// same FY — never touches months before fromMk. Returns the touched month keys.
// Respects frequency: monthly=every month, quarterly=every 3rd month from startMonth,
// annual=only the startMonth each year.
function applyRecurringForward(fromMk) {
  const months = monthsOfFY(fyLabelForMonth(fromMk)).filter(mk => mk >= fromMk);

  months.forEach(mk => {
    const m = ensureMonthExists(mk);
    const mkDate = new Date(mk + '-01');
    const mkMonthIdx = mkDate.getMonth();
    const mkMonthShort = MONTHS_SHORT[mkMonthIdx];

    DB.recurringTemplate.forEach(tpl => {
      const freq = tpl.frequency || 'monthly';
      const startMonth = tpl.startMonth || MONTHS_SHORT[0];
      const startIdx = MONTHS_SHORT.indexOf(startMonth);

      let applies = false;
      if (freq === 'monthly') {
        applies = true;
      } else if (freq === 'quarterly') {
        const diff = (mkMonthIdx - startIdx + 12) % 12;
        applies = diff % 3 === 0;
      } else if (freq === 'annual') {
        applies = mkMonthShort === startMonth;
      }
      if (!applies) return;

      // Match by template name (strip month suffix from existing entries for comparison)
      const existing = m.payments.find(p => p.recurring &&
        p.name.toLowerCase().replace(/\s+for\s+.+$/i, '') === tpl.name.toLowerCase());
      if (existing) {
        existing.amount = tpl.amount;
        existing.frequency = freq; // ensure frequency is up to date
      } else {
        m.payments.push({
          id: uid(),
          name: `${tpl.name} for ${monthLabel(mk)}`,
          amount: tpl.amount,
          status: 'planned',
          recurring: true,
          frequency: freq,
          tds: !!tpl.tds,
        });
      }
    });
  });
  return months;
}

/* ===========================================================
   ACTION: Schedule Payment (future commitment, no amount yet)
   =========================================================== */
function openSchedulePayment() {
  const months = monthsOfFY(DB.currentFY).filter(mk => mk >= todayMonthKey() || mk >= DB.selectedMonth);
  const body = `
    <div class="field"><label>Vendor name</label><input type="text" id="sp-vendor" placeholder="e.g. New supplier"></div>
    <div class="field"><label>Target month</label>
      <select id="sp-month">${monthsOfFY(DB.currentFY).map(mk => `<option value="${mk}" ${mk===DB.selectedMonth?'selected':''}>${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="field"><label>Amount (leave blank if not yet known)</label><input type="number" id="sp-amount" placeholder="0"></div>
    <div class="checkrow"><input type="checkbox" id="sp-tds"><label for="sp-tds">TDS applicable</label></div>`;
  openModal('Schedule payment', body, `<button class="btn" id="sp-cancel">Cancel</button><button class="btn btn-primary" id="sp-save">Save</button>`);
  document.getElementById('sp-cancel').onclick = closeModal;
  document.getElementById('sp-save').onclick = () => {
    const vendor = document.getElementById('sp-vendor').value.trim();
    const mk = document.getElementById('sp-month').value;
    const amount = parseFloat(document.getElementById('sp-amount').value) || 0;
    const tds = document.getElementById('sp-tds').checked;
    if (!vendor) { toast('Vendor name is required'); return; }
    const m = ensureMonthExists(mk);
    m.payments.push({ id: uid(), name: vendor, amount, status: 'planned', recurring: false, tds });
    saveDB([mk]); renderAll(); closeModal();
    toast(`Scheduled ${vendor} for ${monthLabel(mk)}`);
  };
}

/* ===========================================================
   ACTION: Schedule Standard Payment (from vendor dropdown)
   =========================================================== */
function openScheduleStandard() {
  const vendors = DB.standardVendors;
  const body = `
    <div class="field"><label>Standard vendor</label>
      <select id="ss-vendor">${vendors.map(v => `<option value="${v.id}">${escapeHtml(v.name)}${v.amount ? ' &mdash; ' + fmtMoney(v.amount) : ''}</option>`).join('')}</select>
    </div>
    <div class="field"><label>Month</label>
      <select id="ss-month">${monthsOfFY(DB.currentFY).map(mk => `<option value="${mk}" ${mk===DB.selectedMonth?'selected':''}>${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="field"><label>Amount</label><input type="number" id="ss-amount" placeholder="0"></div>`;
  openModal('Schedule standard payment', body, `<button class="btn" id="ss-cancel">Cancel</button><button class="btn btn-primary" id="ss-save">Save</button>`);

  const sel = document.getElementById('ss-vendor');
  const amtInput = document.getElementById('ss-amount');
  function fillAmt() {
    const v = vendors.find(v => v.id === sel.value);
    amtInput.value = v && v.amount ? v.amount : '';
  }
  sel.addEventListener('change', fillAmt);
  fillAmt();

  document.getElementById('ss-cancel').onclick = closeModal;
  document.getElementById('ss-save').onclick = () => {
    const v = vendors.find(v => v.id === sel.value);
    const mk = document.getElementById('ss-month').value;
    const amount = parseFloat(amtInput.value) || 0;
    const m = ensureMonthExists(mk);
    m.payments.push({ id: uid(), name: v.name, amount, status: 'planned', recurring: false, tds: !!v.tds });
    saveDB([mk]); renderAll(); closeModal();
    toast(`Added ${v.name} payment for ${monthLabel(mk)}`);
  };
}

/* ===========================================================
   ACTION: Add Vendor
   =========================================================== */
function openAddVendor() {
  const body = `
    <div class="field"><label>Vendor name</label><input type="text" id="av-name" placeholder="e.g. New Courier Co."></div>
    <div class="radio-pills">
      <div class="radio-pill"><input type="radio" name="av-type" id="av-onetime" value="onetime" checked><label for="av-onetime">One time</label></div>
      <div class="radio-pill"><input type="radio" name="av-type" id="av-regular" value="regular"><label for="av-regular">Regular</label></div>
    </div>
    <div class="field"><label>Default amount (optional)</label><input type="number" id="av-amount" placeholder="0"></div>
    <div class="hint">Regular vendors are added to the Standard Payments dropdown for future months.</div>`;
  openModal('Add vendor', body, `<button class="btn" id="av-cancel">Cancel</button><button class="btn btn-primary" id="av-save">Save</button>`);
  document.getElementById('av-cancel').onclick = closeModal;
  document.getElementById('av-save').onclick = () => {
    const name = document.getElementById('av-name').value.trim();
    const amount = parseFloat(document.getElementById('av-amount').value) || 0;
    const isRegular = document.getElementById('av-regular').checked;
    if (!name) { toast('Vendor name is required'); return; }
    if (isRegular) {
      DB.standardVendors.push({ id: uid(), name, amount, tds: false });
      toast(`${name} added to Standard Payments`);
    } else {
      const m = getMonth(DB.selectedMonth);
      m.payments.push({ id: uid(), name, amount, status: 'planned', recurring: false, tds: false });
      toast(`${name} added as a one-time payment for ${monthLabel(DB.selectedMonth)}`);
    }
    saveDB(); renderAll(); closeModal();
  };
}

/* ===========================================================
   ACTION: General Payment
   =========================================================== */
function openGeneralPayment() {
  const body = `
    <div class="field"><label>Pay to</label><input type="text" id="gp-payto" placeholder="e.g. Vendor / individual"></div>
    <div class="field"><label>Amount</label><input type="number" id="gp-amount" placeholder="0"></div>
    <div class="checkrow"><input type="checkbox" id="gp-tds"><label for="gp-tds">TDS applicable (optional)</label></div>`;
  openModal('General payment', body, `<button class="btn" id="gp-cancel">Cancel</button><button class="btn btn-primary" id="gp-save">Save</button>`);
  document.getElementById('gp-cancel').onclick = closeModal;
  document.getElementById('gp-save').onclick = () => {
    const payTo = document.getElementById('gp-payto').value.trim();
    const amount = parseFloat(document.getElementById('gp-amount').value) || 0;
    const tds = document.getElementById('gp-tds').checked;
    if (!payTo) { toast('Pay to is required'); return; }
    const m = getMonth(DB.selectedMonth);
    m.payments.push({ id: uid(), name: payTo, amount, status: 'paid', recurring: false, tds });
    if (tds) bumpNextMonthTDS(amount * 0.1 > 0 ? null : null); // TDS bump handled by recalcTDS() globally on render
    saveDB(); renderAll(); closeModal();
    toast(`Recorded payment of ${fmtMoney(amount)} to ${payTo}`);
  };
}

// All payments flagged TDS in a month roll into a TDS line in the next month (per rules: "ALL TDS in payments gets auto added to next month TDS amount")
function bumpNextMonthTDS() { /* superseded by recalcAllTDS, kept as no-op for compatibility */ }


/* ===========================================================
   ACTION: Import Order (Kerry / Ceva estimates)
   =========================================================== */
function openImportOrder() {
  const body = `
    <div class="field"><label>Order month</label>
      <select id="io-month">${monthsOfFY(DB.currentFY).map(mk => `<option value="${mk}" ${mk===DB.selectedMonth?'selected':''}>${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="field-row">
      <div class="field"><label>Quantity</label><input type="number" id="io-qty" placeholder="0"></div>
      <div class="field"><label>Order amount</label><input type="number" id="io-amount" placeholder="0"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>KERRY estimate</label><input type="number" id="io-kerry" placeholder="0"></div>
      <div class="field"><label>CEVA estimate</label><input type="number" id="io-ceva" placeholder="0"></div>
    </div>`;
  openModal('Import order', body, `<button class="btn" id="io-cancel">Cancel</button><button class="btn btn-primary" id="io-save">Save</button>`);
  document.getElementById('io-cancel').onclick = closeModal;
  document.getElementById('io-save').onclick = () => {
    const mk = document.getElementById('io-month').value;
    const qty = parseFloat(document.getElementById('io-qty').value) || 0;
    const amount = parseFloat(document.getElementById('io-amount').value) || 0;
    const kerry = parseFloat(document.getElementById('io-kerry').value) || 0;
    const ceva = parseFloat(document.getElementById('io-ceva').value) || 0;
    const m = ensureMonthExists(mk);
    if (!m.imports) m.imports = [];
    const orderNum = m.imports.length + 1;
    const id = uid();
    m.imports.push({ id, qty, amount, kerry, ceva });
    // also reflect in payments ledger as line items so they show in the cashflow
    if (amount) m.payments.push({ id: uid(), name: `Import order ${orderNum} (Qty ${qty})`, amount, status: 'planned', recurring: false, tds: false, _importId: id });
    if (kerry) m.payments.push({ id: uid(), name: `Kerry import order ${orderNum}`, amount: kerry, status: 'planned', recurring: false, tds: false, _importId: id });
    if (ceva) m.payments.push({ id: uid(), name: `Ceva customs order ${orderNum}`, amount: ceva, status: 'planned', recurring: false, tds: false, _importId: id });
    saveDB([mk]); renderAll(); closeModal();
    toast(`Import order ${orderNum} added for ${monthLabel(mk)}`);
  };
}

/* ===========================================================
   ACTION: Ad Hoc Receivable
   A miscellaneous receivable with no PI/GST impact —
   e.g. tax refund, reimbursement from third party.
   Carries forward like a regular receivable if unpaid.
   =========================================================== */
function openAdHocReceivable() {
  const months = monthsOfFY(DB.currentFY);
  const body = `
    <div class="field"><label>Description</label>
      <input type="text" id="ahr-name" placeholder="e.g. Tax refund, Customs duty refund">
    </div>
    <div class="field"><label>Amount</label>
      <input type="number" id="ahr-amount" placeholder="0">
    </div>
    <div class="field"><label>Month</label>
      <select id="ahr-month">${months.map(mk => `<option value="${mk}" ${mk===DB.selectedMonth?'selected':''}>${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="hint">No GST or deduction applied — amount entered is the full receivable amount. Carries forward to the following month if unpaid, same as a regular receivable.</div>`;
  openModal('Add ad hoc receivable', body, `<button class="btn" id="ahr-cancel">Cancel</button><button class="btn btn-primary" id="ahr-save">Save</button>`);
  document.getElementById('ahr-cancel').onclick = closeModal;
  document.getElementById('ahr-save').onclick = () => {
    const name = document.getElementById('ahr-name').value.trim();
    const amount = parseFloat(document.getElementById('ahr-amount').value) || 0;
    const mk = document.getElementById('ahr-month').value;
    if (!name) { toast('Description is required'); return; }
    if (!amount) { toast('Amount is required'); return; }
    const m = ensureMonthExists(mk);
    m.receivables.push({ id: uid(), name, amount, _adhoc: true });
    if (mk === todayMonthKey()) syncCarriedReceivables();
    saveDB([mk, nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
    toast(`Added ad hoc receivable: ${name} (${fmtMoney(amount)})`);
  };
}

function openEditImport(importId) {
  const m = getMonth(DB.selectedMonth);
  const o = (m.imports || []).find(o => o.id === importId);
  if (!o) return;
  const body = `
    <div class="field-row">
      <div class="field"><label>Quantity</label><input type="number" id="ei-qty" value="${o.qty}"></div>
      <div class="field"><label>Order amount</label><input type="number" id="ei-amount" value="${o.amount}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>KERRY estimate</label><input type="number" id="ei-kerry" value="${o.kerry}"></div>
      <div class="field"><label>CEVA estimate</label><input type="number" id="ei-ceva" value="${o.ceva}"></div>
    </div>`;
  openModal('Edit import order', body, `<button class="btn" id="ei-cancel">Cancel</button><button class="btn btn-primary" id="ei-save">Save</button>`);
  document.getElementById('ei-cancel').onclick = closeModal;
  document.getElementById('ei-save').onclick = () => {
    o.qty = parseFloat(document.getElementById('ei-qty').value) || 0;
    o.amount = parseFloat(document.getElementById('ei-amount').value) || 0;
    o.kerry = parseFloat(document.getElementById('ei-kerry').value) || 0;
    o.ceva = parseFloat(document.getElementById('ei-ceva').value) || 0;
    // sync linked payment rows
    m.payments = m.payments.filter(p => p._importId !== o.id);
    const orderNum = m.imports.indexOf(o) + 1;
    if (o.amount) m.payments.push({ id: uid(), name: `Import order ${orderNum} (Qty ${o.qty})`, amount: o.amount, status: 'planned', recurring: false, tds: false, _importId: o.id });
    if (o.kerry) m.payments.push({ id: uid(), name: `Kerry import order ${orderNum}`, amount: o.kerry, status: 'planned', recurring: false, tds: false, _importId: o.id });
    if (o.ceva) m.payments.push({ id: uid(), name: `Ceva customs order ${orderNum}`, amount: o.ceva, status: 'planned', recurring: false, tds: false, _importId: o.id });
    saveDB(); renderAll(); closeModal();
  };
}
function deleteImport(importId) {
  if (!confirm('Delete this import order?')) return;
  const m = getMonth(DB.selectedMonth);
  m.imports = (m.imports || []).filter(o => o.id !== importId);
  m.payments = m.payments.filter(p => p._importId !== importId);
  saveDB(); renderAll();
}

/* ===========================================================
   ACTION: Fixed Deposits
   =========================================================== */
function openFdAdd() {
  const body = `
    <div class="field"><label>Type</label><input type="text" id="fd-type" placeholder="e.g. sweepin FD"></div>
    <div class="field-row">
      <div class="field"><label>Amount</label><input type="number" id="fd-amount" placeholder="0"></div>
      <div class="field"><label>Term</label><input type="text" id="fd-term" placeholder="e.g. 12 months"></div>
    </div>`;
  openModal('Add fixed deposit', body, `<button class="btn" id="fd-cancel">Cancel</button><button class="btn btn-primary" id="fd-save">Save</button>`);
  document.getElementById('fd-cancel').onclick = closeModal;
  document.getElementById('fd-save').onclick = () => {
    const type = document.getElementById('fd-type').value.trim();
    const amount = parseFloat(document.getElementById('fd-amount').value) || 0;
    const term = document.getElementById('fd-term').value.trim();
    if (!type) { toast('Type is required'); return; }
    DB.fixedDeposits.push({ id: uid(), type, amount, term });
    saveDB(); renderAll(); closeModal();
    toast(`Fixed deposit added: ${type}`);
  };
}
function openEditFd(fdId) {
  const fd = (DB.fixedDeposits || []).find(f => f.id === fdId);
  if (!fd) return;
  const body = `
    <div class="field"><label>Type</label><input type="text" id="fd-type" value="${escapeHtml(fd.type)}"></div>
    <div class="field-row">
      <div class="field"><label>Amount</label><input type="number" id="fd-amount" value="${fd.amount}"></div>
      <div class="field"><label>Term</label><input type="text" id="fd-term" value="${escapeHtml(fd.term || '')}"></div>
    </div>
    <div class="hint">This is a standing balance — changes apply everywhere, not just one month. If this FD has matured or closed, use Delete.</div>`;
  openModal('Edit fixed deposit', body, `<button class="btn btn-danger" id="fd-delete">Delete</button><button class="btn" id="fd-cancel">Cancel</button><button class="btn btn-primary" id="fd-save">Save</button>`);
  document.getElementById('fd-cancel').onclick = closeModal;
  document.getElementById('fd-delete').onclick = () => {
    if (!confirm(`Delete "${fd.type}"? This removes it everywhere.`)) return;
    DB.fixedDeposits = DB.fixedDeposits.filter(f => f.id !== fdId);
    saveDB(); renderAll(); closeModal();
    toast(`Fixed deposit removed: ${fd.type}`);
  };
  document.getElementById('fd-save').onclick = () => {
    const type = document.getElementById('fd-type').value.trim();
    const amount = parseFloat(document.getElementById('fd-amount').value) || 0;
    const term = document.getElementById('fd-term').value.trim();
    if (!type) { toast('Type is required'); return; }
    fd.type = type; fd.amount = amount; fd.term = term;
    saveDB(); renderAll(); closeModal();
    toast(`Fixed deposit updated: ${type}`);
  };
}
function openBgAdd() {
  const body = `
    <div class="field"><label>Name</label><input type="text" id="bg-name" placeholder="e.g. FD bank guarantee"></div>
    <div class="field"><label>Amount</label><input type="number" id="bg-amount" placeholder="0"></div>`;
  openModal('Add bank guarantee', body, `<button class="btn" id="bg-cancel">Cancel</button><button class="btn btn-primary" id="bg-save">Save</button>`);
  document.getElementById('bg-cancel').onclick = closeModal;
  document.getElementById('bg-save').onclick = () => {
    const name = document.getElementById('bg-name').value.trim();
    const amount = parseFloat(document.getElementById('bg-amount').value) || 0;
    if (!name) { toast('Name is required'); return; }
    DB.bankGuarantees.push({ id: uid(), name, amount });
    saveDB(); renderAll(); closeModal();
    toast(`Bank guarantee added: ${name}`);
  };
}
function openEditBg(bgId) {
  const bg = (DB.bankGuarantees || []).find(b => b.id === bgId);
  if (!bg) return;
  const body = `
    <div class="field"><label>Name</label><input type="text" id="bg-name" value="${escapeHtml(bg.name)}"></div>
    <div class="field"><label>Amount</label><input type="number" id="bg-amount" value="${bg.amount}"></div>
    <div class="hint">This is a standing balance — changes apply everywhere, not just one month. If this guarantee has been released, use Delete.</div>`;
  openModal('Edit bank guarantee', body, `<button class="btn btn-danger" id="bg-delete">Delete</button><button class="btn" id="bg-cancel">Cancel</button><button class="btn btn-primary" id="bg-save">Save</button>`);
  document.getElementById('bg-cancel').onclick = closeModal;
  document.getElementById('bg-delete').onclick = () => {
    if (!confirm(`Delete "${bg.name}"? This removes it everywhere.`)) return;
    DB.bankGuarantees = DB.bankGuarantees.filter(b => b.id !== bgId);
    saveDB(); renderAll(); closeModal();
    toast(`Bank guarantee removed: ${bg.name}`);
  };
  document.getElementById('bg-save').onclick = () => {
    const name = document.getElementById('bg-name').value.trim();
    const amount = parseFloat(document.getElementById('bg-amount').value) || 0;
    if (!name) { toast('Name is required'); return; }
    bg.name = name; bg.amount = amount;
    saveDB(); renderAll(); closeModal();
    toast(`Bank guarantee updated: ${name}`);
  };
}

/* ===========================================================
   EDIT / DELETE row-level handlers
   =========================================================== */
function openEditReceipt(id) {
  const m = getMonth(DB.selectedMonth);
  const r = m.receipts.find(r => r.id === id);
  if (!r) return;
  const isLinkedPart = !!r._receivableId;
  const body = `
    <div class="field"><label>Name</label><input type="text" id="er-name" value="${escapeHtml(r.name)}"></div>
    <div class="field"><label>Amount</label><input type="number" id="er-amount" value="${r.amount}"></div>
    <div class="field"><label>Status</label>
      <select id="er-status">
        <option value="RECD" ${r.status==='RECD'?'selected':''}>RECD</option>
        <option value="expected" ${r.status==='expected'?'selected':''}>Expected</option>
      </select>
    </div>
    ${isLinkedPart ? '<div class="hint">This was a part payment against a receivable. Changing the amount here will adjust that receivable\'s outstanding balance to match.</div>' : ''}`;
  openModal('Edit receipt', body, `<button class="btn" id="er-cancel">Cancel</button><button class="btn btn-primary" id="er-save">Save</button>`);
  document.getElementById('er-cancel').onclick = closeModal;
  document.getElementById('er-save').onclick = () => {
    const oldAmount = Number(r.amount) || 0;
    const newAmount = parseFloat(document.getElementById('er-amount').value) || 0;
    r.name = document.getElementById('er-name').value.trim();
    r.amount = newAmount;
    r.status = document.getElementById('er-status').value;

    if (isLinkedPart) {
      const delta = newAmount - oldAmount;
      const recMonth = getMonth(r._receivableMonth || DB.selectedMonth);
      let rec = recMonth.receivables.find(rv => rv.id === r._receivableId);
      if (!rec && delta < 0) {
        rec = { id: r._receivableId, name: r.name, amount: 0 };
        recMonth.receivables.push(rec);
      }
      if (rec) {
        rec.amount = Math.max(0, (Number(rec.amount) || 0) - delta);
        if (rec.amount === 0) {
          recMonth.receivables = recMonth.receivables.filter(rv => rv.id !== rec.id);
        }
      }
    }

    const recvMonth = r._receivableMonth || DB.selectedMonth;
    if (recvMonth === todayMonthKey()) syncCarriedReceivables();
    saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
  };
}
function deleteReceipt(id) {
  const m = getMonth(DB.selectedMonth);
  const r = m.receipts.find(r => r.id === id);
  if (!r) return;

  const nextMk = nextMonthKey(DB.selectedMonth);
  const nextM = DB.months[nextMk];
  const existingGst = nextM ? nextM.payments.find(p => /^GST/i.test(p.name)) : null;
  const gstHint = existingGst
    ? `Next month currently has a GST line of ${fmtMoney(existingGst.amount)}. Any amount you enter here will be deducted from it.`
    : `Next month has no GST line yet — if you enter an amount it will be skipped (nothing to deduct from).`;

  const body = `
    <div class="field">
      <label>Receipt being deleted</label>
      <div style="padding:8px 0; font-weight:600;">${escapeHtml(r.name)} — ${fmtMoney(r.amount)}</div>
    </div>
    <div class="field">
      <label>GST included in this receipt (enter 0 if none)</label>
      <input type="number" id="del-gst" value="0" min="0">
    </div>
    <div class="hint">${gstHint}</div>`;

  openModal('Delete receipt', body,
    `<button class="btn" id="del-cancel">Cancel</button>
     <button class="btn btn-danger" id="del-confirm">Delete receipt</button>`);

  document.getElementById('del-cancel').onclick = closeModal;
  document.getElementById('del-confirm').onclick = () => {
    const gstAmt = parseFloat(document.getElementById('del-gst').value) || 0;

    // Restore receivable if this was a linked payment
    if (r._receivableId) {
      const recMonth = getMonth(r._receivableMonth || DB.selectedMonth);
      let rec = recMonth.receivables.find(rv => rv.id === r._receivableId);
      if (!rec) {
        rec = { id: r._receivableId, name: r.name, amount: 0 };
        recMonth.receivables.push(rec);
      }
      rec.amount = (Number(rec.amount) || 0) + (Number(r.amount) || 0);
      const restoredTo = r._receivableMonth || DB.selectedMonth;
      if (restoredTo === todayMonthKey()) syncCarriedReceivables();
    }

    // Deduct GST from next month's GST line if amount entered and line exists
    const extraMks = [nextMonthKey(todayMonthKey())];
    if (gstAmt > 0 && existingGst) {
      existingGst.amount = Math.max(0, (Number(existingGst.amount) || 0) - gstAmt);
      if (!extraMks.includes(nextMk)) extraMks.push(nextMk);
      toast(`Receipt deleted — GST of ${fmtMoney(gstAmt)} deducted from ${monthLabel(nextMk)} GST line`);
    } else {
      toast(`Receipt deleted${r._receivableId ? ' — amount restored to receivables' : ''}`);
    }

    m.receipts = m.receipts.filter(x => x.id !== id);
    saveDB(extraMks);
    closeModal();
    setTimeout(() => renderAll(), 50);
  };
}

function openEditPayment(id) {
  const m = getMonth(DB.selectedMonth);
  const p = m.payments.find(p => p.id === id);
  if (!p) return;
  const months = monthsOfFY(DB.currentFY).filter(mk => mk !== DB.selectedMonth);
  const body = `
    <div class="field"><label>Name</label><input type="text" id="ep-name" value="${escapeHtml(p.name)}"></div>
    <div class="field"><label>Amount</label><input type="number" id="ep-amount" value="${p.amount}"></div>
    <div class="field"><label>Status</label>
      <select id="ep-status">
        <option value="planned" ${p.status==='planned'?'selected':''}>Planned</option>
        <option value="paid" ${p.status==='paid'?'selected':''}>Paid</option>
        <option value="on hold" ${p.status==='on hold'?'selected':''}>On hold</option>
        <option value="schedule">Schedule for another month&hellip;</option>
      </select>
    </div>
    <div class="field" id="ep-schedule-field" style="display:none;">
      <label>Move to month</label>
      <select id="ep-schedule-month">${months.map(mk => `<option value="${mk}">${monthLabel(mk)}</option>`).join('')}</select>
    </div>
    <div class="checkrow"><input type="checkbox" id="ep-tds" ${p.tds?'checked':''}><label for="ep-tds">TDS applicable</label></div>
    <div class="hint" id="ep-hint">${p.recurring ? 'This is a recurring item. To change it for all future months, use Edit recurring instead. Saving here changes only ' + monthLabel(DB.selectedMonth) + '.' : 'Edits here apply only to ' + monthLabel(DB.selectedMonth) + '.'}</div>`;
  openModal('Edit payment', body, `<button class="btn" id="ep-cancel">Cancel</button><button class="btn btn-primary" id="ep-save">Save</button>`);
  document.getElementById('ep-cancel').onclick = closeModal;

  const statusSel = document.getElementById('ep-status');
  const scheduleField = document.getElementById('ep-schedule-field');
  const hintEl = document.getElementById('ep-hint');
  statusSel.addEventListener('change', () => {
    const isSchedule = statusSel.value === 'schedule';
    scheduleField.style.display = isSchedule ? 'block' : 'none';
    if (isSchedule) hintEl.textContent = 'This will remove the payment from ' + monthLabel(DB.selectedMonth) + ' and add it to the month you pick below.';
  });

  document.getElementById('ep-save').onclick = () => {
    const newStatus = statusSel.value;
    if (newStatus === 'schedule') {
      const targetMk = document.getElementById('ep-schedule-month').value;
      const name = document.getElementById('ep-name').value.trim();
      const amount = parseFloat(document.getElementById('ep-amount').value) || 0;
      const tds = document.getElementById('ep-tds').checked;
      m.payments = m.payments.filter(x => x.id !== id);
      const target = ensureMonthExists(targetMk);
      target.payments.push({ id: uid(), name, amount, status: 'planned', recurring: p.recurring, tds });
      saveDB([targetMk]); renderAll(); closeModal();
      toast(`Moved "${name}" to ${monthLabel(targetMk)}`);
      return;
    }
    p.name = document.getElementById('ep-name').value.trim();
    p.amount = parseFloat(document.getElementById('ep-amount').value) || 0;
    p.status = newStatus;
    p.tds = document.getElementById('ep-tds').checked;
    saveDB(); renderAll(); closeModal();
  };
}
function deletePayment(id) {
  if (!confirm('Delete this payment?')) return;
  const m = getMonth(DB.selectedMonth);
  m.payments = m.payments.filter(p => p.id !== id);
  saveDB(); renderAll();
}
function openEditReceivable(id) {
  const m = getMonth(DB.selectedMonth);
  const r = m.receivables.find(r => r.id === id);
  if (!r) return;
  const body = `
    <div class="field"><label>Licensee / Item</label><input type="text" id="erv-name" value="${escapeHtml(r.name)}"></div>
    <div class="field"><label>Amount</label><input type="number" id="erv-amount" value="${r.amount}"></div>`;
  openModal('Edit receivable', body, `<button class="btn btn-danger" id="erv-delete">Delete</button><button class="btn" id="erv-cancel">Cancel</button><button class="btn btn-primary" id="erv-save">Save</button>`);
  document.getElementById('erv-cancel').onclick = closeModal;
  document.getElementById('erv-delete').onclick = () => {
    if (!confirm(`Delete "${r.name}"?`)) return;
    m.receivables = m.receivables.filter(x => x.id !== id);
    if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
    saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
    toast(`Receivable removed: ${r.name}`);
  };
  document.getElementById('erv-save').onclick = () => {
    const name = document.getElementById('erv-name').value.trim();
    const amount = parseFloat(document.getElementById('erv-amount').value) || 0;
    if (!name) { toast('Name is required'); return; }
    r.name = name; r.amount = amount;
    if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
    saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
    toast(`Receivable updated: ${name}`);
  };
}
function deleteReceivable(id) {
  const m = getMonth(DB.selectedMonth);
  const r = m.receivables.find(rv => rv.id === id);
  // Check if this receivable is linked to an invoice
  if (r && r._invoiceId) {
    invInit();
    const linkedInv = [...(DB.invoicing.proformas||[]), ...(DB.invoicing.taxInvoices||[])]
      .find(inv => inv.id === r._invoiceId);
    if (linkedInv && linkedInv.status !== 'cancelled') {
      openModal('Delete receivable', `
        <div class="hint" style="margin-bottom:10px;">
          This receivable is linked to <strong>${escapeHtml(linkedInv.invNo)}</strong>.<br>
          Would you also like to cancel that invoice in the Invoicing register?
        </div>`,
        `<button class="btn" id="dr-cancel-modal">Cancel (do nothing)</button>
         <button class="btn" id="dr-del-only">Delete receivable only</button>
         <button class="btn btn-danger" id="dr-del-and-inv">Delete receivable + Cancel invoice</button>`);
      document.getElementById('dr-cancel-modal').onclick = closeModal;
      document.getElementById('dr-del-only').onclick = () => {
        m.receivables = m.receivables.filter(rv => rv.id !== id);
        if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
        saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
      };
      document.getElementById('dr-del-and-inv').onclick = () => {
        m.receivables = m.receivables.filter(rv => rv.id !== id);
        linkedInv.status = 'cancelled';
        if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
        saveDB([nextMonthKey(todayMonthKey())]); renderAll(); closeModal();
        toast(`Receivable deleted and ${linkedInv.invNo} cancelled`);
      };
      return;
    }
  }
  if (!confirm('Delete this receivable?')) return;
  m.receivables = m.receivables.filter(rv => rv.id !== id);
  if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
  saveDB([nextMonthKey(todayMonthKey())]); renderAll();
}


/* ===========================================================
   AUTOMATIC PROVISIONS
   - TDS provisional and Waco marketing fee are now in DB.recurringTemplate
     (monthly and quarterly respectively) — migrated on startup.
   - TDS flagged on payments rolls into next month's TDS provisional line.
   =========================================================== */
function ensureMonthlyProvisions(mk) {
  const m = ensureMonthExists(mk);

  // Mark as provisioned — prevents re-provisioning if user deletes items they don't want
  m._provisioned = true;

  // Recurring (starred) payments — respect frequency and startMonth
  const mkDate = new Date(mk + '-01');
  const mkMonthIdx = mkDate.getMonth(); // 0=Jan
  const mkMonthShort = MONTHS_SHORT[mkMonthIdx];

  DB.recurringTemplate.forEach(tpl => {
    const freq = tpl.frequency || 'monthly';
    const startMonth = tpl.startMonth || MONTHS_SHORT[0];
    const startIdx = MONTHS_SHORT.indexOf(startMonth);

    let applies = false;
    if (freq === 'monthly') {
      applies = true;
    } else if (freq === 'quarterly') {
      const diff = (mkMonthIdx - startIdx + 12) % 12;
      applies = diff % 3 === 0;
    } else if (freq === 'annual') {
      applies = mkMonthShort === startMonth;
    }
    if (!applies) return;

    const existing = m.payments.find(p => p.recurring &&
      p.name.toLowerCase().replace(/\s+for\s+.+$/i, '') === tpl.name.toLowerCase());
    if (!existing) {
      m.payments.push({
        id: uid(),
        name: `${tpl.name} for ${monthLabel(mk)}`,
        amount: tpl.amount,
        status: 'planned',
        recurring: true,
        frequency: freq,
        tds: !!tpl.tds,
      });
    }
  });

  return m;
}

// Roll TDS-flagged payment amounts (10%) from a month into next month's TDS provisional line, once.
function recalcTDSRollup(mk) {
  const m = DB.months[mk];
  if (!m) return;
  const flagged = m.payments.filter(p => p.tds && !p._tdsRolled);
  if (!flagged.length) return;
  const nextMk = nextMonthKey(mk);
  const nm = ensureMonthExists(nextMk);
  let tdsLine = nm.payments.find(p => /^TDS provisional/i.test(p.name));
  if (!tdsLine) {
    // Get amount from template if available, else use constant
    const tdsTpl = (DB.recurringTemplate || []).find(t => /tds provisional/i.test(t.name));
    const defaultAmt = tdsTpl ? tdsTpl.amount : TDS_ESTIMATE;
    tdsLine = { id: uid(), name: 'TDS provisional', amount: defaultAmt, status: 'planned', recurring: true, frequency: 'monthly', tds: false };
    nm.payments.push(tdsLine);
  }
  let addAmt = 0;
  flagged.forEach(p => { addAmt += (Number(p.amount) || 0) * 0.10; p._tdsRolled = true; });
  tdsLine.amount = (Number(tdsLine.amount) || 0) + addAmt;
}

/* ===========================================================
   BACKUP / RESTORE
   =========================================================== */
function exportBackup() {
  const blob = new Blob([JSON.stringify(DB, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `LMI_Cashflow_backup_${stamp}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  toast('Backup downloaded');
}
function importBackup(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const parsed = JSON.parse(e.target.result);
      if (!parsed.months) throw new Error('Invalid file');
      if (!confirm('This will replace all current data with the backup file. Continue?')) return;
      DB = parsed;
      saveDB(); renderAll();
      toast('Backup restored');
    } catch (err) {
      alert('Could not read this file as a valid backup.');
    }
  };
  reader.readAsText(file);
}

/* ===========================================================
   INIT
   =========================================================== */
function wireActionBar() {
  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      ({
        addInvoice: openAddInvoice,
        paymentReceived: openPaymentReceived,
        generalPayment: openGeneralPayment,
        schedulePayment: openSchedulePayment,
        scheduleStandard: openScheduleStandard,
        addVendor: openAddVendor,
        importOrder: openImportOrder,
        adHocReceivable: openAdHocReceivable,
        scenarioPlanner: openScenarioPlanner,
        generatePI: invOpenPIForm,
        generateTI: () => invOpenTIForm(null, null, true),
      })[action]?.();
    });
  });
}

function wireStaticButtons() {
  document.getElementById('btnUpdateOpening').onclick = openUpdateOpening;
  document.getElementById('btnEditRecurring').onclick = openEditRecurring;
  document.getElementById('btnFdAdd').onclick = openFdAdd;
  document.getElementById('btnBgAdd').onclick = openBgAdd;
  document.getElementById('btnExport').onclick = exportBackup;
  document.getElementById('btnImport').onclick = () => document.getElementById('fileImport').click();
  document.getElementById('fileImport').onchange = e => {
    if (e.target.files[0]) importBackup(e.target.files[0]);
    e.target.value = '';
  };
  document.getElementById('btnTestMode').onclick = activateTestMode;
  document.getElementById('btnUndoTest').onclick = undoTestMode;
  const addUserBtn = document.getElementById('btnAddUser');
  const signOutBtn = document.getElementById('btnSignOut');
  if (window.Cloud && Cloud.cloudConfigured()) {
    addUserBtn.style.display = '';
    signOutBtn.style.display = '';
    addUserBtn.onclick = openAddUser;
    signOutBtn.onclick = async () => {
      if (!confirm('Sign out of LMI Cashflow Manager?')) return;
      await Cloud.signOut();
      location.reload();
    };
  } else {
    addUserBtn.style.display = 'none';
    signOutBtn.style.display = 'none';
  }
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });
}

function init() {
  DB = loadDB();
  if (!DB.selectedMonth) DB.selectedMonth = todayMonthKey();
  if (!DB.currentFY) DB.currentFY = fyLabelForMonth(DB.selectedMonth);
  ensureMonthExists(DB.selectedMonth);
  ensureMonthlyProvisions(DB.selectedMonth);
  recalcTDSRollup(prevMonthKey(DB.selectedMonth));
  promoteCarriedReceivables(); // unlock carried entries now that their month is current
  syncCarriedReceivables();
  saveDB([nextMonthKey(todayMonthKey())]);
  wireActionBar();
  wireStaticButtons();
  restoreTestModeIfActive();
  renderAll();
}

// Entry point used once a user is authenticated (or immediately, if running
// in local-only mode without Supabase configured).
async function startApp() {
  if (window.Cloud && Cloud.cloudConfigured() && Cloud.currentUser) {
    try {
      const cloudDB = await Cloud.cloudLoadAll();
      const hasAnyCloudData = Object.keys(cloudDB.months).length > 0;
      if (hasAnyCloudData) {
        DB = cloudDB;
      } else {
        DB = buildSeedDB();
        DB._workspaceId = cloudDB._workspaceId;
      }
      if (!DB.selectedMonth) DB.selectedMonth = todayMonthKey();
      if (!DB.currentFY) DB.currentFY = fyLabelForMonth(DB.selectedMonth);
      migrateRecurringFrequency();  // fix any wrongly-categorised recurring items
      ensureMonthExists(DB.selectedMonth);
      ensureMonthlyProvisions(DB.selectedMonth);
      recalcTDSRollup(prevMonthKey(DB.selectedMonth));
      promoteCarriedReceivables();
      syncCarriedReceivables();
      saveDB([prevMonthKey(DB.selectedMonth), nextMonthKey(todayMonthKey())]);
      Cloud.startRealtime();
    } catch (err) {
      console.error('Cloud load failed, falling back to local data', err);
      toast('Could not reach the cloud — showing your last local copy');
      init();
      return;
    }
    wireActionBar();
    wireStaticButtons();
    restoreTestModeIfActive();
    renderAll();
  } else {
    init();
  }
}

// Migration: ensure every recurring template entry has an explicit frequency.
// Also promotes the hardcoded TDS provisional and Waco marketing fee into
// DB.recurringTemplate so they appear in the left panel with proper frequency.
function migrateRecurringFrequency() {
  if (!DB.recurringTemplate) DB.recurringTemplate = [];

  // 1. Fix any template entries missing frequency → default monthly
  DB.recurringTemplate.forEach(tpl => {
    if (!tpl.frequency) tpl.frequency = 'monthly';
    if (/tds/i.test(tpl.name) && tpl.frequency !== 'monthly') tpl.frequency = 'monthly';
  });

  // 2. Promote TDS provisional into recurringTemplate if not already there
  const hasTdsTpl = DB.recurringTemplate.some(t => /tds provisional/i.test(t.name));
  if (!hasTdsTpl) {
    let tdsAmt = TDS_ESTIMATE;
    Object.values(DB.months || {}).forEach(m => {
      const row = (m.payments || []).find(p => /^TDS provisional/i.test(p.name));
      if (row && row.amount) tdsAmt = row.amount;
    });
    DB.recurringTemplate.push({ id: uid(), name: 'TDS provisional', amount: tdsAmt, frequency: 'monthly', startMonth: 'Apr', tds: false });
  }

  // 3. Promote Waco marketing fee into recurringTemplate if not already there
  const hasWacoTpl = DB.recurringTemplate.some(t => /waco marketing fee/i.test(t.name));
  if (!hasWacoTpl) {
    let wacoAmt = WACO_FEE;
    Object.values(DB.months || {}).forEach(m => {
      const row = (m.payments || []).find(p => /waco marketing fee/i.test(p.name));
      if (row && row.amount) wacoAmt = row.amount;
    });
    DB.recurringTemplate.push({ id: uid(), name: 'Waco marketing fee', amount: wacoAmt, frequency: 'quarterly', startMonth: 'Apr', tds: false });
  }

  // 4. Fix existing month payment rows: mark TDS and WACO as recurring with correct frequency
  Object.entries(DB.months || {}).forEach(([mk, m]) => {
    (m.payments || []).forEach(p => {
      if (/^TDS provisional/i.test(p.name)) { p.recurring = true; p.frequency = 'monthly'; }
      if (/waco marketing fee/i.test(p.name)) { p.recurring = true; p.frequency = 'quarterly'; }
      if (p.recurring && !p.frequency) p.frequency = 'monthly';
    });
  });

  // 5. ── DATA CLEANUP ──────────────────────────────────────────────
  // Remove payments that are in the wrong month (e.g. "Salaries for Aug 26" inside July's data).
  // A recurring payment's name suffix must match the month it lives in.
  // Also remove exact duplicates (same name+amount in same month).
  const MONTH_LABELS_MAP = {};
  Object.keys(DB.months || {}).forEach(mk => {
    MONTH_LABELS_MAP[monthLabel(mk).toLowerCase()] = mk;
  });

  Object.entries(DB.months || {}).forEach(([mk, m]) => {
    if (!m.payments) return;
    const thisMonthLabel = monthLabel(mk).toLowerCase(); // e.g. "aug 2026"

    // Remove recurring entries whose name suffix belongs to a DIFFERENT month
    m.payments = m.payments.filter(p => {
      if (!p.recurring) return true;
      // Check if name ends with "for [month label]" and that label ≠ this month
      const match = p.name.match(/\bfor\s+(.+)$/i);
      if (!match) return true; // no month suffix — keep (e.g. old "TDS provisional")
      const suffix = match[1].trim().toLowerCase();
      // If suffix matches a known month key and it's not this month → wrong month → remove
      if (MONTH_LABELS_MAP[suffix] && MONTH_LABELS_MAP[suffix] !== mk) {
        console.log(`[cleanup] Removing "${p.name}" from ${mk} (belongs to ${MONTH_LABELS_MAP[suffix]})`);
        return false;
      }
      return true;
    });

    // Remove duplicate recurring entries (keep first, remove subsequent same-name)
    const seen = new Set();
    m.payments = m.payments.filter(p => {
      if (!p.recurring) return true;
      const key = p.name.toLowerCase().replace(/\s+for\s+.+$/i, ''); // normalise suffix away
      if (seen.has(key)) {
        console.log(`[cleanup] Removing duplicate "${p.name}" from ${mk}`);
        return false;
      }
      seen.add(key);
      return true;
    });
  });
}

document.addEventListener('DOMContentLoaded', () => {
  try {
    if (window.Cloud && Cloud.cloudConfigured()) {
      if (!window.supabase || typeof window.supabase.createClient !== 'function') {
        showAuthFatalError('The Supabase library failed to load from the CDN. Check your internet connection, then refresh this page. If it keeps happening, your network or an ad-blocker may be blocking cdn.jsdelivr.net.');
        return;
      }
      Cloud.initSupabaseClient();
      wireAuthScreen();
    } else {
      document.getElementById('authGate').style.display = 'none';
      document.getElementById('mainApp').style.display = '';
      startApp();
    }
  } catch (err) {
    console.error('Startup error', err);
    showAuthFatalError('Something went wrong while starting the app: ' + (err && err.message ? err.message : err));
  }
});

function showAuthFatalError(message) {
  const gate = document.getElementById('authGate');
  gate.style.display = 'flex';
  gate.innerHTML = `<div class="auth-card"><div class="auth-error" style="display:block; margin-bottom:0;">${escapeHtml(message)}</div></div>`;
}

/* ===========================================================
   AUTH SCREEN
   =========================================================== */
function describeAuthError(err) {
  if (!err) return 'Something went wrong. Please try again.';
  if (typeof err === 'string') return err;
  if (err.message) return err.message;
  if (err.error_description) return err.error_description;
  if (err.error) return String(err.error);
  try { return JSON.stringify(err); } catch (e) { return 'Something went wrong. Please try again.'; }
}

function wireAuthScreen() {
  const gate = document.getElementById('authGate');
  const mainApp = document.getElementById('mainApp');

  document.getElementById('showFirstRun').onclick = e => {
    e.preventDefault();
    document.getElementById('authSignIn').style.display = 'none';
    document.getElementById('authFirstRun').style.display = 'block';
  };
  document.getElementById('showSignIn').onclick = e => {
    e.preventDefault();
    document.getElementById('authFirstRun').style.display = 'none';
    document.getElementById('authSignIn').style.display = 'block';
  };

  async function completeLogin(user) {
    Cloud.currentUser = { id: user.id, email: user.email, displayName: (user.user_metadata && user.user_metadata.display_name) };
    gate.style.display = 'none';
    mainApp.style.display = '';
    await startApp();
  }

  document.getElementById('authSignInBtn').onclick = async () => {
    const email = document.getElementById('auth-email').value.trim();
    const password = document.getElementById('auth-password').value;
    const errEl = document.getElementById('authError');
    const btn = document.getElementById('authSignInBtn');
    errEl.style.display = 'none';
    if (!email || !password) { errEl.textContent = 'Enter your email and password.'; errEl.style.display = 'block'; return; }
    btn.textContent = 'Signing in…';
    btn.disabled = true;
    try {
      const user = await Cloud.signIn(email, password);
      await completeLogin(user);
    } catch (err) {
      console.error('Sign in failed:', err);
      errEl.textContent = describeAuthError(err);
      errEl.style.display = 'block';
      btn.textContent = 'Sign in';
      btn.disabled = false;
    }
  };

  document.getElementById('authFirstRunBtn').onclick = async () => {
    const email = document.getElementById('fr-email').value.trim();
    const password = document.getElementById('fr-password').value;
    const errEl = document.getElementById('frError');
    errEl.style.display = 'none';
    if (!email || !password) { errEl.textContent = 'Enter an email and password.'; errEl.style.display = 'block'; return; }
    if (password.length < 6) { errEl.textContent = 'Password should be at least 6 characters.'; errEl.style.display = 'block'; return; }
    const btn = document.getElementById('authFirstRunBtn');
    btn.textContent = 'Creating account…';
    btn.disabled = true;
    try {
      const { user, session } = await Cloud.signUpFirstUser(email, password);
      if (!session) {
        errEl.textContent = 'Account created, but email confirmation is still required. In Supabase, go to Authentication → Providers → Email and turn off "Confirm email", then come back and sign in.';
        errEl.style.display = 'block';
        btn.textContent = 'Create account & sign in';
        btn.disabled = false;
        return;
      }
      await completeLogin(user);
    } catch (err) {
      console.error('Sign up failed:', err);
      errEl.textContent = describeAuthError(err);
      errEl.style.display = 'block';
      btn.textContent = 'Create account & sign in';
      btn.disabled = false;
    }
  };

  // Resume an existing session without re-prompting for a password.
  Cloud.client.auth.getSession().then(({ data }) => {
    if (data.session && data.session.user) {
      completeLogin(data.session.user);
    }
  });

  [['auth-password', 'authSignInBtn'], ['fr-password', 'authFirstRunBtn']].forEach(([inputId, btnId]) => {
    document.getElementById(inputId).addEventListener('keydown', e => {
      if (e.key === 'Enter') document.getElementById(btnId).click();
    });
  });
}

function openAddUser() {
  const body = `
    <div class="field"><label>Email</label><input type="email" id="au-email" placeholder="teammate@lmi-india.in"></div>
    <div class="field"><label>Display name</label><input type="text" id="au-name" placeholder="e.g. Priya"></div>
    <div class="field"><label>Temporary password</label><input type="text" id="au-password" placeholder="Set a password to share with them"></div>
    <div class="hint">They'll have full access, same as you. Share this password with them directly (WhatsApp, in person) — they can change it after signing in.</div>
    <div id="au-error" class="auth-error" style="display:none;"></div>`;
  openModal('Add user', body, `<button class="btn" id="au-cancel">Cancel</button><button class="btn btn-primary" id="au-save">Add user</button>`);
  document.getElementById('au-cancel').onclick = closeModal;
  document.getElementById('au-save').onclick = async () => {
    const email = document.getElementById('au-email').value.trim();
    const name = document.getElementById('au-name').value.trim();
    const password = document.getElementById('au-password').value;
    const errEl = document.getElementById('au-error');
    errEl.style.display = 'none';
    if (!email || !password) { errEl.textContent = 'Email and password are required.'; errEl.style.display = 'block'; return; }
    if (password.length < 6) { errEl.textContent = 'Password should be at least 6 characters.'; errEl.style.display = 'block'; return; }
    try {
      const { hadSession } = await Cloud.addTeammate(email, password, name);
      if (!hadSession) {
        toast(`${name || email} created — but they may need to confirm their email before signing in`);
      } else {
        toast(`${name || email} added — share their password with them`);
      }
      closeModal();
    } catch (err) {
      console.error('Add user failed:', err);
      errEl.textContent = describeAuthError(err);
      errEl.style.display = 'block';
    }
  };
}

/* ===========================================================
   SCENARIO PLANNER
   A what-if cashflow modelling workspace. Reads live data
   but writes nothing back. Scenarios saved to DB.scenarios.
   =========================================================== */

// Active scenario state — reset each time planner is opened or loaded
let SP = null;

function spEmptyScenario() {
  return {
    name: '',
    spendLines: [],       // [{id, label, amount}]
    checkedRecvIds: [],   // receivable ids selected for inflow
    unlistedInflows: [],  // [{id, label, amount}]
  };
}

function openScenarioPlanner() {
  if (!SP) SP = spEmptyScenario();
  if (!DB.scenarios) DB.scenarios = [];
  document.getElementById('scenarioOverlay').style.display = 'block';
  document.body.style.overflow = 'hidden';
  spPopulateSavedList();
  spRender();
  spWireButtons();
}

function closeScenarioPlanner() {
  document.getElementById('scenarioOverlay').style.display = 'none';
  document.body.style.overflow = '';
}

function spWireButtons() {
  document.getElementById('sp-close-btn').onclick = () => {
    if (confirm('Close Scenario Planner? Unsaved changes will be lost.')) closeScenarioPlanner();
  };
  document.getElementById('sp-add-spend').onclick = spAddSpendLine;
  document.getElementById('sp-add-unlisted').onclick = spAddUnlisted;
  document.getElementById('sp-save-btn').onclick = spSaveScenario;
  document.getElementById('sp-load-btn').onclick = spLoadSelected;
  document.getElementById('sp-delete-btn').onclick = spDeleteSelected;
}

/* ---------- Spend lines ---------- */
function spAddSpendLine(label = '', amount = 0) {
  const id = uid();
  SP.spendLines.push({ id, label, amount });
  spRenderSpend();
  spRenderReport();
  // Focus the new label input
  setTimeout(() => {
    const input = document.querySelector(`[data-sp-spend-label="${id}"]`);
    if (input) input.focus();
  }, 50);
}

function spRenderSpend() {
  const wrap = document.getElementById('sp-spend-lines');
  if (!SP.spendLines.length) {
    wrap.innerHTML = '<div class="empty-note">No spend lines yet. Click + Add line.</div>';
    document.getElementById('sp-spend-total').textContent = '₹0';
    return;
  }
  wrap.innerHTML = SP.spendLines.map(s => `
    <div style="display:grid; grid-template-columns:1fr auto auto; gap:6px; align-items:center; margin-bottom:8px;">
      <input type="text" data-sp-spend-label="${s.id}" value="${escapeHtml(s.label)}"
        placeholder="e.g. Import order, Marketing"
        style="padding:7px 9px; border:1px solid var(--line); border-radius:4px; font-size:13px; font-family:var(--sans);">
      <input type="number" data-sp-spend-amt="${s.id}" value="${s.amount || ''}"
        placeholder="0"
        style="width:110px; padding:7px 9px; border:1px solid var(--line); border-radius:4px; font-size:13px; font-family:var(--mono); text-align:right;">
      <button data-sp-del-spend="${s.id}" style="border:none; background:none; color:#aab2bd; cursor:pointer; font-size:14px; padding:4px;">✕</button>
    </div>`).join('');

  // Wire inputs
  wrap.querySelectorAll('[data-sp-spend-label]').forEach(inp => {
    inp.oninput = () => {
      const s = SP.spendLines.find(x => x.id === inp.dataset.spSpendLabel);
      if (s) { s.label = inp.value; spRenderReport(); }
    };
  });
  wrap.querySelectorAll('[data-sp-spend-amt]').forEach(inp => {
    inp.oninput = () => {
      const s = SP.spendLines.find(x => x.id === inp.dataset.spSpendAmt);
      if (s) { s.amount = parseFloat(inp.value) || 0; spRenderReport(); spUpdateSpendTotal(); }
    };
  });
  wrap.querySelectorAll('[data-sp-del-spend]').forEach(btn => {
    btn.onclick = () => {
      SP.spendLines = SP.spendLines.filter(x => x.id !== btn.dataset.spDelSpend);
      spRenderSpend(); spRenderReport();
    };
  });
  spUpdateSpendTotal();
}

function spUpdateSpendTotal() {
  const total = SP.spendLines.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  document.getElementById('sp-spend-total').textContent = fmtMoney(total);
}

/* ---------- Receivables checklist ---------- */
function spRenderRecvChecks() {
  const todayMk = todayMonthKey();
  const m = DB.months[todayMk];
  const wrap = document.getElementById('sp-recv-checks');
  const ownRecvs = m ? m.receivables.filter(r => !r._carriedFrom && (Number(r.amount) || 0) > 0) : [];

  if (!ownRecvs.length) {
    wrap.innerHTML = '<div class="empty-note">No receivables in current month.</div>';
    return;
  }
  wrap.innerHTML = ownRecvs.map(r => {
    const checked = SP.checkedRecvIds.includes(r.id);
    return `<div style="display:flex; align-items:center; gap:8px; margin-bottom:7px; font-size:13px;">
      <input type="checkbox" id="sp-recv-${r.id}" ${checked ? 'checked' : ''} style="width:auto; cursor:pointer;">
      <label for="sp-recv-${r.id}" style="flex:1; cursor:pointer; display:flex; justify-content:space-between;">
        <span>${escapeHtml(r.name)}</span>
        <span style="font-family:var(--mono); color:var(--ink-soft);">${fmtMoney(r.amount)}</span>
      </label>
    </div>`;
  }).join('');

  wrap.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.onchange = () => {
      const recvId = cb.id.replace('sp-recv-', '');
      if (cb.checked) {
        if (!SP.checkedRecvIds.includes(recvId)) SP.checkedRecvIds.push(recvId);
      } else {
        SP.checkedRecvIds = SP.checkedRecvIds.filter(x => x !== recvId);
      }
      spUpdateInflowTotal();
      spRenderReport();
    };
  });
}

/* ---------- Unlisted inflows ---------- */
function spAddUnlisted(label = '', amount = 0) {
  const id = uid();
  SP.unlistedInflows.push({ id, label, amount });
  spRenderUnlisted();
  spRenderReport();
  setTimeout(() => {
    const input = document.querySelector(`[data-sp-ul-label="${id}"]`);
    if (input) input.focus();
  }, 50);
}

function spRenderUnlisted() {
  const wrap = document.getElementById('sp-unlisted-lines');
  if (!SP.unlistedInflows.length) {
    wrap.innerHTML = '<div class="empty-note">None added.</div>';
    spUpdateInflowTotal();
    return;
  }
  wrap.innerHTML = SP.unlistedInflows.map(u => `
    <div style="display:grid; grid-template-columns:1fr auto auto; gap:6px; align-items:center; margin-bottom:8px;">
      <input type="text" data-sp-ul-label="${u.id}" value="${escapeHtml(u.label)}"
        placeholder="e.g. Expected bank transfer"
        style="padding:7px 9px; border:1px solid var(--line); border-radius:4px; font-size:13px; font-family:var(--sans);">
      <input type="number" data-sp-ul-amt="${u.id}" value="${u.amount || ''}"
        placeholder="0"
        style="width:110px; padding:7px 9px; border:1px solid var(--line); border-radius:4px; font-size:13px; font-family:var(--mono); text-align:right;">
      <button data-sp-del-ul="${u.id}" style="border:none; background:none; color:#aab2bd; cursor:pointer; font-size:14px; padding:4px;">✕</button>
    </div>`).join('');

  wrap.querySelectorAll('[data-sp-ul-label]').forEach(inp => {
    inp.oninput = () => {
      const u = SP.unlistedInflows.find(x => x.id === inp.dataset.spUlLabel);
      if (u) { u.label = inp.value; spRenderReport(); }
    };
  });
  wrap.querySelectorAll('[data-sp-ul-amt]').forEach(inp => {
    inp.oninput = () => {
      const u = SP.unlistedInflows.find(x => x.id === inp.dataset.spUlAmt);
      if (u) { u.amount = parseFloat(inp.value) || 0; spUpdateInflowTotal(); spRenderReport(); }
    };
  });
  wrap.querySelectorAll('[data-sp-del-ul]').forEach(btn => {
    btn.onclick = () => {
      SP.unlistedInflows = SP.unlistedInflows.filter(x => x.id !== btn.dataset.spDelUl);
      spRenderUnlisted(); spRenderReport();
    };
  });
  spUpdateInflowTotal();
}

function spInflowTotal() {
  const todayMk = todayMonthKey();
  const m = DB.months[todayMk];
  const recvTotal = m
    ? m.receivables
        .filter(r => !r._carriedFrom && SP.checkedRecvIds.includes(r.id))
        .reduce((s, r) => s + (Number(r.amount) || 0), 0)
    : 0;
  const unlistedTotal = SP.unlistedInflows.reduce((s, u) => s + (Number(u.amount) || 0), 0);
  return { recvTotal, unlistedTotal, total: recvTotal + unlistedTotal };
}

function spUpdateInflowTotal() {
  document.getElementById('sp-inflow-total').textContent = fmtMoney(spInflowTotal().total);
}

/* ---------- Report ---------- */
function spRenderReport() {
  const todayMk = todayMonthKey();
  const next1Mk = nextMonthKey(todayMk);
  const next2Mk = nextMonthKey(next1Mk);

  const totalSpend = SP.spendLines.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const { recvTotal, unlistedTotal, total: inflowTotal } = spInflowTotal();

  // Base projections from live data
  const todayClosing = computeClosing(todayMk);
  const next1Closing = computeClosing(next1Mk);
  const next2Closing = computeClosing(next2Mk);

  // Scenario projections
  const todayScenario = todayClosing + inflowTotal - totalSpend;
  // Next months cascade from the scenario closing of the prior month
  const next1Base = getOpening(next1Mk);
  const { receiptsTotal: r1, paymentsTotal: p1 } = monthTotals(next1Mk);
  const next1Scenario = todayScenario + (r1 - p1); // uses scenario closing as next month's opening
  const next2Base = getOpening(next2Mk);
  const { receiptsTotal: r2, paymentsTotal: p2 } = monthTotals(next2Mk);
  const next2Scenario = next1Scenario + (r2 - p2);

  function row(label, value, bold = false, highlight = null) {
    const color = highlight === 'good' ? 'var(--green)' : highlight === 'bad' ? 'var(--red)' : 'var(--ink)';
    const fwt = bold ? '700' : '400';
    return `<div style="display:flex; justify-content:space-between; padding:7px 0; border-bottom:1px solid #f0f1f3; font-size:13px; font-weight:${fwt}; color:${color};">
      <span>${label}</span><span style="font-family:var(--mono);">${fmtMoney(value)}</span>
    </div>`;
  }
  function sectionHead(label) {
    return `<div style="font-size:10.5px; text-transform:uppercase; letter-spacing:.07em; color:var(--ink-soft); margin:14px 0 6px 0; font-weight:600;">${label}</div>`;
  }
  function divider() {
    return `<div style="border-top:2px solid var(--line); margin:10px 0;"></div>`;
  }

  const spendBreakdown = SP.spendLines.length
    ? SP.spendLines.map(s => row(`&nbsp;&nbsp;&nbsp;${escapeHtml(s.label) || 'Unnamed spend'}`, -(Number(s.amount) || 0))).join('')
    : '';

  const inflowBreakdown = SP.checkedRecvIds.length || SP.unlistedInflows.length
    ? (SP.checkedRecvIds.length ? row('&nbsp;&nbsp;&nbsp;Selected receivables', recvTotal) : '')
      + (SP.unlistedInflows.length ? row('&nbsp;&nbsp;&nbsp;Unlisted inflows', unlistedTotal) : '')
    : '';

  document.getElementById('sp-report-body').innerHTML = `
    ${sectionHead(monthLabel(todayMk) + ' — current month')}
    ${row('Live closing projection', todayClosing)}
    ${inflowBreakdown}
    ${spendBreakdown}
    ${divider()}
    ${row('Scenario closing', todayScenario, true, todayScenario >= 0 ? 'good' : 'bad')}

    ${sectionHead(monthLabel(next1Mk) + ' — next month')}
    ${row('Live closing projection', next1Closing)}
    ${row('After scenario carry-forward', next1Scenario, true, next1Scenario >= 0 ? 'good' : 'bad')}

    ${sectionHead(monthLabel(next2Mk) + ' — month after')}
    ${row('Live closing projection', next2Closing)}
    ${row('After scenario carry-forward', next2Scenario, true, next2Scenario >= 0 ? 'good' : 'bad')}
  `;
}

/* ---------- Full render ---------- */
function spRender() {
  spRenderSpend();
  spRenderRecvChecks();
  spRenderUnlisted();
  spRenderReport();
}

/* ---------- Save / Load / Delete ---------- */
function spSaveScenario() {
  const name = prompt('Save scenario as:', SP.name || '');
  if (!name || !name.trim()) return;
  if (!DB.scenarios) DB.scenarios = [];
  const existing = DB.scenarios.find(s => s.name === name.trim());
  if (existing) {
    if (!confirm(`Overwrite existing scenario "${name.trim()}"?`)) return;
    Object.assign(existing, { ...SP, name: name.trim(), savedAt: new Date().toISOString() });
  } else {
    DB.scenarios.push({ ...SP, name: name.trim(), savedAt: new Date().toISOString() });
  }
  SP.name = name.trim();
  saveDB();
  spPopulateSavedList();
  toast(`Scenario "${name.trim()}" saved`);
}

function spPopulateSavedList() {
  const sel = document.getElementById('sp-saved-list');
  const current = sel.value;
  sel.innerHTML = '<option value="">— Saved scenarios —</option>' +
    (DB.scenarios || []).map(s =>
      `<option value="${escapeHtml(s.name)}" ${s.name === current ? 'selected' : ''}>${escapeHtml(s.name)} (${new Date(s.savedAt).toLocaleDateString('en-IN')})</option>`
    ).join('');
}

function spLoadSelected() {
  const name = document.getElementById('sp-saved-list').value;
  if (!name) { toast('Select a scenario to load'); return; }
  const found = (DB.scenarios || []).find(s => s.name === name);
  if (!found) return;
  SP = { ...found };
  spRender();
  toast(`Loaded: ${name}`);
}

function spDeleteSelected() {
  const name = document.getElementById('sp-saved-list').value;
  if (!name) { toast('Select a scenario to delete'); return; }
  if (!confirm(`Delete scenario "${name}"?`)) return;
  DB.scenarios = (DB.scenarios || []).filter(s => s.name !== name);
  saveDB();
  spPopulateSavedList();
  SP = spEmptyScenario();
  spRender();
  toast(`Deleted: ${name}`);
}

/* ===========================================================
   PENDING ACTIONS
   Shared task board with sections per team member + Completed.
   Stored in DB.pendingActions (workspace-level, synced to cloud).
   =========================================================== */

const PA_SECTIONS = ['NIRALI', 'ASHOK', 'SANDEEP'];

function paInit() {
  if (!DB.pendingActions) {
    DB.pendingActions = { NIRALI: [], ASHOK: [], SANDEEP: [], COMPLETED: [] };
  }
  // Ensure all sections exist (in case new ones added later)
  [...PA_SECTIONS, 'COMPLETED'].forEach(s => {
    if (!DB.pendingActions[s]) DB.pendingActions[s] = [];
  });
}

function openPendingActions() {
  paInit();
  document.getElementById('pendingActionsOverlay').style.display = 'block';
  document.body.style.overflow = 'hidden';
  document.getElementById('pa-close-btn').onclick = closePendingActions;
  renderMonthTabs(); // re-render so Pending Actions tab shows as active
  paRender();
}

function closePendingActions() {
  document.getElementById('pendingActionsOverlay').style.display = 'none';
  document.body.style.overflow = '';
  renderMonthTabs(); // re-render so current month tab shows as active again
}

function paRender() {
  paInit();
  const wrap = document.getElementById('pa-columns');
  const sectionColors = {
    NIRALI:    { hdr: '#1e4f8a', bg: '#e8f0fb' },
    ASHOK:     { hdr: '#1f7a4d', bg: '#e6f4ec' },
    SANDEEP:   { hdr: '#9a6b14', bg: '#fbf0dd' },
    COMPLETED: { hdr: '#5b6470', bg: '#f6f4ee' },
  };

  const allSections = [...PA_SECTIONS, 'COMPLETED'];
  wrap.innerHTML = allSections.map(section => {
    const items = DB.pendingActions[section] || [];
    const col = sectionColors[section];
    const isCompleted = section === 'COMPLETED';

    const itemsHtml = items.length ? items.map(item => `
      <div style="display:flex; align-items:flex-start; gap:8px; padding:9px 0; border-bottom:1px solid #eef0f2;">
        ${!isCompleted ? `<input type="checkbox" data-pa-done="${item.id}" data-pa-section="${section}" style="margin-top:3px; cursor:pointer; flex-shrink:0;">` : `<span style="color:var(--green); font-size:14px; flex-shrink:0;">&#10003;</span>`}
        <div style="flex:1; min-width:0;">
          <div style="font-size:13px; color:var(--ink); word-break:break-word;">${escapeHtml(item.text)}</div>
          <div style="font-size:10.5px; color:var(--ink-soft); margin-top:3px;">
            ${isCompleted ? `Done by <strong>${escapeHtml(item.doneBy || '')}</strong> &middot; ${item.doneAt ? new Date(item.doneAt).toLocaleDateString('en-IN', {day:'numeric',month:'short',year:'2-digit'}) : ''}` : `Added ${item.addedAt ? new Date(item.addedAt).toLocaleDateString('en-IN', {day:'numeric',month:'short'}) : ''}`}
          </div>
        </div>
        <button data-pa-delete="${item.id}" data-pa-section="${section}" title="Delete permanently" style="border:none; background:none; color:#ccc; cursor:pointer; font-size:13px; padding:2px 4px; flex-shrink:0;">&#10005;</button>
      </div>`).join('')
    : `<div style="color:var(--ink-soft); font-size:12.5px; font-style:italic; padding:10px 0;">${isCompleted ? 'No completed items yet.' : 'No pending items.'}</div>`;

    const addHtml = !isCompleted ? `
      <div style="display:flex; gap:6px; margin-top:12px;">
        <input type="text" id="pa-new-${section}" placeholder="Add item for ${section}..."
          style="flex:1; padding:8px 10px; border:1px solid var(--line); border-radius:5px; font-size:13px; font-family:var(--sans);">
        <button data-pa-add="${section}" class="btn btn-sm btn-primary">Add</button>
      </div>` : '';

    return `
      <div class="panel">
        <div class="panel-head" style="background:${col.bg};">
          <h2 style="color:${col.hdr};">${section}</h2>
          <span style="font-size:11px; color:${col.hdr}; font-weight:600;">${items.length} item${items.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="panel-body">
          <div id="pa-items-${section}">${itemsHtml}</div>
          ${addHtml}
        </div>
      </div>`;
  }).join('');

  // Wire add buttons
  wrap.querySelectorAll('[data-pa-add]').forEach(btn => {
    const section = btn.dataset.paAdd;
    const input = document.getElementById(`pa-new-${section}`);
    const doAdd = () => {
      const text = input.value.trim();
      if (!text) return;
      paInit();
      DB.pendingActions[section].push({
        id: uid(),
        text,
        addedAt: new Date().toISOString(),
        addedBy: Cloud && Cloud.currentUser ? (Cloud.currentUser.displayName || Cloud.currentUser.email || '') : 'User',
      });
      input.value = '';
      saveDB(); paRender();
    };
    btn.onclick = doAdd;
    input.onkeydown = e => { if (e.key === 'Enter') doAdd(); };
  });

  // Wire done (checkbox) buttons
  wrap.querySelectorAll('[data-pa-done]').forEach(cb => {
    cb.onchange = () => {
      if (!cb.checked) return;
      const id = cb.dataset.paDone;
      const section = cb.dataset.paSection;
      paInit();
      const idx = DB.pendingActions[section].findIndex(x => x.id === id);
      if (idx === -1) return;
      const [item] = DB.pendingActions[section].splice(idx, 1);
      item.doneAt = new Date().toISOString();
      item.doneBy = Cloud && Cloud.currentUser ? (Cloud.currentUser.displayName || Cloud.currentUser.email || '').split('@')[0] : 'User';
      DB.pendingActions.COMPLETED.unshift(item); // newest at top
      saveDB(); paRender();
      toast(`Moved to Completed`);
    };
  });

  // Wire delete buttons
  wrap.querySelectorAll('[data-pa-delete]').forEach(btn => {
    btn.onclick = () => {
      if (!confirm('Permanently delete this item?')) return;
      const id = btn.dataset.paDelete;
      const section = btn.dataset.paSection;
      paInit();
      DB.pendingActions[section] = DB.pendingActions[section].filter(x => x.id !== id);
      saveDB(); paRender();
    };
  });
}

// Hook into realtime: when workspace changes, re-render if Pending Actions is open
const _origHandleWorkspaceChange = typeof handleWorkspaceChange !== 'undefined' ? handleWorkspaceChange : null;

/* ===========================================================
   INVOICING MODULE  v2.3.0
   PI/TI generation, client master, Excel export, cashflow
   integration. Data stored in DB.invoicing (workspace-level).
   =========================================================== */

// ── Constants ──────────────────────────────────────────────
const GORU = {
  name: 'GORU TRAINING PRIVATE LIMITED',
  addr1: 'No 1108, Floor No 11th, Sureshwari\nTechno IT Park Premises,',
  addr2: 'Link Road, Near Eskay Resorts, Borivali\n400092 Mumbai',
  state: 'Maharashtra',
  gstin: '27AAGCG2703D1Z2',
  pan: 'AAGCG2703D',
  tan: 'MUMG18684B',
  bank: 'HDFC Bank',
  branch: 'Churchgate, Mumbai 400020',
  acno: '50200048157133',
  ifsc: 'HDFC0000501',
};

// ── Data initialisation ────────────────────────────────────
function invInit() {
  if (!DB.invoicing) {
    DB.invoicing = {
      clients: [],
      piSequence: {},   // { '26-27': 0 }
      tiSequence: {},
      proformas: [],    // PI records
      taxInvoices: [],  // TI records
    };
  }
  if (!DB.invoicing.clients) DB.invoicing.clients = [];
  if (!DB.invoicing.products) DB.invoicing.products = [];
  if (!DB.invoicing.piSequence) DB.invoicing.piSequence = {};
  if (!DB.invoicing.tiSequence) DB.invoicing.tiSequence = {};
  if (!DB.invoicing.proformas) DB.invoicing.proformas = [];
  if (!DB.invoicing.taxInvoices) DB.invoicing.taxInvoices = [];
}

function currentInvFY() {
  return DB.currentFY || '26-27';
}

function nextPINumber() {
  const fy = currentInvFY();
  const n = (DB.invoicing.piSequence[fy] || 0) + 1;
  DB.invoicing.piSequence[fy] = n;
  return `PI/${fy}/${String(n).padStart(3, '0')}`;
}

function nextTINumber() {
  const fy = currentInvFY();
  const n = (DB.invoicing.tiSequence[fy] || 0) + 1;
  DB.invoicing.tiSequence[fy] = n;
  return `TI/${fy}/${String(n).padStart(3, '0')}`;
}

function gstType(clientState) {
  if (!clientState) return 'igst';
  return clientState.trim().toLowerCase() === 'maharashtra' ? 'intra' : 'igst';
}

// ── Open / Close ───────────────────────────────────────────
let INV_TAB = 'pi'; // 'pi' or 'ti'

function openInvoicingModule() {
  invInit();
  document.getElementById('invoicingOverlay').style.display = 'block';
  document.body.style.overflow = 'hidden';
  document.getElementById('inv-fy-label').textContent =
    `Goru Training Pvt. Ltd. · FY ${currentInvFY()}`;
  invWireButtons();
  invRenderRegister();
}

function closeInvoicingModule() {
  document.getElementById('invoicingOverlay').style.display = 'none';
  document.body.style.overflow = '';
  renderMonthTabs();
}

function invWireButtons() {
  document.getElementById('inv-close-btn').onclick = closeInvoicingModule;
  document.getElementById('inv-btn-clients').onclick = invOpenClients;
  document.getElementById('inv-btn-products').onclick = invOpenAddProduct;
  document.getElementById('inv-btn-product-list').onclick = invOpenProductList;
  document.getElementById('inv-btn-startnums').onclick = invOpenStartNumbers;
  document.getElementById('inv-btn-new-pi').onclick = () => invOpenPIForm();
  document.getElementById('inv-btn-new-ti').onclick = () => invOpenTIForm();
  document.getElementById('inv-btn-report').onclick = invOpenReport;
  document.querySelectorAll('.inv-reg-tab').forEach(tab => {
    tab.onclick = () => {
      INV_TAB = tab.dataset.tab;
      document.querySelectorAll('.inv-reg-tab').forEach(t => {
        t.style.borderBottom = t === tab ? '2px solid var(--navy)' : 'none';
        t.style.color = t === tab ? 'var(--navy)' : 'var(--ink-soft)';
      });
      invRenderRegister();
    };
  });
}

// ── Register table ─────────────────────────────────────────
function invRenderRegister() {
  invInit();
  const isPi = INV_TAB === 'pi';
  const list = isPi ? DB.invoicing.proformas : DB.invoicing.taxInvoices;
  const fy = currentInvFY();
  const fyList = list.filter(inv => inv.invNo.includes(fy)).reverse();

  const head = document.getElementById('inv-register-head');
  const body = document.getElementById('inv-register-body');

  head.innerHTML = `<tr>
    <th>Invoice No</th><th>Date</th><th>Client</th>
    <th>Description</th><th style="text-align:right;">Gross (₹)</th>
    <th>Status</th><th style="text-align:center;">Actions</th>
  </tr>`;

  if (!fyList.length) {
    body.innerHTML = `<tr><td colspan="7" class="empty-note" style="padding:20px; text-align:center;">No ${isPi ? 'Proforma' : 'Tax'} invoices for FY ${fy} yet.</td></tr>`;
    return;
  }

  const statusColor = { draft: '#9a6b14', sent: '#1e4f8a', paid: '#1f7a4d', cancelled: '#a4302a', converted: '#5b6470' };

  body.innerHTML = fyList.map(inv => `
    <tr style="font-size:13px;">
      <td style="font-family:var(--mono); font-weight:600;">${escapeHtml(inv.invNo)}</td>
      <td>${inv.date || ''}</td>
      <td>${escapeHtml(inv.clientName || '')}</td>
      <td style="max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${escapeHtml(inv.desc || '')}</td>
      <td style="font-family:var(--mono); text-align:right;">${fmtMoney(inv.gross || 0)}</td>
      <td><span class="badge" style="background:${statusColor[inv.status] || '#ddd'}22; color:${statusColor[inv.status] || '#666'};">${(inv.status || 'draft').toUpperCase()}</span></td>
      <td style="text-align:center;">
        <span class="row-actions" style="justify-content:center;">
          ${inv.status !== 'cancelled' && inv.status !== 'converted' ? `<button data-inv-edit="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Edit">&#9998;</button>` : ''}
          ${isPi && inv.status !== 'cancelled' && inv.status !== 'converted' ? `<button data-inv-to-ti="${inv.id}" title="Convert to TI">&#8594;TI</button>` : ''}
          <button data-inv-word="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Download Word">&#128229;</button>
          <button data-inv-email="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Email">&#9993;</button>
          <button data-inv-del="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="${inv.status === 'cancelled' ? 'Delete permanently' : 'Cancel/Delete'}" style="color:var(--red);">&#10005;</button>
        </span>
      </td>
    </tr>`).join('');

  body.querySelectorAll('[data-inv-edit]').forEach(b =>
    b.onclick = () => b.dataset.invType === 'pi'
      ? invOpenPIForm(b.dataset.invEdit)
      : invOpenTIForm(null, b.dataset.invEdit));
  body.querySelectorAll('[data-inv-to-ti]').forEach(b =>
    b.onclick = () => invOpenTIForm(b.dataset.invToTi));
  body.querySelectorAll('[data-inv-word]').forEach(b =>
    b.onclick = () => invDownloadWord(b.dataset.invWord, b.dataset.invType));
  body.querySelectorAll('[data-inv-email]').forEach(b =>
    b.onclick = () => invEmail(b.dataset.invEmail, b.dataset.invType));
  body.querySelectorAll('[data-inv-del]').forEach(b =>
    b.onclick = () => invCancel(b.dataset.invDel, b.dataset.invType));
}

// ── Client management ──────────────────────────────────────
function invOpenClients() {
  invInit();
  const clients = DB.invoicing.clients;
  const body = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
      <span style="font-size:12.5px; color:var(--ink-soft);">${clients.length} client${clients.length !== 1 ? 's' : ''} on file</span>
      <button class="btn btn-sm btn-primary" id="cl-add">+ Add client</button>
    </div>
    <div style="max-height:420px; overflow-y:auto; border:1px solid var(--line); border-radius:6px;">
      ${clients.length ? clients.map(c => `
        <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 14px; border-bottom:1px solid #eef0f2; font-size:13px;">
          <div>
            <div style="font-weight:600;">${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</div>
            <div style="font-size:11.5px; color:var(--ink-soft);">${escapeHtml(c.state || '')} · ${escapeHtml(c.gstin || 'No GSTIN')}</div>
          </div>
          <span class="row-actions">
            <button data-cl-edit="${c.id}" title="Edit">&#9998;</button>
            <button data-cl-del="${c.id}" title="Delete" style="color:var(--red);">&#10005;</button>
          </span>
        </div>`).join('') : '<div class="empty-note" style="padding:20px;">No clients yet. Click + Add client to import from the bundled list or add manually.</div>'}
    </div>`;

  openModal('Client master', body, `
    <button class="btn" id="cl-import">&#8659; Import bundled list</button>
    <button class="btn" id="cl-close">Close</button>`);

  document.getElementById('cl-close').onclick = closeModal;
  document.getElementById('cl-add').onclick = () => invOpenClientForm();
  document.getElementById('cl-import').onclick = invImportBundledClients;
  document.querySelectorAll('[data-cl-edit]').forEach(b =>
    b.onclick = () => invOpenClientForm(b.dataset.clEdit));
  document.querySelectorAll('[data-cl-del]').forEach(b => {
    b.onclick = () => {
      if (!confirm('Delete this client?')) return;
      DB.invoicing.clients = DB.invoicing.clients.filter(c => c.id !== b.dataset.clDel);
      saveDB(); invOpenClients();
    };
  });
}

function invOpenClientForm(clientId) {
  invInit();
  const existing = clientId ? DB.invoicing.clients.find(c => c.id === clientId) : null;
  const v = existing || {};
  const body = `
    <div class="field-row">
      <div class="field"><label>Short name</label><input id="cf-short" value="${escapeHtml(v.shortName||'')}"></div>
      <div class="field"><label>State</label><input id="cf-state" value="${escapeHtml(v.state||'')}"></div>
    </div>
    <div class="field"><label>Company name</label><input id="cf-company" value="${escapeHtml(v.companyName||'')}"></div>
    <div class="field"><label>Address line 1</label><input id="cf-addr1" value="${escapeHtml(v.addr1||'')}"></div>
    <div class="field"><label>Address line 2</label><input id="cf-addr2" value="${escapeHtml(v.addr2||'')}"></div>
    <div class="field"><label>Address line 3</label><input id="cf-addr3" value="${escapeHtml(v.addr3||'')}"></div>
    <div class="field-row">
      <div class="field"><label>GSTIN</label><input id="cf-gstin" value="${escapeHtml(v.gstin||'')}"></div>
      <div class="field"><label>PAN</label><input id="cf-pan" value="${escapeHtml(v.pan||'')}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>TAN</label><input id="cf-tan" value="${escapeHtml(v.tan||'')}"></div>
      <div class="field"><label>Kind Attn</label><input id="cf-attn" value="${escapeHtml(v.attn||'')}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Mobile</label><input id="cf-mobile" value="${escapeHtml(v.mobile||'')}"></div>
      <div class="field"><label>Email</label><input id="cf-email" value="${escapeHtml(v.email||'')}"></div>
    </div>`;
  openModal(existing ? 'Edit client' : 'Add client', body,
    `<button class="btn" id="cf-cancel">Cancel</button><button class="btn btn-primary" id="cf-save">Save</button>`);
  document.getElementById('cf-cancel').onclick = () => { closeModal(); invOpenClients(); };
  document.getElementById('cf-save').onclick = () => {
    const get = id => document.getElementById(id).value.trim();
    const data = {
      id: existing ? existing.id : uid(),
      shortName: get('cf-short'), companyName: get('cf-company'),
      addr1: get('cf-addr1'), addr2: get('cf-addr2'), addr3: get('cf-addr3'),
      state: get('cf-state'), gstin: get('cf-gstin'), pan: get('cf-pan'),
      tan: get('cf-tan'), attn: get('cf-attn'), mobile: get('cf-mobile'), email: get('cf-email'),
    };
    if (!data.shortName) { toast('Short name is required'); return; }
    if (existing) {
      Object.assign(existing, data);
    } else {
      DB.invoicing.clients.push(data);
    }
    saveDB(); closeModal(); invOpenClients();
    toast(`Client ${data.shortName} saved`);
  };
}

// Bundled client list from the uploaded licensee details
const BUNDLED_CLIENTS = [
  {shortName:'AMEEN',companyName:'THE BOTTOM LINE LEADERSHIP CONSULTING LLP',addr1:'84-G, DASTOOR BLOCKS, NAIGAUM CROSS ROAD,',addr2:'DADAR, MUMBAI 400014',addr3:'',state:'Maharashtra',gstin:'27AAIFT9075C1ZW',pan:'AAIFT9075C',tan:'',attn:'Ameen Merchant',mobile:'9820987775',email:'ameenmerchant7@gmail.com'},
  {shortName:'ANIL',companyName:'ANIL JHINGAN',addr1:'116, Malviya Nagar,',addr2:'Bhopal, Madhya Pradesh - 462003',addr3:'',state:'Madhya Pradesh',gstin:'23ACUPJ5185F2ZB',pan:'ACUPJ5185F',tan:'',attn:'Anil Jhingan',mobile:'9826255067',email:'jhingan.anil@gmail.com'},
  {shortName:'ANUPAMA',companyName:'ANUPAMA SRIVASTAVA',addr1:'C002, Salarpuria Sattva Gold Summit,',addr2:'Doddagubbi, Kothanur',addr3:'Bengaluru - 560077',state:'Karnataka',gstin:'',pan:'ASSPS8483R',tan:'',attn:'Anupama Srivastava',mobile:'9845954450',email:'srivastavaanupamar@gmail.com'},
  {shortName:'ARYAN',companyName:'GRIT UNLIMITED',addr1:'103, Adarsh Society, Kalinga',addr2:'Athwa lines, Surat - 395001',addr3:'',state:'Gujarat',gstin:'24DUSPB4176H1ZF',pan:'DUSPB4176H',tan:'',attn:'Aryan Naidu',mobile:'9824102318',email:'baryannaidu@gmail.com'},
  {shortName:'ANAND',companyName:'LEADERS LEAGUE',addr1:'103, Adarsh Society,',addr2:'Athwa lines, Surat - 395002',addr3:'',state:'Gujarat',gstin:'24ABBPN3317M1ZQ',pan:'ABBPN3317M',tan:'SRTA02534A',attn:'B. Anand Naidu',mobile:'9824133238',email:'b.anandnaidu@gmail.com'},
  {shortName:'BHAGYASHRI',companyName:'BHAGYASHRI NEELESH VARTAK',addr1:'B-1502, Adney 2, Holy cross Road, Near SBI Bank,',addr2:'I.C.Colony, Borivali West, Mumbai 400103',addr3:'',state:'Maharashtra',gstin:'',pan:'AMAPR9005Q',tan:'',attn:'Bhagyashri Vartak',mobile:'9967662274',email:'bhagyashrinv@gmail.com'},
  {shortName:'BHAVIN',companyName:'M/S IRA ENTERPRISE',addr1:'100, MAHASUKHNAGAR SOCIETY, NR. NOBEL SCHOOL,',addr2:'ON PARSHWANATH HARIVILLA ROAD, KRISHNANAGAR,',addr3:'PO: SAIJPUR BOGHA, AHMEDABAD - 382345',state:'Gujarat',gstin:'24HGUPS5996A1ZL',pan:'HGUPS5996A',tan:'',attn:'Bhavin Soni',mobile:'9824650620',email:'bnsoni1975@yahoo.com'},
  {shortName:'MAHESH',companyName:'BRIDGE PEOPLE TECHNOLOGY SOLUTIONS PVT LTD',addr1:'Unit 3F04, Tower F, Century Central,',addr2:'Konanakunte Cross, Kanakapura Main Road,',addr3:'Bangalore - 560062',state:'Karnataka',gstin:'29AAFCB0486R1ZZ',pan:'AAFCB0486R',tan:'',attn:'C. N. Mahesh',mobile:'8431318575',email:'mahesh@grorgconsulting.in'},
  {shortName:'DEBASIS',companyName:'CBE LEARNING PVT LTD.',addr1:'320, Block 1B, 73 East Avenue, Genda Circle,',addr2:'Bhailal Amin Marg, Vadodara - 390017',addr3:'',state:'Gujarat',gstin:'24AAICC5630M1ZJ',pan:'AAICC5630M',tan:'',attn:'Debasis Majumdar',mobile:'9687044466',email:'debasis@sixsigmaconcept.com'},
  {shortName:'DEEPAK',companyName:'CEE EM EXPORTS PVT LTD',addr1:'512 Deepshikha Building,',addr2:'8 Rajendra Place,',addr3:'New Delhi-110008',state:'Delhi',gstin:'07AAACC0992L1ZI',pan:'AAACC0992L',tan:'',attn:'Deepak Talwar',mobile:'8810241176',email:'talwardeepak478@gmail.com'},
  {shortName:'DIPANKAR',companyName:'SEEKGROWTH LEARNING SOLUTIONS',addr1:'Pocket 40 / House No. 99 (Basement),',addr2:'Chittaranjan Park,',addr3:'New Delhi 110019',state:'Delhi',gstin:'07AGGPD9121P1Z5',pan:'AGGPD9121P',tan:'',attn:'Dipankar Das',mobile:'9811500184',email:'sirdipankar@gmail.com'},
  {shortName:'DILIP',companyName:'CARPE DIEM',addr1:'2A Apsaras, No. 1, Sambandam Street',addr2:'Off G. N Chetty Road, T. Nagar, Chennai - 600017',addr3:'',state:'Tamil Nadu',gstin:'33AKHPD1440Q1Z8',pan:'AKHPD1440Q',tan:'CHED08636F',attn:'P. Dilip Krishna',mobile:'9840022248',email:'dilipkrishna@carpediemindia.in'},
  {shortName:'HARISH',companyName:'HARISH CN',addr1:'B 001, Ground Floor, B Block,',addr2:'Renaissance Park 1, Malleshwaram West,',addr3:'Bangalore - 560055',state:'Karnataka',gstin:'29AACPH2632P1ZH',pan:'AACPH2632P',tan:'',attn:'Harish Closepet',mobile:'9811500184',email:'harishcn1210@gmail.com'},
  {shortName:'JOHNSON',companyName:'EXCEL TALENT PLUS',addr1:'302, MG Gajapthy Nivas, Plot No.6, Indrapuri Railway Colony,',addr2:'West Marredpally, Secunderabad,',addr3:'Hyderabad-500026',state:'Telangana',gstin:'',pan:'ACTPB6652J',tan:'',attn:'Johnson Baby',mobile:'9490484401',email:'hrjohnsonbaby@gmail.com'},
  {shortName:'KADAMBARI',companyName:'KADAMBARI DEODHAR',addr1:'Flat no.12, Dar-ul-Khalil,',addr2:'Shahid Bhagatsingh Road,',addr3:'Colaba, Mumbai 400 001',state:'Maharashtra',gstin:'27AEQPD4118L1ZA',pan:'AEQPD4118L',tan:'',attn:'Kadambari Deodhar',mobile:'9820129239',email:'kadambari.deodhar@lmi-india.in'},
  {shortName:'NARESH',companyName:'NARESH KUMAR RATTAN',addr1:'H.No. 110, Sector 30,',addr2:'Gurugram 122001',addr3:'',state:'Haryana',gstin:'06AACPR5068D1ZS',pan:'AACPR1068D',tan:'',attn:'Naresh Kumar Rattan',mobile:'9878337710',email:'naresh_rattan@yahoo.com'},
  {shortName:'PAYANK',companyName:'LIFE A SCHOOL OF ATTITUDE AND VALUES PVT LTD',addr1:'120 Fortune Business Hub, Nr Shell Petrol Pump,',addr2:'Science City Road, Ahmedabad - 380060',addr3:'',state:'Gujarat',gstin:'24AADCL8539M1Z1',pan:'AADCL8539M',tan:'',attn:'Payank Patel',mobile:'9904983310',email:'payank.patel140581@gmail.com'},
  {shortName:'RAVI',companyName:'ALACRITY CONSULTING',addr1:'34, Mangalmurti Krishnaji Nagar,',addr2:'Scheme No. 77, Near Khajrana temple,',addr3:'Indore - 452016',state:'Madhya Pradesh',gstin:'23AIEPG7876N1ZR',pan:'AIEPG7876N',tan:'',attn:'Ravi Gupta',mobile:'9893011073',email:'ravi@lmi-india.in'},
  {shortName:'SALIL',companyName:'SALIL CHANDRA',addr1:'L-301, Microtek Greenburg',addr2:'Sector 86',addr3:'Gurgaon',state:'Haryana',gstin:'06AENPC9828D1ZG',pan:'AENPC9828D',tan:'',attn:'Salil Chandra',mobile:'9999114183',email:'salil@lmi-india.in'},
  {shortName:'SUDHIR',companyName:'SUMMIT CONSULTANTS',addr1:'1002 Trishna View CHS, Bhagwan Mahavir Marg,',addr2:'J. B. Nagar, Andheri East, Mumbai 400059',addr3:'',state:'Maharashtra',gstin:'27AABPR2614Q1ZC',pan:'AABPR2614Q',tan:'MUMS82361G',attn:'Sudhir Rao',mobile:'9820282709',email:'rao.ssu@gmail.com'},
  {shortName:'SUNIL J',companyName:'RE; FORM - TRANSFORM TO THRIVE',addr1:'9011, Garden Villas,',addr2:'DLF phase 4,',addr3:'Gurgaon 122009',state:'Haryana',gstin:'06AAAHS8485D1ZZ',pan:'AAAHS8485D',tan:'',attn:'Sunil Jain',mobile:'9717797744',email:'sunil@lmi-india.in'},
  {shortName:'SUNIL N',companyName:'GOLDEN SKY VENTURES',addr1:'Plot No. 76, Phase 1, Adithya Homes,',addr2:'Nerige PO, Kamanahalli Village Circle, Sarjapura,',addr3:'Bengaluru- 562125',state:'Karnataka',gstin:'29ADEPN6889M1ZI',pan:'ADEPN6889M',tan:'',attn:'Sunil Nair',mobile:'9972189026',email:'sknglobal@gmail.com'},
  {shortName:'SUPARNA',companyName:'ENVISAGE TALENT SOLUTIONS PVT LTD',addr1:'B909 Onkar, Shivdham Sankul, Above Axis Bank,',addr2:'Opposite Fire Brigade, Near Oberoi Mall,',addr3:'Malad East, Mumbai 400097',state:'Maharashtra',gstin:'27AACCE4605N1ZI',pan:'AACCE4605N',tan:'',attn:'Suparna Samant',mobile:'9321381326',email:'suparna@lmi-india.in'},
  {shortName:'VINOD',companyName:'VINOD SETHUMADHAVAN WARRIER',addr1:'B -503, Akruti Aneri, Behind Seven Hills Hospital,',addr2:'Marol Maroshi Road, Marol Andheri (East) Mumbai - 400059',addr3:'',state:'Maharashtra',gstin:'27AAAPW7337E1ZJ',pan:'AAAPW7337E',tan:'',attn:'Vinod Warrier',mobile:'9324060153',email:'vinodswarrier@gmail.com'},
  {shortName:'ZUBAIR',companyName:'CONSULTZUBAIR PRIVATE LTD.',addr1:'Mandir Bagh, Baghat,',addr2:'Barzulla, Srinagar, India',addr3:'',state:'Jammu & Kashmir',gstin:'',pan:'AALCC4514F',tan:'',attn:'Zubair Iqbal',mobile:'8491999000',email:'zubair.iqball@gmail.com'},
];

function invImportBundledClients() {
  invInit();
  let added = 0;
  BUNDLED_CLIENTS.forEach(c => {
    const exists = DB.invoicing.clients.find(x => x.shortName === c.shortName);
    if (!exists) {
      DB.invoicing.clients.push({ ...c, id: uid() });
      added++;
    }
  });
  saveDB();
  closeModal();
  invOpenClients();
  toast(`${added} client${added !== 1 ? 's' : ''} imported`);
}

// ── Start Numbers ──────────────────────────────────────────
function invOpenStartNumbers() {
  invInit();
  const fy = currentInvFY();
  const piCurrent = DB.invoicing.piSequence[fy] || 0;
  const tiCurrent = DB.invoicing.tiSequence[fy] || 0;
  const body = `
    <div class="field"><label>Financial year</label>
      <select id="sn-fy">${['26-27','27-28','28-29','29-30'].map(f =>
        `<option value="${f}" ${f===fy?'selected':''}>${f}</option>`).join('')}</select>
    </div>
    <div class="field-row">
      <div class="field"><label>PI — next number</label>
        <input type="number" id="sn-pi" value="${piCurrent + 1}" min="1">
        <div style="font-size:11px; color:var(--ink-soft); margin-top:4px;">Current last used: ${piCurrent === 0 ? 'none' : `PI/${fy}/${String(piCurrent).padStart(3,'0')}`}</div>
      </div>
      <div class="field"><label>TI — next number</label>
        <input type="number" id="sn-ti" value="${tiCurrent + 1}" min="1">
        <div style="font-size:11px; color:var(--ink-soft); margin-top:4px;">Current last used: ${tiCurrent === 0 ? 'none' : `TI/${fy}/${String(tiCurrent).padStart(3,'0')}`}</div>
      </div>
    </div>
    <div class="hint">Setting the next number to e.g. 18 means the next invoice generated will be PI/${fy}/018. This does not create any invoices — just sets the counter.</div>`;
  openModal('Set invoice start numbers', body,
    `<button class="btn" id="sn-cancel">Cancel</button><button class="btn btn-primary" id="sn-save">Save</button>`);
  document.getElementById('sn-cancel').onclick = closeModal;
  document.getElementById('sn-save').onclick = () => {
    const fy2 = document.getElementById('sn-fy').value;
    const pi = Math.max(0, (parseInt(document.getElementById('sn-pi').value) || 1) - 1);
    const ti = Math.max(0, (parseInt(document.getElementById('sn-ti').value) || 1) - 1);
    DB.invoicing.piSequence[fy2] = pi;
    DB.invoicing.tiSequence[fy2] = ti;
    saveDB(); closeModal();
    toast(`Counters set: next PI will be PI/${fy2}/${String(pi+1).padStart(3,'0')}, next TI will be TI/${fy2}/${String(ti+1).padStart(3,'0')}`);
  };
}

// ── Invoice form (PI) ──────────────────────────────────────
// ── Nature of Invoice helpers ──────────────────────────────
const DEFAULT_NATURES = ['PROG SALE', 'SALES SUPPORT', 'LICENSEE FEE', 'EMAIL Cost'];

function invGetNatures() {
  invInit();
  if (!DB.invoicing.natures) DB.invoicing.natures = [...DEFAULT_NATURES];
  return DB.invoicing.natures;
}

function invGetNatureOptions(selected) {
  const natures = invGetNatures();
  const opts = natures.map(n =>
    `<option value="${escapeHtml(n)}" ${n === selected ? 'selected' : ''}>${escapeHtml(n)}</option>`
  ).join('');
  // Always include OTHER as last option — triggers manual entry
  const otherSel = selected === '__other__' ? 'selected' : '';
  return opts + `<option value="__other__" ${otherSel}>OTHER (manual entry)</option>`;
}

function invAddNature(currentSelectId) {
  const name = prompt('Add new Nature of Invoice:');
  if (!name || !name.trim()) return;
  const natures = invGetNatures();
  if (natures.includes(name.trim())) { toast('Already exists'); return; }
  natures.push(name.trim());
  DB.invoicing.natures = natures;
  saveDB();
  // Refresh dropdown
  const sel = document.getElementById(currentSelectId);
  if (sel) {
    sel.innerHTML = invGetNatureOptions(name.trim());
    sel.value = name.trim();
  }
  toast(`Nature "${name.trim()}" added`);
}

function invCalcReceivable(taxable, gstType, tdsDeducted) {
  const intra = gstType === 'intra';
  const cgst = intra ? taxable * 0.09 : 0;
  const sgst = intra ? taxable * 0.09 : 0;
  const igst = intra ? 0 : taxable * 0.18;
  const grossBeforeTDS = taxable + cgst + sgst + igst;
  if (tdsDeducted === 'yes') {
    const tds = taxable * 0.10;
    return Math.round((taxable - tds + cgst + sgst + igst) * 100) / 100;
  }
  return Math.round(grossBeforeTDS * 100) / 100;
}

function invOpenPIForm(editId) {
  invInit();
  // If called from dashboard (not from invoicing tab), open the module first
  const overlay = document.getElementById('invoicingOverlay');
  if (overlay && overlay.style.display === 'none') {
    openInvoicingModule();
  }
  const existing = editId ? DB.invoicing.proformas.find(p => p.id === editId) : null;
  const v = existing || {};
  const clients = DB.invoicing.clients;
  const today = new Date().toISOString().slice(0, 10);
  const fy = currentInvFY();
  const nextPiNum = `PI/${fy}/${String((DB.invoicing.piSequence[fy]||0)+1).padStart(3,'0')}`;

  if (!clients.length) {
    toast('No clients found — please import or add clients first (click Invoicing tab → Clients)');
    return;
  }

  const body = `
    <div class="hint" style="background:#e8f0fb; margin-bottom:10px;">
      ${existing ? `Editing ${escapeHtml(existing.invNo)}` : `Next PI number will be: <strong>${nextPiNum}</strong>`}
    </div>
    <div class="field-row">
      <div class="field"><label>Client</label>
        <select id="pf-client">${clients.map(c =>
          `<option value="${c.id}" ${v.clientId===c.id?'selected':''}>${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</option>`
        ).join('')}</select>
      </div>
      <div class="field"><label>Invoice date</label>
        <input type="date" id="pf-date" value="${v.date || today}"
          onchange="document.getElementById('pf-supply').value=this.value">
      </div>
    </div>
    <div class="field-row">
      <div class="field"><label>Date of supply</label>
        <input type="date" id="pf-supply" value="${v.supplyDate || v.date || today}">
      </div>
      <div class="field"><label>Payment terms</label>
        <input id="pf-terms" value="${escapeHtml(v.paymentTerms || 'Advance')}">
      </div>
    </div>

    <!-- Dynamic line items -->
    <div style="border:1px solid var(--line); border-radius:6px; padding:12px; margin-bottom:8px;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <span style="font-size:12px; font-weight:700; color:var(--navy); text-transform:uppercase; letter-spacing:.05em;">Program / Service lines</span>
        <div style="display:flex;gap:6px;">
          <button type="button" id="pf-add-line" class="btn btn-sm">+ Add program</button>
          <button type="button" id="pf-add-freight" class="btn btn-sm" style="background:var(--gold);border-color:var(--gold);color:#fff;">+ Add freight</button>
        </div>
      </div>
      <div id="pf-lines-wrap"></div>
    </div>

    <div class="field-row">
      <div class="field"><label>Nature of invoice</label>
        <div style="display:flex;gap:6px;">
          <select id="pf-nature" style="flex:1;">
            ${invGetNatureOptions(v.nature||'')}
          </select>
          <button type="button" id="pf-nature-add" class="btn btn-sm" title="Add new nature">+</button>
        </div>
        <input type="text" id="pf-nature-manual" placeholder="Enter nature manually"
          value="${v.nature && !invGetNatures().includes(v.nature) ? escapeHtml(v.nature) : ''}"
          style="margin-top:6px; width:100%; display:${(v.nature && !invGetNatures().includes(v.nature) && v.nature !== '') ? 'block' : 'none'};">
      </div>
      <div class="field"><label>SAC code</label>
        <input id="pf-sac" value="${v.sac || '998399'}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>TDS deducted?</label>
        <select id="pf-tds">
          <option value="no" ${(v.tdsDeducted||'no')==='no'?'selected':''}>No</option>
          <option value="yes" ${(v.tdsDeducted||'no')==='yes'?'selected':''}>Yes — 10% TDS on taxable amount</option>
        </select>
        <div style="font-size:11px;color:var(--ink-soft);margin-top:3px;">If YES: receivable = (taxable − 10% TDS) + GST</div>
      </div>
    </div>
    <div id="pf-gst-preview" style="background:var(--paper); border-radius:6px; padding:10px 14px; font-size:12.5px; margin-top:4px;"></div>`;

  openModal(existing ? 'Edit Proforma Invoice' : 'New Proforma Invoice', body,
    `<button class="btn" id="pf-cancel">Cancel</button><button class="btn btn-primary" id="pf-save">${existing ? 'Save changes' : 'Generate PI'}</button>`);

  // Initialise line items from existing data or blank
  // Each line has: _id, _type ('prog'|'freight'), desc, shortName, rate, qty, disc
  const pfProducts = DB.invoicing.products || [];
  const offlineOnlineProds = pfProducts.filter(p => p.category !== 'OTHER');
  const otherProds = pfProducts.filter(p => p.category === 'OTHER');

  let pfLines = [];
  if (existing && existing.lineItems && existing.lineItems.length) {
    pfLines = existing.lineItems.map(l => ({
      ...l, _id: uid(),
      _type: (l._type === 'freight' || l.desc === 'Freight / Courier charges') ? 'freight' : 'prog'
    }));
  } else {
    pfLines = [{ _id: uid(), _type: 'prog', desc: '', shortName: pfProducts.length ? '' : '__manual__', rate: 0, qty: 1, disc: 0 }];
  }

  function pfProgCount() { return pfLines.filter(l => l._type === 'prog').length; }
  function pfHasFreight() { return pfLines.some(l => l._type === 'freight'); }

  function renderPfLines() {
    const wrap = document.getElementById('pf-lines-wrap');
    if (!wrap) return;

    // Update Add program button — disabled at 3 program lines
    const addBtn = document.getElementById('pf-add-line');
    if (addBtn) {
      const progCount = pfProgCount();
      addBtn.disabled = progCount >= 3;
      addBtn.title = progCount >= 3 ? 'Maximum 3 program lines (A, B, C)' : '';
      addBtn.style.opacity = progCount >= 3 ? '0.5' : '1';
    }
    // Update Add freight button — disabled if already has freight
    const freightBtn = document.getElementById('pf-add-freight');
    if (freightBtn) {
      freightBtn.disabled = pfHasFreight();
      freightBtn.style.opacity = pfHasFreight() ? '0.5' : '1';
      freightBtn.title = pfHasFreight() ? 'Freight already added' : 'Add freight / courier line';
    }

    // Build row labels — A, B, C for programs, then freight gets next letter
    let progIdx = 0, freightLabel = '';
    const labels = pfLines.map(l => {
      if (l._type === 'prog') return String.fromCharCode(65 + progIdx++);
      freightLabel = String.fromCharCode(65 + progIdx++);
      return freightLabel;
    });
    // Reset progIdx for actual program count
    progIdx = 0;

    wrap.innerHTML = pfLines.map((line, i) => {
      const label = labels[i];
      const isFreight = line._type === 'freight';
      const canDelete = pfLines.length > 1 || isFreight;

      if (isFreight) {
        // Freight row — simple amount entry
        return `
          <div style="border:1px solid #dda63a; border-radius:5px; padding:10px; margin-bottom:8px; background:#fffbf0;" data-pf-line="${line._id}">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:6px;">
              <span style="font-size:11px; font-weight:700; color:#9a6b14;">ROW ${label} — Freight / Courier</span>
              <button type="button" data-pf-del-line="${line._id}" style="border:none;background:none;color:#aab2bd;cursor:pointer;font-size:13px;">✕</button>
            </div>
            <div style="display:flex;gap:8px;align-items:center;">
              <div class="field" style="flex:2;margin:0"><label style="font-size:11px;">Description</label>
                <input type="text" data-pf-line-desc="${line._id}" value="${escapeHtml(line.desc||'Freight / Courier charges')}" style="width:100%;"></div>
              <div class="field" style="flex:1;margin:0"><label style="font-size:11px;">Amount (₹)</label>
                <input type="number" data-pf-line-rate="${line._id}" value="${line.rate||''}" placeholder="0"></div>
            </div>
          </div>`;
      }

      // Program row — PROGRAM dropdown + Qty + Rate
      const allProds = pfProducts;
      const selectedProd = allProds.find(p => p.shortName === line.shortName);
      const isManual = !selectedProd || line.shortName === '__manual__' || !line.shortName;
      const shortDisplayVal = isManual ? (line.shortName === '__manual__' ? '' : (line.shortName||'')) : (line.shortName||'');

      return `
        <div style="border:1px solid ${isManual?'#9a6b14':'var(--line)'}; border-radius:5px; padding:10px; margin-bottom:8px; background:${isManual?'#fffbf0':'var(--paper)'};" data-pf-line="${line._id}">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
            <span style="font-size:11px; font-weight:700; color:var(--navy);">ROW ${label}</span>
            ${canDelete ? `<button type="button" data-pf-del-line="${line._id}" style="border:none;background:none;color:#aab2bd;cursor:pointer;font-size:13px;">✕</button>` : ''}
          </div>
          ${allProds.length ? `
          <div class="field" style="margin-bottom:8px;">
            <label style="font-size:11px;">Program (or choose Enter manually for one-off)</label>
            <select data-pf-line-prog="${line._id}" style="width:100%; font-size:13px;">
              <option value="">— Select program —</option>
              ${offlineOnlineProds.length ? `<optgroup label="OFFLINE / ONLINE">
                ${offlineOnlineProds.map(p => `<option value="${p.shortName}" ${p.shortName===line.shortName?'selected':''}>${escapeHtml(p.shortName)} — ${escapeHtml(p.longName)}</option>`).join('')}
              </optgroup>` : ''}
              ${otherProds.length ? `<optgroup label="OTHER">
                ${otherProds.map(p => `<option value="${p.shortName}" ${p.shortName===line.shortName?'selected':''}>${escapeHtml(p.shortName)} — ${escapeHtml(p.longName)}</option>`).join('')}
              </optgroup>` : ''}
              <option value="__manual__" ${isManual?'selected':''}>— Enter manually —</option>
            </select>
          </div>` : ''}
          <div style="display:flex;gap:8px;margin-bottom:8px;">
            <div class="field" style="flex:0 0 120px;margin:0">
              <label style="font-size:11px;">Short name${isManual?' <span style="color:#9a6b14;font-weight:700;">(type here)</span>':''}</label>
              <input type="text" data-pf-line-short="${line._id}"
                value="${escapeHtml(shortDisplayVal)}"
                placeholder="${isManual?'e.g. EPP':'auto-filled'}"
                style="${isManual?'border-color:#9a6b14;background:#fff8f0;':'background:var(--paper-dim,#f5f5f5);color:var(--ink-soft);'}"
                ${!isManual ? 'readonly' : ''}>
            </div>
            <div class="field" style="flex:1;margin:0">
              <label style="font-size:11px;">Rate (₹)</label>
              <input type="number" data-pf-line-rate="${line._id}" value="${line.rate||''}" placeholder="0">
            </div>
            <div class="field" style="flex:0 0 80px;margin:0">
              <label style="font-size:11px;">Qty</label>
              <input type="number" data-pf-line-qty="${line._id}" value="${line.qty||1}" min="1">
            </div>
          </div>
          <div class="field" style="margin:0">
            <label style="font-size:11px;">Description${isManual?' <span style="color:#9a6b14;font-weight:700;">(type here)</span>':' (auto-filled, editable)'}</label>
            <input type="text" data-pf-line-desc="${line._id}" value="${escapeHtml(line.desc||'')}"
              placeholder="${isManual?'Enter full description e.g. One-time setup charges':'Auto-filled from product'}">
          </div>
        </div>`;
    }).join('');

    // Wire delete
    wrap.querySelectorAll('[data-pf-del-line]').forEach(b => {
      b.onclick = () => {
        pfLines = pfLines.filter(l => l._id !== b.dataset.pfDelLine);
        renderPfLines(); pfPreview();
      };
    });

    // Wire program selector
    wrap.querySelectorAll('[data-pf-line-prog]').forEach(sel => {
      sel.onchange = () => {
        const val = sel.value;
        const l = pfLines.find(x => x._id === sel.dataset.pfLineProg);
        if (!l) return;
        l.shortName = val;
        if (val === '__manual__' || !val) {
          l.desc = '';
          l.rate = 0;
          // Don't clear shortName display — user will type it
        } else {
          const prod = pfProducts.find(p => p.shortName === val);
          if (prod) {
            l.rate = prod.rate;
            l.desc = prod.category === 'OTHER' ? prod.longName : `Sale of ${prod.longName}`;
          }
        }
        renderPfLines(); pfPreview();
      };
    });

    // Wire short name input (manual entry)
    wrap.querySelectorAll('[data-pf-line-short]').forEach(inp => {
      inp.oninput = () => {
        const l = pfLines.find(x => x._id === inp.dataset.pfLineShort);
        if (l) { l.shortName = inp.value || '__manual__'; }
      };
    });

    // Wire desc input (manual edit)
    wrap.querySelectorAll('[data-pf-line-desc]').forEach(inp => {
      inp.oninput = () => {
        const l = pfLines.find(x => x._id === inp.dataset.pfLineDesc);
        if (l) { l.desc = inp.value; pfPreview(); }
      };
    });

    // Wire rate input
    wrap.querySelectorAll('[data-pf-line-rate]').forEach(inp => {
      inp.oninput = () => {
        const l = pfLines.find(x => x._id === inp.dataset.pfLineRate);
        if (l) { l.rate = parseFloat(inp.value) || 0; pfPreview(); }
      };
    });

    // Wire qty input
    wrap.querySelectorAll('[data-pf-line-qty]').forEach(inp => {
      inp.oninput = () => {
        const l = pfLines.find(x => x._id === inp.dataset.pfLineQty);
        if (l) { l.qty = parseFloat(inp.value) || 1; pfPreview(); }
      };
    });
  }

  const pfPreview = () => {
    const cid = document.getElementById('pf-client').value;
    const cl = clients.find(c => c.id === cid);
    const tdsDeducted = document.getElementById('pf-tds') ? document.getElementById('pf-tds').value : 'no';
    // Programs: rate × qty. Freight: just the rate value (qty=1)
    const linesTaxable = pfLines.reduce((s, l) => {
      if (l._type === 'freight') return s + (l.rate || 0);
      return s + ((l.rate||0) * (l.qty||1)) - (l.disc||0);
    }, 0);
    const taxable = linesTaxable;
    const type = cl ? gstType(cl.state) : 'igst';
    const intra = type === 'intra';
    const cgst = intra ? taxable * 0.09 : 0;
    const sgst = intra ? taxable * 0.09 : 0;
    const igst = intra ? 0 : taxable * 0.18;
    const gross = taxable + cgst + sgst + igst;
    const receivable = invCalcReceivable(taxable, type, tdsDeducted);
    const gstLabel = intra ? 'CGST 9% + SGST 9%' : 'IGST 18%';
    const tdsNote = tdsDeducted === 'yes'
      ? ` | TDS 10% = ${fmtMoney(taxable*0.1)} | <strong>Receivable: ${fmtMoney(receivable)}</strong>`
      : ` | <strong>Receivable: ${fmtMoney(receivable)}</strong>`;
    document.getElementById('pf-gst-preview').innerHTML =
      `<b>Preview:</b> Taxable ${fmtMoney(taxable)} + ${gstLabel} = Gross ${fmtMoney(gross)}${tdsNote}`;
    return { linesTaxable, taxable, cgst, sgst, igst, gross, type, tdsDeducted, receivable };
  };

  renderPfLines();
  pfPreview();

  // Add program button — max 3 program rows
  document.getElementById('pf-add-line').onclick = () => {
    if (pfProgCount() >= 3) { toast('Maximum 3 program lines (A, B, C) per invoice'); return; }
    // Insert before freight if it exists, otherwise at end
    const freightIdx = pfLines.findIndex(l => l._type === 'freight');
    const newLine = { _id: uid(), _type: 'prog', desc: '', shortName: pfProducts.length ? '' : '__manual__', rate: 0, qty: 1, disc: 0 };
    if (freightIdx >= 0) pfLines.splice(freightIdx, 0, newLine);
    else pfLines.push(newLine);
    renderPfLines(); pfPreview();
  };

  // Add freight button — only one freight row, always goes at end
  document.getElementById('pf-add-freight').onclick = () => {
    if (pfHasFreight()) { toast('Freight line already added'); return; }
    pfLines.push({ _id: uid(), _type: 'freight', desc: 'Freight / Courier charges', rate: 0, qty: 1, disc: 0 });
    renderPfLines(); pfPreview();
  };

  // Wire Nature add button and OTHER manual input
  const pfNatureAddBtn = document.getElementById('pf-nature-add');
  if (pfNatureAddBtn) pfNatureAddBtn.onclick = () => invAddNature('pf-nature');
  const pfNatureSel = document.getElementById('pf-nature');
  const pfNatureManual = document.getElementById('pf-nature-manual');
  if (pfNatureSel && pfNatureManual) {
    pfNatureSel.onchange = () => {
      pfNatureManual.style.display = pfNatureSel.value === '__other__' ? 'block' : 'none';
      if (pfNatureSel.value === '__other__') pfNatureManual.focus();
    };
  }

  ['pf-client','pf-tds'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', pfPreview);
    if (el && el.tagName === 'SELECT') el.addEventListener('change', pfPreview);
  });

  document.getElementById('pf-cancel').onclick = closeModal;
  document.getElementById('pf-save').onclick = () => {
    const cid = document.getElementById('pf-client').value;
    const cl = clients.find(c => c.id === cid);
    if (!cl) { toast('Select a client'); return; }
    if (!pfLines.length || !pfLines[0].desc.trim()) { toast('At least one line item with a description is required'); return; }
    const calc = pfPreview();
    // desc/unit/rate for backward compat (first line item)
    const firstLine = pfLines[0];
    const natureRaw = document.getElementById('pf-nature') ? document.getElementById('pf-nature').value : '';
    const nature = natureRaw === '__other__'
      ? (document.getElementById('pf-nature-manual') ? document.getElementById('pf-nature-manual').value.trim() : '')
      : natureRaw;
    const tdsDeducted = document.getElementById('pf-tds') ? document.getElementById('pf-tds').value : 'no';
    const rec = {
      id: existing ? existing.id : uid(),
      invNo: existing ? existing.invNo : nextPINumber(),
      date: document.getElementById('pf-date').value,
      supplyDate: document.getElementById('pf-supply').value || document.getElementById('pf-date').value,
      paymentTerms: document.getElementById('pf-terms').value,
      nature, tdsDeducted,
      clientId: cid, clientName: cl.companyName, clientShort: cl.shortName,
      // Multi-line items — preserve _type for freight rows
      lineItems: pfLines.map(l => ({
        _type: l._type || 'prog',
        desc: l.desc,
        shortName: l.shortName || '',
        rate: l.rate,
        qty: l._type === 'freight' ? 1 : (l.qty||1),
        disc: l.disc||0
      })),
      // Single-item fields for backward compat
      desc: pfLines.map(l => l.desc).filter(Boolean).join('; '),
      unit: firstLine.unit,
      rate: firstLine.rate,
      qty: firstLine.qty || 1,
      sac: document.getElementById('pf-sac').value || '998399',
      disc: pfLines.reduce((s,l)=>s+(l.disc||0),0),
      taxable: calc.taxable, cgst: calc.cgst, sgst: calc.sgst,
      igst: calc.igst, gross: calc.gross, gstType: calc.type,
      receivableAmount: calc.receivable, // may differ from gross if TDS
      status: existing ? existing.status : 'draft',
    };
    if (existing) {
      const oldRecv = existing.receivableAmount || existing.gross;
      Object.assign(existing, rec);
      if (oldRecv !== rec.receivableAmount) invUpdateReceivableAmount(rec.invNo, rec.receivableAmount);
    } else {
      DB.invoicing.proformas.push(rec);
      invAddToReceivables(rec, 'PI');
    }
    DB.invoicing.piSequence[currentInvFY()] = parseInt(rec.invNo.split('/').pop(), 10);
    saveDB(); closeModal(); invRenderRegister();
    toast(`${rec.invNo} ${existing ? 'updated' : 'generated'}`);
  };
}

// ── Invoice form (TI) ──────────────────────────────────────
function invOpenTIForm(fromPiId, editTiId, fromDashboard) {
  invInit();
  if (fromDashboard) {
    const overlay = document.getElementById('invoicingOverlay');
    if (overlay && overlay.style.display === 'none') openInvoicingModule();
  }
  const existingTI = editTiId ? DB.invoicing.taxInvoices.find(t => t.id === editTiId) : null;
  const clients = DB.invoicing.clients;

  if (!clients.length) {
    toast('No clients found — add clients first (Invoicing tab → Clients)');
    return;
  }

  // If we already have a source PI or existing TI, skip the YES/NO step
  if (fromPiId || existingTI) {
    _invOpenTIFormInner(fromPiId, existingTI);
    return;
  }

  // ── Step 1: Ask "From existing PI?" ──────────────────────────
  const openPIs = DB.invoicing.proformas
    .filter(p => p.status !== 'cancelled' && p.status !== 'converted')
    .sort((a, b) => (b.invNo||'').localeCompare(a.invNo||''));

  const body = `
    <div style="text-align:center; padding:10px 0 18px;">
      <div style="font-size:15px; font-weight:700; color:var(--navy); margin-bottom:6px;">
        Is this Tax Invoice based on an existing Proforma Invoice?
      </div>
      <div style="font-size:13px; color:var(--ink-soft);">
        YES — pick a PI and auto-populate the details<br>
        NO — create a fresh Tax Invoice
      </div>
    </div>
    <div style="display:flex; gap:12px; justify-content:center;">
      <button class="btn btn-primary" id="ti-from-pi-yes" style="min-width:100px; font-size:15px;">YES</button>
      <button class="btn" id="ti-from-pi-no" style="min-width:100px; font-size:15px;">NO</button>
    </div>`;

  openModal('New Tax Invoice', body,
    `<button class="btn" id="ti-step1-cancel">Cancel</button>`);

  document.getElementById('ti-step1-cancel').onclick = closeModal;

  document.getElementById('ti-from-pi-no').onclick = () => {
    closeModal();
    _invOpenTIFormInner(null, null);
  };

  document.getElementById('ti-from-pi-yes').onclick = () => {
    // Show PI picker
    if (!openPIs.length) {
      toast('No open Proforma Invoices found — use NO to create a fresh TI');
      return;
    }
    const pickerBody = `
      <div class="hint" style="margin-bottom:12px;">Select the Proforma Invoice to convert to a Tax Invoice.</div>
      <div style="max-height:360px; overflow-y:auto;">
        ${openPIs.map(pi => `
          <div class="sub-list-item" style="cursor:pointer; display:flex; align-items:center; gap:10px;"
               data-pi-pick="${pi.id}">
            <div style="flex:1;">
              <div style="font-weight:700; color:var(--navy); font-size:13px;">${escapeHtml(pi.invNo)}</div>
              <div style="font-size:12px; color:var(--ink-soft);">${escapeHtml(pi.clientName)} — ${pi.date} — ${fmtMoney(pi.gross)}</div>
              ${pi.desc ? `<div style="font-size:11px; color:var(--ink-soft);">${escapeHtml(pi.desc.slice(0,60))}</div>` : ''}
            </div>
            <button class="btn btn-sm btn-primary" data-pi-pick="${pi.id}">Select →</button>
          </div>`).join('')}
      </div>`;

    openModal('Select Proforma Invoice', pickerBody,
      `<button class="btn" id="ti-pi-pick-back">← Back</button>`);

    document.getElementById('ti-pi-pick-back').onclick = () => invOpenTIForm(null, null, false);

    document.querySelectorAll('[data-pi-pick]').forEach(el => {
      el.onclick = () => {
        const piId = el.dataset.piPick;
        closeModal();
        _invOpenTIFormInner(piId, null);
      };
    });
  };
}

// ── TI inner form — shared by both PI-sourced and fresh TI ────
function _invOpenTIFormInner(fromPiId, existingTI) {
  invInit();
  const sourcePi = fromPiId ? DB.invoicing.proformas.find(p => p.id === fromPiId) : null;
  const v = existingTI || sourcePi || {};
  const clients = DB.invoicing.clients;
  const today = new Date().toISOString().slice(0, 10);
  const fy = currentInvFY();
  const nextTiNum = `TI/${fy}/${String((DB.invoicing.tiSequence[fy]||0)+1).padStart(3,'0')}`;

  const pfProducts = DB.invoicing.products || [];
  const offlineOnlineProds = pfProducts.filter(p => p.category !== 'OTHER');
  const otherProds = pfProducts.filter(p => p.category === 'OTHER');

  // Initialise line items from source
  let tiLines = [];
  if (existingTI && existingTI.lineItems && existingTI.lineItems.length) {
    tiLines = existingTI.lineItems.map(l => ({ ...l, _id: uid() }));
  } else if (sourcePi && sourcePi.lineItems && sourcePi.lineItems.length) {
    tiLines = sourcePi.lineItems.map(l => ({ ...l, _id: uid() }));
  } else if (v.desc) {
    tiLines = [{ _id: uid(), _type: 'prog', desc: v.desc, shortName: v.unit||'__manual__', rate: v.rate||0, qty: v.qty||1, disc: v.disc||0 }];
  } else {
    tiLines = [{ _id: uid(), _type: 'prog', desc: '', shortName: pfProducts.length ? '' : '__manual__', rate: 0, qty: 1, disc: 0 }];
  }

  const body = `
    <div class="hint" style="background:#e8f0fb; margin-bottom:10px;">
      ${sourcePi ? `<strong>From PI:</strong> ${escapeHtml(sourcePi.invNo)} — ${escapeHtml(sourcePi.clientName)} · ` : ''}
      ${existingTI ? `Editing ${escapeHtml(existingTI.invNo)}` : `Next TI: <strong>${nextTiNum}</strong>`}
    </div>
    <div class="field-row">
      <div class="field"><label>Client</label>
        <select id="ti-client">${clients.map(c =>
          `<option value="${c.id}" ${(v.clientId||'')=== c.id?'selected':''}>${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</option>`
        ).join('')}</select>
      </div>
      <div class="field"><label>Invoice date</label>
        <input type="date" id="ti-date" value="${existingTI ? existingTI.date : today}"
          onchange="document.getElementById('ti-supply').value=this.value">
      </div>
    </div>
    <div class="field-row">
      <div class="field"><label>Date of supply</label>
        <input type="date" id="ti-supply" value="${v.supplyDate || (existingTI ? existingTI.date : today)}">
      </div>
      <div class="field"><label>Payment terms</label>
        <input id="ti-terms" value="${escapeHtml(v.paymentTerms || 'Advance')}">
      </div>
    </div>

    <!-- Dynamic line items (same as PI) -->
    <div style="border:1px solid var(--line); border-radius:6px; padding:12px; margin-bottom:8px;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <span style="font-size:12px; font-weight:700; color:var(--navy); text-transform:uppercase; letter-spacing:.05em;">Program / Service lines</span>
        <div style="display:flex;gap:6px;">
          <button type="button" id="ti-add-line" class="btn btn-sm">+ Add program</button>
          <button type="button" id="ti-add-freight" class="btn btn-sm" style="background:var(--gold);border-color:var(--gold);color:#fff;">+ Add freight</button>
        </div>
      </div>
      <div id="ti-lines-wrap"></div>
    </div>

    <div class="field-row">
      <div class="field"><label>Nature of invoice</label>
        <div style="display:flex;gap:6px;">
          <select id="ti-nature" style="flex:1;">
            ${invGetNatureOptions(v.nature||'')}
          </select>
          <button type="button" id="ti-nature-add" class="btn btn-sm" title="Add nature">+</button>
        </div>
        <input type="text" id="ti-nature-manual" placeholder="Enter nature manually"
          value="${v.nature && !invGetNatures().includes(v.nature) ? escapeHtml(v.nature) : ''}"
          style="margin-top:6px; width:100%; display:${(v.nature && !invGetNatures().includes(v.nature) && v.nature !== '') ? 'block' : 'none'};">
      </div>
      <div class="field"><label>SAC code</label>
        <input id="ti-sac" value="${v.sac || '998399'}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>TDS deducted?</label>
        <select id="ti-tds">
          <option value="no" ${(v.tdsDeducted||'no')==='no'?'selected':''}>No</option>
          <option value="yes" ${(v.tdsDeducted||'no')==='yes'?'selected':''}>Yes — 10% TDS on taxable amount</option>
        </select>
      </div>
    </div>
    <div id="ti-gst-preview" style="background:var(--paper); border-radius:6px; padding:10px 14px; font-size:12.5px; margin-top:4px;"></div>`;

  openModal(existingTI ? 'Edit Tax Invoice' : 'New Tax Invoice', body,
    `<button class="btn" id="ti-cancel">${sourcePi ? '← Back' : 'Cancel'}</button>
     <button class="btn btn-primary" id="ti-save">${existingTI ? 'Save changes' : 'Generate TI'}</button>`);

  // ── Line rendering (mirrors PI form) ──────────────────────
  function tiProgCount() { return tiLines.filter(l => l._type === 'prog').length; }
  function tiHasFreight() { return tiLines.some(l => l._type === 'freight'); }

  function renderTiLines() {
    const wrap = document.getElementById('ti-lines-wrap');
    if (!wrap) return;
    const addBtn = document.getElementById('ti-add-line');
    if (addBtn) { addBtn.disabled = tiProgCount() >= 3; addBtn.style.opacity = tiProgCount() >= 3 ? '0.5' : '1'; }
    const freightBtn = document.getElementById('ti-add-freight');
    if (freightBtn) { freightBtn.disabled = tiHasFreight(); freightBtn.style.opacity = tiHasFreight() ? '0.5' : '1'; }

    let progIdx = 0;
    const labels = tiLines.map(l => {
      const label = String.fromCharCode(65 + progIdx++);
      return label;
    });

    wrap.innerHTML = tiLines.map((line, i) => {
      const label = labels[i];
      const isFreight = line._type === 'freight';
      const canDelete = tiLines.length > 1 || isFreight;
      const selectedProd = pfProducts.find(p => p.shortName === line.shortName);
      const isManual = !selectedProd || line.shortName === '__manual__' || !line.shortName;

      if (isFreight) return `
        <div style="border:1px solid #dda63a; border-radius:5px; padding:10px; margin-bottom:8px; background:#fffbf0;" data-ti-line="${line._id}">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:6px;">
            <span style="font-size:11px; font-weight:700; color:#9a6b14;">ROW ${label} — Freight / Courier</span>
            <button type="button" data-ti-del-line="${line._id}" style="border:none;background:none;color:#aab2bd;cursor:pointer;font-size:13px;">✕</button>
          </div>
          <div style="display:flex;gap:8px;">
            <div class="field" style="flex:2;margin:0"><label style="font-size:11px;">Description</label>
              <input type="text" data-ti-line-desc="${line._id}" value="${escapeHtml(line.desc||'Freight / Courier charges')}" style="width:100%;"></div>
            <div class="field" style="flex:1;margin:0"><label style="font-size:11px;">Amount (₹)</label>
              <input type="number" data-ti-line-rate="${line._id}" value="${line.rate||''}" placeholder="0"></div>
          </div>
        </div>`;

      return `
        <div style="border:1px solid ${isManual?'#9a6b14':'var(--line)'}; border-radius:5px; padding:10px; margin-bottom:8px; background:${isManual?'#fffbf0':'var(--paper)'};" data-ti-line="${line._id}">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
            <span style="font-size:11px; font-weight:700; color:var(--navy);">ROW ${label}</span>
            ${canDelete ? `<button type="button" data-ti-del-line="${line._id}" style="border:none;background:none;color:#aab2bd;cursor:pointer;font-size:13px;">✕</button>` : ''}
          </div>
          ${pfProducts.length ? `
          <div class="field" style="margin-bottom:8px;">
            <label style="font-size:11px;">Program (or choose Enter manually for one-off)</label>
            <select data-ti-line-prog="${line._id}" style="width:100%; font-size:13px;">
              <option value="">— Select —</option>
              ${offlineOnlineProds.length ? `<optgroup label="OFFLINE / ONLINE">
                ${offlineOnlineProds.map(p => `<option value="${p.shortName}" ${p.shortName===line.shortName?'selected':''}>${escapeHtml(p.shortName)} — ${escapeHtml(p.longName)}</option>`).join('')}
              </optgroup>` : ''}
              ${otherProds.length ? `<optgroup label="OTHER">
                ${otherProds.map(p => `<option value="${p.shortName}" ${p.shortName===line.shortName?'selected':''}>${escapeHtml(p.shortName)} — ${escapeHtml(p.longName)}</option>`).join('')}
              </optgroup>` : ''}
              <option value="__manual__" ${isManual?'selected':''}>— Enter manually —</option>
            </select>
          </div>` : ''}
          <div style="display:flex;gap:8px;margin-bottom:8px;">
            <div class="field" style="flex:0 0 120px;margin:0">
              <label style="font-size:11px;">Short name${isManual?' <span style="color:#9a6b14;font-weight:700;">(type here)</span>':''}</label>
              <input type="text" data-ti-line-short="${line._id}"
                value="${escapeHtml(isManual ? (line.shortName==='__manual__'?'':line.shortName||'') : line.shortName||'')}"
                placeholder="${isManual?'e.g. EPP':'auto-filled'}"
                style="${isManual?'border-color:#9a6b14;background:#fff8f0;':'background:var(--paper-dim,#f5f5f5);color:var(--ink-soft);'}"
                ${!isManual ? 'readonly' : ''}>
            </div>
            <div class="field" style="flex:1;margin:0"><label style="font-size:11px;">Rate (₹)</label>
              <input type="number" data-ti-line-rate="${line._id}" value="${line.rate||''}" placeholder="0"></div>
            <div class="field" style="flex:0 0 80px;margin:0"><label style="font-size:11px;">Qty</label>
              <input type="number" data-ti-line-qty="${line._id}" value="${line.qty||1}" min="1"></div>
          </div>
          <div class="field" style="margin:0">
            <label style="font-size:11px;">Description${isManual?' <span style="color:#9a6b14;font-weight:700;">(type here)</span>':' (auto-filled, editable)'}</label>
            <input type="text" data-ti-line-desc="${line._id}" value="${escapeHtml(line.desc||'')}"
              placeholder="${isManual?'Enter full description':'Auto-filled from product'}">
          </div>
        </div>`;
    }).join('');

    // Wire all TI line inputs
    wrap.querySelectorAll('[data-ti-del-line]').forEach(b => {
      b.onclick = () => { tiLines = tiLines.filter(l => l._id !== b.dataset.tiDelLine); renderTiLines(); tiPreview(); };
    });
    wrap.querySelectorAll('[data-ti-line-prog]').forEach(sel => {
      sel.onchange = () => {
        const l = tiLines.find(x => x._id === sel.dataset.tiLineProg);
        if (!l) return;
        l.shortName = sel.value;
        if (sel.value && sel.value !== '__manual__') {
          const prod = pfProducts.find(p => p.shortName === sel.value);
          if (prod) { l.rate = prod.rate; l.desc = prod.category === 'OTHER' ? prod.longName : `Sale of ${prod.longName}`; }
        } else { l.desc = ''; l.rate = 0; }
        renderTiLines(); tiPreview();
      };
    });
    wrap.querySelectorAll('[data-ti-line-short]').forEach(inp => {
      inp.oninput = () => { const l = tiLines.find(x => x._id === inp.dataset.tiLineShort); if (l) l.shortName = inp.value || '__manual__'; };
    });
    wrap.querySelectorAll('[data-ti-line-desc]').forEach(inp => {
      inp.oninput = () => { const l = tiLines.find(x => x._id === inp.dataset.tiLineDesc); if (l) { l.desc = inp.value; tiPreview(); } };
    });
    wrap.querySelectorAll('[data-ti-line-rate]').forEach(inp => {
      inp.oninput = () => { const l = tiLines.find(x => x._id === inp.dataset.tiLineRate); if (l) { l.rate = parseFloat(inp.value)||0; tiPreview(); } };
    });
    wrap.querySelectorAll('[data-ti-line-qty]').forEach(inp => {
      inp.oninput = () => { const l = tiLines.find(x => x._id === inp.dataset.tiLineQty); if (l) { l.qty = parseFloat(inp.value)||1; tiPreview(); } };
    });
  }

  const tiPreview = () => {
    const cid = document.getElementById('ti-client').value;
    const cl = clients.find(c => c.id === cid);
    const tds = document.getElementById('ti-tds') ? document.getElementById('ti-tds').value : 'no';
    const taxable = tiLines.reduce((s, l) => {
      if (l._type === 'freight') return s + (l.rate||0);
      return s + ((l.rate||0)*(l.qty||1)) - (l.disc||0);
    }, 0);
    const type = cl ? gstType(cl.state) : 'igst';
    const intra = type === 'intra';
    const cgst = intra ? taxable*0.09 : 0;
    const sgst = intra ? taxable*0.09 : 0;
    const igst = intra ? 0 : taxable*0.18;
    const gross = taxable + cgst + sgst + igst;
    const receivable = invCalcReceivable(taxable, type, tds);
    const gstLabel = intra ? 'CGST 9% + SGST 9%' : 'IGST 18%';
    const tdsNote = tds === 'yes'
      ? ` | TDS 10% = ${fmtMoney(taxable*0.1)} | <strong>Receivable: ${fmtMoney(receivable)}</strong>`
      : ` | <strong>Receivable: ${fmtMoney(receivable)}</strong>`;
    const el = document.getElementById('ti-gst-preview');
    if (el) el.innerHTML = `<b>Preview:</b> Taxable ${fmtMoney(taxable)} + ${gstLabel} = Gross ${fmtMoney(gross)}${tdsNote}`;
    return { taxable, cgst, sgst, igst, gross, type, tdsDeducted: tds, receivable };
  };

  renderTiLines();
  tiPreview();

  document.getElementById('ti-add-line').onclick = () => {
    if (tiProgCount() >= 3) { toast('Maximum 3 program lines (A, B, C)'); return; }
    const fi = tiLines.findIndex(l => l._type === 'freight');
    const nl = { _id: uid(), _type: 'prog', desc: '', shortName: pfProducts.length ? '' : '__manual__', rate: 0, qty: 1, disc: 0 };
    if (fi >= 0) tiLines.splice(fi, 0, nl); else tiLines.push(nl);
    renderTiLines(); tiPreview();
  };
  document.getElementById('ti-add-freight').onclick = () => {
    if (tiHasFreight()) { toast('Freight line already added'); return; }
    tiLines.push({ _id: uid(), _type: 'freight', desc: 'Freight / Courier charges', rate: 0, qty: 1, disc: 0 });
    renderTiLines(); tiPreview();
  };

  const tiNatureBtn = document.getElementById('ti-nature-add');
  if (tiNatureBtn) tiNatureBtn.onclick = () => invAddNature('ti-nature');
  const tiNatureSel = document.getElementById('ti-nature');
  const tiNatureManual = document.getElementById('ti-nature-manual');
  if (tiNatureSel && tiNatureManual) {
    tiNatureSel.onchange = () => {
      tiNatureManual.style.display = tiNatureSel.value === '__other__' ? 'block' : 'none';
      if (tiNatureSel.value === '__other__') tiNatureManual.focus();
    };
  }

  ['ti-client','ti-tds'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', tiPreview);
  });

  document.getElementById('ti-cancel').onclick = () => {
    closeModal();
    if (sourcePi) invOpenTIForm(null, null, false); // back to YES/NO
  };

  document.getElementById('ti-save').onclick = () => {
    const cid = document.getElementById('ti-client').value;
    const cl = clients.find(c => c.id === cid);
    if (!cl) { toast('Select a client'); return; }
    if (!tiLines.length || (!tiLines[0].desc.trim() && tiLines[0]._type !== 'freight')) {
      toast('At least one line item with a description is required'); return;
    }
    const calc = tiPreview();
    const natureRaw = document.getElementById('ti-nature') ? document.getElementById('ti-nature').value : '';
    const nature = natureRaw === '__other__'
      ? (document.getElementById('ti-nature-manual') ? document.getElementById('ti-nature-manual').value.trim() : '')
      : natureRaw;
    const firstLine = tiLines[0];
    const rec = {
      id: existingTI ? existingTI.id : uid(),
      invNo: existingTI ? existingTI.invNo : nextTINumber(),
      date: document.getElementById('ti-date').value,
      supplyDate: document.getElementById('ti-supply').value || document.getElementById('ti-date').value,
      paymentTerms: document.getElementById('ti-terms').value,
      nature, tdsDeducted: calc.tdsDeducted,
      clientId: cid, clientName: cl.companyName, clientShort: cl.shortName,
      lineItems: tiLines.map(l => ({
        _type: l._type || 'prog',
        desc: l.desc, shortName: l.shortName || '',
        rate: l.rate, qty: l._type === 'freight' ? 1 : (l.qty||1), disc: l.disc||0,
      })),
      desc: tiLines.filter(l=>l._type!=='freight').map(l=>l.desc).filter(Boolean).join('; '),
      unit: firstLine.shortName || '', rate: firstLine.rate, qty: firstLine.qty||1,
      sac: document.getElementById('ti-sac').value || '998399',
      disc: tiLines.reduce((s,l)=>s+(l.disc||0),0),
      taxable: calc.taxable, cgst: calc.cgst, sgst: calc.sgst,
      igst: calc.igst, gross: calc.gross, gstType: calc.type,
      receivableAmount: calc.receivable,
      status: existingTI ? existingTI.status : 'draft',
      fromPiId: sourcePi ? sourcePi.id : (existingTI ? existingTI.fromPiId : null),
      fromPiNo: sourcePi ? sourcePi.invNo : (existingTI ? existingTI.fromPiNo : null),
    };
    if (existingTI) {
      const oldRecv = existingTI.receivableAmount || existingTI.gross;
      Object.assign(existingTI, rec);
      if (oldRecv !== rec.receivableAmount) invUpdateReceivableAmount(rec.invNo, rec.receivableAmount);
    } else {
      DB.invoicing.taxInvoices.push(rec);
      if (sourcePi) {
        sourcePi.status = 'converted';
        invRemoveFromReceivables(sourcePi.invNo);
      }
      invAddToReceivables(rec, 'TI');
    }
    DB.invoicing.tiSequence[currentInvFY()] = parseInt(rec.invNo.split('/').pop(), 10);
    saveDB(); closeModal(); INV_TAB = 'ti'; invRenderRegister();
    toast(`${rec.invNo} ${existingTI ? 'updated' : 'generated'}`);
  };
}

// ── Cashflow integration ───────────────────────────────────

// When an invoice amount is edited, update its linked receivable in all months
function invUpdateReceivableAmount(invNo, newGross) {
  let changed = false;
  Object.entries(DB.months).forEach(([mk, m]) => {
    if (!m.receivables) return;
    m.receivables.forEach(r => {
      if (r.name && r.name.startsWith(invNo)) {
        r.amount = newGross;
        changed = true;
      }
    });
  });
  if (changed) {
    // Re-sync carry-forward if today's month was affected
    syncCarriedReceivables();
    toast('Invoice and cashflow receivable updated');
  }
}

function invAddToReceivables(inv, type) {
  const mk = inv.date ? inv.date.slice(0, 7) : todayMonthKey();
  const m = ensureMonthExists(mk);
  const amount = inv.receivableAmount || inv.gross; // use TDS-adjusted amount if set
  m.receivables.push({
    id: uid(),
    name: `${inv.invNo} — ${inv.clientShort || inv.clientName}`,
    amount,
    _invoiceId: inv.id,
    _invoiceType: type,
  });
  if (mk === todayMonthKey()) syncCarriedReceivables();
  saveDB([mk, nextMonthKey(todayMonthKey())]);
}

function invRemoveFromReceivables(invNo) {
  // Remove the receivable created by this invoice number across all months
  Object.values(DB.months).forEach(m => {
    if (m.receivables) {
      m.receivables = m.receivables.filter(r => !r.name || !r.name.startsWith(invNo));
    }
  });
}

function invCancel(invId, type) {
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;

  // Already cancelled — go straight to permanent delete option
  if (inv.status === 'cancelled') {
    const body = `<div class="hint" style="margin-bottom:10px;">
      <strong>${escapeHtml(inv.invNo)}</strong> is already cancelled.<br>
      You can permanently delete it to remove it from the register entirely.
    </div>`;
    openModal(`Delete ${inv.invNo}`, body,
      `<button class="btn" id="icd-close">Close</button>
       <button class="btn btn-danger" id="icd-delete">Permanently delete</button>`);
    document.getElementById('icd-close').onclick = closeModal;
    document.getElementById('icd-delete').onclick = () => {
      if (!confirm(`Permanently delete ${inv.invNo}? This cannot be undone.`)) return;
      if (type === 'pi') DB.invoicing.proformas = DB.invoicing.proformas.filter(p => p.id !== invId);
      else DB.invoicing.taxInvoices = DB.invoicing.taxInvoices.filter(t => t.id !== invId);
      saveDB(); closeModal(); invRenderRegister();
      toast(`${inv.invNo} permanently deleted`);
    };
    return;
  }

  const body = `
    <div class="field">
      <label>What would you like to do with ${escapeHtml(inv.invNo)}?</label>
    </div>
    <div class="hint" style="margin-top:8px;">
      <strong>Cancel</strong> — marks cancelled, removes receivable from cashflow, keeps in register for audit trail.<br><br>
      <strong>Permanently delete</strong> — removes from register entirely. Use only for test entries or duplicates.
    </div>`;

  openModal(`Cancel / Delete ${inv.invNo}`, body,
    `<button class="btn" id="icd-close">Close</button>
     <button class="btn" id="icd-cancel" style="background:#9a6b14; color:#fff; border-color:#9a6b14;">Cancel invoice</button>
     <button class="btn btn-danger" id="icd-delete">Permanently delete</button>`);

  document.getElementById('icd-close').onclick = closeModal;

  document.getElementById('icd-cancel').onclick = () => {
    inv.status = 'cancelled';
    invRemoveFromReceivables(inv.invNo);
    saveDB(); closeModal(); invRenderRegister();
    toast(`${inv.invNo} cancelled — receivable removed from cashflow`);
  };

  document.getElementById('icd-delete').onclick = () => {
    if (!confirm(`Permanently delete ${inv.invNo}? This cannot be undone.`)) return;
    invRemoveFromReceivables(inv.invNo);
    if (type === 'pi') DB.invoicing.proformas = DB.invoicing.proformas.filter(p => p.id !== invId);
    else DB.invoicing.taxInvoices = DB.invoicing.taxInvoices.filter(t => t.id !== invId);
    saveDB(); closeModal(); invRenderRegister();
    toast(`${inv.invNo} permanently deleted`);
  };
}

// ── Email (mailto) ─────────────────────────────────────────
// ── Excel generation ───────────────────────────────────────
async function invDownloadWord(invId, type) {
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;
  const cl = DB.invoicing.clients.find(c => c.id === inv.clientId) || {};
  toast('Generating Word document…');
  try {
    const docxLib = window.docx;
    if (!docxLib) { toast('Word library not loaded — try refreshing the page'); return; }
    const isPi = type === 'pi';
    const doc = await invBuildWordDoc(inv, cl, isPi, docxLib);
    // Use toBlob (browser-compatible) instead of toBuffer (Node-only)
    const blob = await docxLib.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${inv.invNo.replace(/\//g, '-')}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    toast(`${inv.invNo}.docx downloaded`);
  } catch (e) {
    console.error('Word generation error:', e);
    toast('Word generation failed — ' + (e.message || String(e)));
  }
}

async function invBuildWordDoc(inv, cl, isPi, D) {
  // D = window.docx (UMD bundle)
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, ImageRun } = D;

  const NAVY = '0F2540', GOLD = 'B8860B', INK = '1B2430', WHITE = 'FFFFFF', LIGHT = 'F2F4F6';
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
  const thickBorder = { style: BorderStyle.SINGLE, size: 12, color: NAVY };
  const intra = inv.gstType === 'intra';
  const title = isPi ? 'PROFORMA INVOICE' : 'TAX INVOICE';
  const invLabel = isPi ? 'Proforma Invoice No :' : 'Tax Invoice No :';
  const dateLabel = isPi ? 'Proforma Invoice Date :' : 'Tax Invoice Date :';

  function fmtDate(d) {
    if (!d) return '';
    const parts = d.split('-');
    if (parts.length === 3) return `${parts[1]}/${parts[2]}/${parts[0]}`;
    return d;
  }

  // Landscape column widths (total = 15238 DXA = A4 landscape usable)
  const COLS = [400, 2700, 700, 600, 850, 1400, 1150, 1500, 380, 950, 380, 950, 380, 950, 1948];
  const TOTAL = COLS.reduce((a, b) => a + b, 0);

  function run2(text, opts = {}) {
    return new TextRun({ text: String(text ?? ''), font: 'Calibri',
      size: opts.size || 19, bold: opts.bold || false,
      color: opts.color || INK, italics: opts.italics || false });
  }

  function para(children, align = AlignmentType.LEFT) {
    let runs;
    if (Array.isArray(children)) {
      runs = children;
    } else if (typeof children === 'string') {
      runs = [run2(children)];
    } else {
      // TextRun object passed directly
      runs = [children];
    }
    return new Paragraph({ alignment: align, spacing: { before: 40, after: 40 }, children: runs });
  }

  function money2(n) {
    return Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function pct2(n) { return n ? (Number(n) * 100).toFixed(0) + '%' : ''; }

  function tc(children, w, opts = {}) {
    const { fill, borders: b, align = AlignmentType.LEFT, colspan = 1, rowspan = 1 } = opts;
    const paras = Array.isArray(children) ? children : [para(children, align)];
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      columnSpan: colspan, rowSpan: rowspan,
      verticalAlign: VerticalAlign.CENTER,
      shading: fill ? { fill, type: ShadingType.CLEAR, color: fill } : undefined,
      borders: b || { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder },
      margins: { top: 50, bottom: 50, left: 80, right: 80 },
      children: paras
    });
  }

  function hdrCell(text, w, cs = 1) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA }, columnSpan: cs,
      verticalAlign: VerticalAlign.CENTER,
      shading: { fill: NAVY, type: ShadingType.CLEAR, color: NAVY },
      borders: { top: noBorder, bottom: thinBorder, left: noBorder, right: thinBorder },
      margins: { top: 60, bottom: 60, left: 60, right: 60 },
      children: [para(run2(text, { bold: true, size: 17, color: WHITE }), AlignmentType.CENTER)]
    });
  }

  function amtTC(val, w, bold = false, fill = null) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: fill ? { fill, type: ShadingType.CLEAR, color: fill } : undefined,
      borders: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder },
      margins: { top: 50, bottom: 50, left: 60, right: 60 },
      children: [para(run2(val, { bold, size: 18 }), AlignmentType.RIGHT)]
    });
  }

  const emptyRow = () => new TableRow({ children: [new TableCell({
    columnSpan: 15, width: { size: TOTAL, type: WidthType.DXA },
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
    children: [new Paragraph({ spacing: { before: 60, after: 60 }, children: [] })]
  })]});

  // ── Title ──
  const titleRow = new TableRow({ children: [new TableCell({
    columnSpan: 15, width: { size: TOTAL, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR, color: NAVY },
    borders: { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [para(run2(title, { bold: true, size: 28, color: WHITE }), AlignmentType.CENTER)]
  })]});

  // ── Line item rows ──
  const lineHdr1 = new TableRow({ children:[
    hdrCell('SR',COLS[0]), hdrCell('DESCRIPTION',COLS[1]),
    hdrCell('SAC',COLS[2]), hdrCell('QTY',COLS[3]),
    hdrCell('RATE',COLS[4]), hdrCell('NET AMOUNT',COLS[5]),
    hdrCell('DISC',COLS[6]), hdrCell('TAXABLE VALUE',COLS[7]),
    hdrCell('CGST %',COLS[8]), hdrCell('CGST AMT',COLS[9]),
    hdrCell('SGST %',COLS[10]), hdrCell('SGST AMT',COLS[11]),
    hdrCell('IGST %',COLS[12]), hdrCell('IGST AMT',COLS[13]),
    hdrCell('GROSS AMOUNT',COLS[14]),
  ]});

  function navyEmpty(w) {
    return new TableCell({ width:{size:w,type:WidthType.DXA},
      shading:{fill:NAVY,type:ShadingType.CLEAR,color:NAVY},
      borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:thinBorder},
      children:[new Paragraph({children:[]})] });
  }

  // lineHdr2 removed — % merged into row 1 headers

  function lineRow2(sr,desc,sac,unit,rate,netAmt,disc,taxable,
                   cp,ca,sp,sa,ip,ia,gross,fill=null) {
    const nb = {top:thinBorder,bottom:thinBorder,left:thinBorder,right:thinBorder};
    function c(v,w,right=false,b=false) {
      return new TableCell({ width:{size:w,type:WidthType.DXA},
        shading:fill?{fill,type:ShadingType.CLEAR,color:fill}:undefined,
        borders:nb, margins:{top:50,bottom:50,left:60,right:60},
        children:[para(run2(v,{size:18,bold:b||!!fill}),
          right?AlignmentType.RIGHT:AlignmentType.LEFT)] });
    }
    return new TableRow({ children:[
      c(sr,COLS[0]), c(desc,COLS[1]), c(sac,COLS[2]), c(unit,COLS[3]),
      c(typeof rate==='number'?money2(rate):rate, COLS[4], true),
      c(typeof netAmt==='number'?money2(netAmt):netAmt||'', COLS[5], true),
      c(typeof disc==='number'?money2(disc):disc||'', COLS[6], true),
      c(typeof taxable==='number'?money2(taxable):taxable||'', COLS[7], true),
      c(pct2(cp),COLS[8],true), c(ca?money2(ca):'',COLS[9],true),
      c(pct2(sp),COLS[10],true), c(sa?money2(sa):'',COLS[11],true),
      c(pct2(ip),COLS[12],true), c(ia?money2(ia):'',COLS[13],true),
      c(typeof gross==='number'?money2(gross):gross||'',COLS[14],true),
    ]});
  }

  const net = inv.rate * (inv.qty || 1);

  // Build line item rows from lineItems array (multi-item support) or fallback to single item
  const lineItemsData = (inv.lineItems && inv.lineItems.length)
    ? inv.lineItems
    : [{ desc: inv.desc, qty: inv.qty||1, rate: inv.rate||0, disc: inv.disc||0 }];

  const lineRows = lineItemsData.map((item, i) => {
    const isFreight = item._type === 'freight';
    const itemNet = isFreight ? (item.rate||0) : (item.rate||0) * (item.qty||1);
    const itemTaxable = itemNet - (item.disc||0);
    const itemCgst = intra ? itemTaxable * 0.09 : 0;
    const itemSgst = intra ? itemTaxable * 0.09 : 0;
    const itemIgst = intra ? 0 : itemTaxable * 0.18;
    const itemGross = itemTaxable + itemCgst + itemSgst + itemIgst;
    const qtyDisplay = isFreight ? '' : String(item.qty||1);
    return lineRow2(
      String.fromCharCode(65 + i), item.desc||'', inv.sac||'998399',
      qtyDisplay, item.rate||0, itemNet, item.disc||0, itemTaxable,
      intra?0.09:0, itemCgst, intra?0.09:0, itemSgst,
      intra?0:0.18, itemIgst, itemGross
    );
  });

  // Totals from all lines
  const totalTaxable2 = lineItemsData.reduce((s,item) => {
    const isFreight = item._type === 'freight';
    return s + (isFreight ? (item.rate||0) : ((item.rate||0)*(item.qty||1))) - (item.disc||0);
  }, 0);
  const totalCgst2 = intra ? totalTaxable2 * 0.09 : 0;
  const totalSgst2 = intra ? totalTaxable2 * 0.09 : 0;
  const totalIgst2 = intra ? 0 : totalTaxable2 * 0.18;
  const totalGross2 = totalTaxable2 + totalCgst2 + totalSgst2 + totalIgst2;

  const totalRow2 = lineRow2('TOTAL','','','','','','',(totalTaxable2||0),
    0,totalCgst2||0, 0,totalSgst2||0, 0,totalIgst2||0, totalGross2||0, LIGHT);

  // Words row merged as bottom of line items table
  const wordsRow2 = new TableRow({ children:[new TableCell({
    columnSpan:15, width:{size:TOTAL,type:WidthType.DXA},
    borders:{top:thinBorder,bottom:thinBorder,left:thinBorder,right:thinBorder},
    margins:{top:30,bottom:30,left:100,right:100},
    children:[para([run2('Gross Amount in Words : ',{bold:true,size:17}),
      run2(numberToWords(Math.round(totalGross2||0)),{size:17,italics:true})],AlignmentType.LEFT,{before:15,after:15})]
  })]});

  function wideRow2(text, bold=false) {
    return new TableRow({ children:[new TableCell({
      columnSpan:15, width:{size:TOTAL,type:WidthType.DXA},
      borders:{top:noBorder,bottom:noBorder,left:thinBorder,right:thinBorder},
      margins:{top:14,bottom:14,left:100,right:100},
      children:[para(run2(text,{bold,size:15}),AlignmentType.LEFT,{before:3,after:3})]
    })]});
  }

  function splitRow2(leftTxt, rightTxt) {
    const lw = Math.floor(TOTAL * 0.65);
    const rw = TOTAL - lw;
    return new TableRow({ children:[
      new TableCell({ width:{size:lw,type:WidthType.DXA},
        borders:{top:thinBorder,bottom:noBorder,left:thinBorder,right:noBorder},
        margins:{top:14,bottom:14,left:100,right:80},
        children:[para(run2(leftTxt,{size:15}),AlignmentType.LEFT,{before:5,after:5})] }),
      new TableCell({ width:{size:rw,type:WidthType.DXA},
        borders:{top:thinBorder,bottom:noBorder,left:noBorder,right:thinBorder},
        margins:{top:14,bottom:14,left:80,right:80},
        children:[para(run2(rightTxt,{bold:true,size:15,color:NAVY}),AlignmentType.RIGHT,{before:5,after:5})] }),
    ]});
  }

  const sigRow2 = new TableRow({ children:[
    new TableCell({ width:{size:Math.floor(TOTAL*0.65),type:WidthType.DXA},
      borders:{top:noBorder,bottom:thinBorder,left:thinBorder,right:noBorder},
      margins:{top:20,bottom:20,left:100,right:80},
      children:[
        para(run2(''),AlignmentType.LEFT,{before:0,after:0}),
        para(run2(''),AlignmentType.LEFT,{before:0,after:0}),
      ] }),
    new TableCell({ width:{size:TOTAL-Math.floor(TOTAL*0.65),type:WidthType.DXA},
      borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:thinBorder},
      margins:{top:20,bottom:20,left:80,right:80},
      children:[
        para(run2(''),AlignmentType.RIGHT,{before:0,after:0}),
        para(run2('AUTHORISED SIGNATORY',{bold:true,size:16,color:NAVY}),AlignmentType.RIGHT,{before:5,after:5}),
      ] }),
  ]});

  // ── Table 1: Title ──────────────────────────────────────────
  const titleTable = new Table({
    width:{size:TOTAL,type:WidthType.DXA}, columnWidths:[TOTAL],
    rows:[titleRow]
  });

  // ── Table 2: Invoice meta (No + Date) ───────────────────────
  const metaW = Math.floor(TOTAL/4);
  const metaTable = new Table({
    width:{size:TOTAL,type:WidthType.DXA},
    columnWidths:[metaW, metaW, metaW, TOTAL - metaW*3],
    rows:[new TableRow({ children:[
      tc(run2(invLabel,{bold:true,size:17}), metaW,
        {borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:noBorder},top:20,bottom:20,left:60,right:60}),
      tc(run2(inv.invNo||'',{bold:true,size:17,color:'CC0000'}), metaW,
        {borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:noBorder},top:20,bottom:20,left:60,right:60}),
      tc(run2(dateLabel,{bold:true,size:17}), metaW,
        {borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:noBorder},top:20,bottom:20,left:60,right:60}),
      tc(run2(fmtDate(inv.date||''),{size:17}), TOTAL-metaW*3,
        {borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:noBorder},top:20,bottom:20,left:60,right:60}),
    ]})]
  });

  // ── Table 3: From / To ──────────────────────────────────────
  const fromW = Math.floor(TOTAL * 0.40);
  const toW   = TOTAL - fromW;

  function addrRow2(fromLabel, fromVal, toLabel, toVal, fromBold=false, toBold=false) {
    const bFrom = {top:noBorder,bottom:noBorder,left:thinBorder,right:noBorder};
    const bTo   = {top:noBorder,bottom:noBorder,left:thinBorder,right:thinBorder};
    return new TableRow({ children:[
      new TableCell({ width:{size:fromW,type:WidthType.DXA},
        borders:bFrom, margins:{top:6,bottom:6,left:80,right:40},
        children:[para([run2(fromLabel,{bold:true,size:16}),run2(' ',{size:16}),run2(fromVal||'',{bold:fromBold,size:16})],AlignmentType.LEFT,{before:3,after:3})]
      }),
      new TableCell({ width:{size:toW,type:WidthType.DXA},
        borders:bTo, margins:{top:6,bottom:6,left:80,right:40},
        children:[para([run2(toLabel,{bold:true,size:16}),run2(' ',{size:16}),run2(toVal||'',{bold:toBold,size:16})],AlignmentType.LEFT,{before:3,after:3})]
      }),
    ]});
  }

  const fromToTable = new Table({
    width:{size:TOTAL,type:WidthType.DXA}, columnWidths:[fromW,toW],
    rows:[
      new TableRow({ children:[
        new TableCell({ width:{size:fromW,type:WidthType.DXA},
          borders:{top:thinBorder,bottom:noBorder,left:thinBorder,right:noBorder},
          margins:{top:8,bottom:4,left:80,right:40},
          children:[para(run2('From,',{bold:true,size:17}),AlignmentType.LEFT,{before:3,after:3})]
        }),
        new TableCell({ width:{size:toW,type:WidthType.DXA},
          borders:{top:thinBorder,bottom:noBorder,left:thinBorder,right:thinBorder},
          margins:{top:8,bottom:4,left:80,right:40},
          children:[para(run2('To,',{bold:true,size:17}),AlignmentType.LEFT,{before:3,after:3})]
        }),
      ]}),
      addrRow2('',GORU.name,'',cl.companyName||'',true,true),
      addrRow2('',GORU.addr1,'',cl.addr1||''),
      addrRow2('',GORU.addr2,'', (cl.addr2||'')+(cl.addr3?', '+cl.addr3:'')),
      addrRow2('State: '+GORU.state,'','State: '+(cl.state||''),''),
      addrRow2('GSTIN: '+GORU.gstin+'  |  PAN: '+GORU.pan+'  |  TAN: '+GORU.tan,'',
               'GSTIN: '+(cl.gstin||'')+'  |  PAN: '+(cl.pan||'')+'  |  TAN: '+(cl.tan||''),''),
      addrRow2('','','Kind Attn: '+(cl.attn||'')+'  |  '+(cl.email||''),''),
      addrRow2('','','Mob: '+(cl.mobile||''),''),
      // TI: Date of Supply
      ...(isPi ? [] : [addrRow2('','','Date of Supply: '+fmtDate(inv.supplyDate||''),'')]),
      new TableRow({ children:[
        new TableCell({ width:{size:fromW,type:WidthType.DXA},
          borders:{top:noBorder,bottom:thinBorder,left:thinBorder,right:noBorder},
          margins:{top:4,bottom:8,left:80,right:40},
          children:[para(run2('Place of Supply: '+GORU.state,{size:16}),AlignmentType.LEFT,{before:3,after:3})]
        }),
        new TableCell({ width:{size:toW,type:WidthType.DXA},
          borders:{top:noBorder,bottom:thinBorder,left:thinBorder,right:thinBorder},
          margins:{top:4,bottom:8,left:80,right:40},
          children:[para(run2('Supply Destination: '+(cl.state||''),{size:16}),AlignmentType.LEFT,{before:3,after:3})]
        }),
      ]}),
    ]
  });

  // ── Table 4: Line items ─────────────────────────────────────
  const lineTable = new Table({
    width:{size:TOTAL,type:WidthType.DXA}, columnWidths:COLS,
    rows:[lineHdr1, ...lineRows, totalRow2, wordsRow2]
  });

  // ── Table 5: Payment terms / Bank / Signatory ───────────────
  const termsSigW1 = 9904;
  const termsSigW2 = 5334;

  // Stamp image (TI only) — embedded base64
  const STAMP_B64 = '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';
  const stampBytes = STAMP_B64 !== 'STAMP_PLACEHOLDER'
    ? Uint8Array.from(atob(STAMP_B64), c => c.charCodeAt(0)) : null;

  const bottomTable = new Table({
    width:{size:TOTAL,type:WidthType.DXA}, columnWidths:[termsSigW1,termsSigW2],
    rows:[new TableRow({ children:[
      // Left: payment terms + bank
      new TableCell({ width:{size:termsSigW1,type:WidthType.DXA},
        borders:{top:thinBorder,bottom:thinBorder,left:thinBorder,right:noBorder},
        margins:{top:14,bottom:14,left:100,right:80},
        children:[
          para([run2('Payment Terms : ',{bold:true,size:16}),run2(inv.paymentTerms||'Advance',{size:16})],AlignmentType.LEFT,{before:5,after:4}),
          para(run2('Bank transfer/cheque in favour of "Goru Training Pvt. Ltd." payable at Mumbai',{size:15}),AlignmentType.LEFT,{before:3,after:3}),
          para(run2('Bank: '+GORU.bank+', '+GORU.branch+'   A/c: '+GORU.acno+'   IFSC: '+GORU.ifsc,{size:15}),AlignmentType.LEFT,{before:3,after:3}),
          // PI only: computer generated notice; TI has physical stamp instead
          ...(isPi ? [
            para(run2('This is a computer generated invoice  |  contact@lmi-india.in  |  Tel: 022 66364393',{size:14,italics:true,color:'5B6470'}),AlignmentType.LEFT,{before:3,after:5}),
          ] : []),
        ]
      }),
      // Right: FOR GORU + stamp (TI) or spacing (PI) + AUTHORISED SIGNATORY
      new TableCell({ width:{size:termsSigW2,type:WidthType.DXA},
        borders:{top:thinBorder,bottom:thinBorder,left:thinBorder,right:thinBorder},
        margins:{top:6,bottom:6,left:60,right:60},
        verticalAlign:VerticalAlign.TOP,
        children:[
          para(run2('FOR GORU TRAINING PRIVATE LIMITED',{bold:true,size:16,color:NAVY}),AlignmentType.RIGHT,{before:5,after:4}),
          // TI: stamp centred between the two text lines
          ...(!isPi && stampBytes ? [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing:{before:2,after:2},
              children:[new ImageRun({data:stampBytes,transformation:{width:60,height:60},type:'jpg'})],
            }),
          ] : [
            para('',AlignmentType.RIGHT,{before:0,after:24}),
          ]),
          para(run2('AUTHORISED SIGNATORY',{bold:true,size:16,color:NAVY}),AlignmentType.RIGHT,{before:2,after:3}),
        ]
      }),
    ]})]
  });

  const { Header, Footer } = D;

  // ── Header: company letterhead ──
  // ── Header: LMI India letterhead (matching uploaded sample) ──
  const LMI_BLUE = '1F5FA6';
  // LMI South Asia logo embedded as base64 — auto-appears in every generated invoice header
  const LMI_LOGO_B64 = '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';
  const logoBytes = Uint8Array.from(atob(LMI_LOGO_B64), c => c.charCodeAt(0));
  const docHeader = new Header({
    children: [
      new Table({
        width: { size: TOTAL, type: WidthType.DXA },
        columnWidths: [1800, TOTAL - 1800],
        rows: [new TableRow({ children: [

          // ── Left: Logo placeholder ──
          new TableCell({
            width: { size: 1800, type: WidthType.DXA },
            borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 6, color: LMI_BLUE }, left: noBorder, right: { style: BorderStyle.SINGLE, size: 6, color: LMI_BLUE } },
            margins: { top: 60, bottom: 80, left: 60, right: 60 },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 40, after: 40 },
                children: [
                  new ImageRun({
                    data: logoBytes,
                    transformation: { width: 110, height: 90 },
                    type: 'jpg',
                  })
                ]
              }),
            ]
          }),

          // ── Right: LMI India branding + address ──
          new TableCell({
            width: { size: TOTAL - 1800, type: WidthType.DXA },
            borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 6, color: LMI_BLUE }, left: noBorder, right: noBorder },
            margins: { top: 40, bottom: 80, left: 120, right: 60 },
            children: [
              // "LMI INDIA  -  A Division of Goru Training Pvt. Ltd."
              new Paragraph({
                spacing: { before: 40, after: 20 },
                children: [
                  run2('LMI SOUTH ASIA', { bold: true, size: 30, color: LMI_BLUE }),
                  run2('  -  ', { size: 22, color: LMI_BLUE }),
                  run2('A Division of Goru Training Pvt. Ltd.', { bold: true, italics: true, size: 22, color: LMI_BLUE }),
                ]
              }),
              // CIN
              new Paragraph({
                spacing: { before: 0, after: 60 },
                children: [run2('CIN: U74120MH2016PTC273207', { size: 16, color: INK })]
              }),
              // Business Address (with underline via paragraph bottom border)
              new Paragraph({
                spacing: { before: 0, after: 20 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: INK, space: 2 } },
                children: [
                  run2('Business Address: ', { bold: true, size: 17, color: INK }),
                  run2('Apeejay House, 2', { size: 17, color: INK }),
                  new TextRun({ text: 'nd', font: 'Calibri', size: 17, color: INK, superScript: true }),
                  run2(' flr, 36-37 W, 3 Dinshaw Vachha Road, Mumbai - 400020', { size: 17, color: INK }),
                ]
              }),
              // Email / Website / Tel
              new Paragraph({
                spacing: { before: 40, after: 0 },
                children: [
                  run2('Email: ', { bold: true, size: 16, color: INK }),
                  run2('contact@lmi-india.in', { size: 16, color: LMI_BLUE }),
                  run2('   Website: ', { bold: true, size: 16, color: INK }),
                  run2('www.lmi-india.in', { size: 16, color: LMI_BLUE }),
                  run2('   Tel: 022 66364393', { size: 16, color: INK }),
                ]
              }),
            ]
          }),

        ]})]
      })
    ]
  });

  // ── Footer: computer generated disclaimer ──
  const docFooter = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA', space: 4 } },
        spacing: { before: 60 },
        children: [
          run2('This is a computer generated invoice.   ', { size: 16, italics: true, color: '5B6470' }),
          run2('LMI India · Goru Training Pvt. Ltd. · contact@lmi-india.in · Tel: 022 66364393', { size: 16, color: '5B6470' }),
        ]
      })
    ]
  });

  const { PageOrientation } = D;

  const doc = new Document({
    sections:[{
      properties:{ page:{ size:{
        width: 11906,    // docx library swaps internally — give portrait values + orient=landscape
        height: 16838,   // result: w:w="16838" h:h="11906" orient="landscape" in XML
        orientation: PageOrientation ? PageOrientation.LANDSCAPE : 'landscape',
      },
        margin:{top:600,right:700,bottom:600,left:700} } },
      headers: { default: docHeader },
      children:[
        titleTable,
        new Paragraph({spacing:{before:0,after:0,line:240,lineRule:'exact'},children:[]}),
        metaTable,
        new Paragraph({spacing:{before:0,after:0,line:240,lineRule:'exact'},children:[]}),
        fromToTable,
        new Paragraph({spacing:{before:0,after:0,line:240,lineRule:'exact'},children:[]}),
        lineTable,
        new Paragraph({spacing:{before:0,after:0,line:240,lineRule:'exact'},children:[]}),
        bottomTable,
      ]
    }]
  });

  return doc;
}

function invBuildSheetData(inv, cl, isPi) {
  const title = isPi ? 'PROFORMA INVOICE' : 'TAX INVOICE';
  const numLabel = isPi ? 'Proforma Invoice No :-' : 'Tax Invoice No :-';
  const dateLabel = isPi ? 'Proforma Invoice Date :-' : 'Tax Invoice Date :-';
  const intra = inv.gstType === 'intra';
  const amtWords = numberToWords(Math.round(inv.gross));

  // Client address block
  const addrLines = [cl.companyName||'', cl.addr1||'', cl.addr2||'', cl.addr3||''].filter(Boolean);

  return [
    ['','','','','','','','','','','','','','',''],        // row 1
    ['','','','','','','','','','','','','','',''],        // row 2
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['',title,'','','','','','','','','','','','',''],    // row 9: title
    ['','','','','','','','','','','','','','',''],
    ['From,','','','','','To,','','','','','',numLabel,'','',inv.invNo],
    [GORU.name,'','','','',addrLines[0]||'','','','','','',dateLabel,'','',inv.date],
    ['','','','','',addrLines[1]||'','','','','','','Date of Supply :-','','',inv.supplyDate],
    ['','','','','',addrLines[2]||'','','','','','','Place of Supply :-','','',cl.state||''],
    ['State :-','Maharashtra','','','','State :-','','','','','','Kind Attn:','','',cl.attn||''],
    ['GSTIN :-',GORU.gstin,'','','','GSTIN :-',cl.gstin||'','','','','','Mob :-','','',cl.mobile||''],
    ['PAN :-',GORU.pan,'','','','PAN :-',cl.pan||'','','','','','Email:-','','',cl.email||''],
    ['TAN :-',GORU.tan,'','','','TAN :-',cl.tan||'','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['SR','DESCRIPTION','HNS/SAC CODE','UNIT','RATE','NET AMOUNT','DISC','TAXABLE VALUE','CGST','','SGST','','IGST','','GROSS AMOUNT'],
    ['','','','','','','','','%','AMOUNT','%','AMOUNT','%','AMOUNT',''],
    ['A',inv.desc,inv.sac,inv.unit,inv.rate,inv.rate*(inv.qty||1),inv.disc,inv.taxable,
      intra?0.09:0, intra?(inv.taxable*0.09):0,
      intra?0.09:0, intra?(inv.taxable*0.09):0,
      intra?0:0.18, intra?0:(inv.taxable*0.18),
      inv.gross - (inv.freight||0) - (inv.other||0)],
    ['','','','','','','','','','','','','','',''],
    ['B','Freight charges','','','','','',(inv.freight||0), 0,0, 0,0, intra?0:0.18, intra?0:((inv.freight||0)*0.18), (inv.freight||0)*(intra?1:1.18)],
    ['C','Other Taxable Charges','','','','','','', inv.other||0, 0,0, 0,0, intra?0:0.18, intra?0:((inv.other||0)*0.18)],
    ['TOTAL','','','','','','',(inv.taxable||0), 0,(inv.cgst||0), 0,(inv.sgst||0), 0,(inv.igst||0),(inv.gross||0)],
    ['','','','','','','','','','','','','','',''],
    [`Gross Amount in Words :- ${amtWords}`,'','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','',''],
    [`Payment Terms :  ${inv.paymentTerms||''}`, '','','','','','','','','','','','','','FOR GORU TRAINING PRIVATE LIMITED'],
    [`All payments by bank transfer/draft/cheque payable at Mumbai in favour of "Goru Training Pvt. Ltd."`],
    [`Bank Name: ${GORU.bank}    Branch: ${GORU.branch}`],
    [`A/c. No.: ${GORU.acno};   IFSC Code: ${GORU.ifsc};`],
    ['','','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','AUTHORISED SIGNATORY',''],
  ];
}

// Simple number to words for Indian currency
function numberToWords(n) {
  if (n === 0) return 'Zero Rupees Only';
  const ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
    'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  function w(num) {
    if (num < 20) return ones[num];
    if (num < 100) return tens[Math.floor(num/10)] + (num%10 ? ' ' + ones[num%10] : '');
    if (num < 1000) return ones[Math.floor(num/100)] + ' Hundred' + (num%100 ? ' ' + w(num%100) : '');
    return '';
  }
  const cr = Math.floor(n / 10000000); n %= 10000000;
  const lac = Math.floor(n / 100000); n %= 100000;
  const th = Math.floor(n / 1000); n %= 1000;
  const parts = [];
  if (cr) parts.push(w(cr) + ' Crore');
  if (lac) parts.push(w(lac) + ' Lakh');
  if (th) parts.push(w(th) + ' Thousand');
  if (n) parts.push(w(n));
  return parts.join(' ') + ' Rupees Only';
}

// ── Report generator ───────────────────────────────────────
function invOpenReport() {
  invInit();
  const fy = currentInvFY();
  const body = `
    <div class="field-row">
      <div class="field"><label>Report type</label>
        <select id="rpt-type">
          <option value="all">All invoices</option>
          <option value="pi">Proforma invoices only</option>
          <option value="ti">Tax invoices only</option>
        </select>
      </div>
      <div class="field"><label>Group by</label>
        <select id="rpt-group">
          <option value="none">None (flat list)</option>
          <option value="client">Client</option>
          <option value="month">Month</option>
          <option value="quarter">Quarter (AMJ = Q1)</option>
        </select>
      </div>
    </div>
    <div class="field-row">
      <div class="field"><label>From date</label><input type="date" id="rpt-from"></div>
      <div class="field"><label>To date</label><input type="date" id="rpt-to"></div>
    </div>
    <div class="field"><label>Client (leave blank for all)</label>
      <select id="rpt-client">
        <option value="">All clients</option>
        ${DB.invoicing.clients.map(c => `<option value="${c.id}">${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</option>`).join('')}
      </select>
    </div>`;
  openModal('Generate report', body,
    `<button class="btn" id="rpt-cancel">Cancel</button><button class="btn btn-primary" id="rpt-run">Download Excel</button>`);
  document.getElementById('rpt-cancel').onclick = closeModal;
  document.getElementById('rpt-run').onclick = invRunReport;
}

function invRunReport() {
  invInit();
  const type = document.getElementById('rpt-type').value;
  const group = document.getElementById('rpt-group').value;
  const from = document.getElementById('rpt-from').value;
  const to = document.getElementById('rpt-to').value;
  const clientFilter = document.getElementById('rpt-client').value;

  let list = [];
  if (type !== 'ti') list = list.concat(DB.invoicing.proformas.map(i => ({ ...i, _type: 'PI' })));
  if (type !== 'pi') list = list.concat(DB.invoicing.taxInvoices.map(i => ({ ...i, _type: 'TI' })));

  // Filters
  if (from) list = list.filter(i => i.date >= from);
  if (to) list = list.filter(i => i.date <= to);
  if (clientFilter) list = list.filter(i => i.clientId === clientFilter);
  list = list.filter(i => i.status !== 'cancelled');
  list.sort((a, b) => (a.date || '').localeCompare(b.date || ''));

  const XLSX = window.XLSX;
  if (!XLSX) { toast('Excel library not loaded'); return; }

  const quarterLabel = date => {
    if (!date) return '';
    const m = parseInt(date.slice(5, 7));
    if ([4,5,6].includes(m)) return 'Q1 (AMJ)';
    if ([7,8,9].includes(m)) return 'Q2 (JAS)';
    if ([10,11,12].includes(m)) return 'Q3 (OND)';
    return 'Q4 (JFM)';
  };
  const monthLabel2 = date => date ? date.slice(0, 7) : '';

  const headers = ['Type','Invoice No','Date','Client','Description','Taxable','CGST','SGST','IGST','Gross','Status'];
  if (group !== 'none') headers.unshift('Group');

  const rows = [headers];
  let groupKey = '';
  let groupTotal = 0;

  const addRow = (i, gk) => {
    const row = [i._type, i.invNo, i.date, i.clientShort||i.clientName, i.desc,
      i.taxable, i.cgst||0, i.sgst||0, i.igst||0, i.gross, i.status];
    if (group !== 'none') row.unshift(gk);
    rows.push(row);
    groupTotal += i.gross || 0;
  };

  list.forEach(i => {
    let gk = '';
    if (group === 'client') gk = i.clientShort || i.clientName;
    else if (group === 'month') gk = monthLabel2(i.date);
    else if (group === 'quarter') gk = quarterLabel(i.date);

    if (group !== 'none' && gk !== groupKey) {
      if (groupKey) {
        const totalRow = new Array(headers.length).fill('');
        totalRow[0] = `${groupKey} TOTAL`;
        totalRow[headers.length - 2] = groupTotal;
        rows.push(totalRow);
        rows.push([]);
      }
      groupKey = gk; groupTotal = 0;
    }
    addRow(i, gk);
  });

  if (group !== 'none' && groupKey) {
    const totalRow = new Array(headers.length).fill('');
    totalRow[0] = `${groupKey} TOTAL`;
    totalRow[headers.length - 2] = groupTotal;
    rows.push(totalRow);
  }

  // Grand total
  const grandTotal = list.reduce((s, i) => s + (i.gross || 0), 0);
  rows.push([]);
  const gtRow = new Array(headers.length).fill('');
  gtRow[0] = 'GRAND TOTAL'; gtRow[headers.length - 2] = grandTotal;
  rows.push(gtRow);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = headers.map((h, i) => ({ wch: [8,14,10,20,30,12,8,8,8,12,10][i] || 12 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Invoice Report');
  const fname = `Invoice_Report_${from||'all'}_to_${to||'all'}.xlsx`;
  XLSX.writeFile(wb, fname);
  closeModal();
  toast(`Report downloaded: ${list.length} invoice${list.length !== 1 ? 's' : ''}`);
}

/* ===========================================================
   TEST MODE
   Snapshots the entire DB before any test activity. UNDO TEST
   restores the snapshot exactly — all changes during test mode
   (invoices, payments, receivables, cashflow edits) are wiped.
   =========================================================== */

let TEST_SNAPSHOT = null; // stringified DB snapshot
let TEST_MODE_ACTIVE = false;

function activateTestMode() {
  if (TEST_MODE_ACTIVE) {
    toast('Test mode is already active');
    return;
  }
  if (!confirm('Activate Test Mode?\n\nA snapshot of all live data will be taken now. You can make any changes freely — clicking "UNDO TEST & RESTORE" will bring everything back to exactly this point.\n\nYour live data is safe.')) return;

  TEST_SNAPSHOT = JSON.stringify(DB);
  TEST_MODE_ACTIVE = true;
  localStorage.setItem('lmi_test_snapshot', TEST_SNAPSHOT);
  localStorage.setItem('lmi_test_mode', '1');

  document.getElementById('testModeBanner').style.display = 'block';
  document.getElementById('btnTestMode').style.background = '#a4302a';
  document.getElementById('btnTestMode').style.color = '#fff';
  document.getElementById('btnTestMode').textContent = 'TEST ●';
  toast('Test mode active — your live data is safely snapshotted');
}

function undoTestMode() {
  if (!TEST_MODE_ACTIVE) {
    toast('Test mode is not active');
    return;
  }
  if (!TEST_SNAPSHOT) {
    toast('No snapshot found — cannot undo');
    return;
  }
  if (!confirm('Undo test mode?\n\nThis will restore ALL data to the state it was in when Test Mode was activated. Every change made during the test session (invoices, payments, edits) will be permanently removed.\n\nAre you sure?')) return;

  try {
    DB = JSON.parse(TEST_SNAPSHOT);
    saveDB();
    TEST_MODE_ACTIVE = false;
    TEST_SNAPSHOT = null;
    localStorage.removeItem('lmi_test_snapshot');
    localStorage.removeItem('lmi_test_mode');
    document.getElementById('testModeBanner').style.display = 'none';
    document.getElementById('btnTestMode').style.background = '';
    document.getElementById('btnTestMode').style.color = '';
    document.getElementById('btnTestMode').textContent = 'TEST';
    renderAll();
    toast('✓ Live data restored — test session undone');
  } catch (e) {
    toast('Could not restore snapshot: ' + e.message);
  }
}

function restoreTestModeIfActive() {
  // On page load, check if test mode was active before a refresh
  if (localStorage.getItem('lmi_test_mode') === '1') {
    TEST_SNAPSHOT = localStorage.getItem('lmi_test_snapshot');
    TEST_MODE_ACTIVE = true;
    setTimeout(() => {
      const banner = document.getElementById('testModeBanner');
      const btn = document.getElementById('btnTestMode');
      if (banner) banner.style.display = 'block';
      if (btn) {
        btn.style.background = '#a4302a';
        btn.style.color = '#fff';
        btn.textContent = 'TEST ●';
      }
    }, 100);
  }
}

/* ===========================================================
   PDF GENERATION
   Reproduces the exact invoice layout as a PDF using jsPDF,
   then downloads it. Used by the Email flow (user attaches
   the downloaded PDF to Gmail manually).
   =========================================================== */

async function invGeneratePDF(inv, cl, isPi) {
  // Dynamically load jsPDF if not already loaded
  if (!window.jspdf) {
    await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const W = 210, ML = 15, MR = 195;
  const intra = inv.gstType === 'intra';
  const title = isPi ? 'PROFORMA INVOICE' : 'TAX INVOICE';

  // Header rule
  doc.setFillColor(15, 37, 64);
  doc.rect(ML, 10, MR - ML, 8, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(12); doc.setFont('helvetica', 'bold');
  doc.text(title, W / 2, 15.5, { align: 'center' });

  // From block
  doc.setTextColor(30, 30, 30);
  doc.setFontSize(7.5); doc.setFont('helvetica', 'bold');
  doc.text('From,', ML, 23);
  doc.setFont('helvetica', 'normal');
  doc.text(GORU.name, ML, 27);
  doc.text('No 1108, Floor No 11th, Sureshwari Techno IT Park Premises,', ML, 31);
  doc.text('Link Road, Near Eskay Resorts, Borivali 400092 Mumbai', ML, 35);
  doc.text(`State: ${GORU.state}   GSTIN: ${GORU.gstin}`, ML, 39);
  doc.text(`PAN: ${GORU.pan}   TAN: ${GORU.tan}`, ML, 43);

  // To block
  const toX = 110;
  doc.setFont('helvetica', 'bold');
  doc.text('To,', toX, 23);
  doc.setFont('helvetica', 'normal');
  const addrLines = [cl.companyName||'', cl.addr1||'', cl.addr2||'', cl.addr3||''].filter(Boolean);
  addrLines.forEach((line, i) => doc.text(line, toX, 27 + i * 4));
  const toY = 27 + addrLines.length * 4;
  if (cl.state) doc.text(`State: ${cl.state}`, toX, toY);
  if (cl.gstin) doc.text(`GSTIN: ${cl.gstin}`, toX, toY + 4);
  if (cl.pan) doc.text(`PAN: ${cl.pan}`, toX, toY + 8);

  // Invoice details (right column)
  const dtX = 155;
  doc.setFont('helvetica', 'bold');
  doc.text(isPi ? 'Proforma Invoice No:' : 'Tax Invoice No:', dtX, 23);
  doc.text(isPi ? 'Proforma Invoice Date:' : 'Tax Invoice Date:', dtX, 27);
  doc.text('Date of Supply:', dtX, 31);
  doc.text('Place of Supply:', dtX, 35);
  doc.text('Kind Attn:', dtX, 39);
  doc.text('Mob:', dtX, 43);
  doc.setFont('helvetica', 'normal');
  doc.text(inv.invNo, dtX + 38, 23);
  doc.text(inv.date || '', dtX + 38, 27);
  doc.text(inv.supplyDate || '', dtX + 38, 31);
  doc.text(cl.state || '', dtX + 38, 35);
  doc.text(cl.attn || '', dtX + 38, 39);
  doc.text(cl.mobile || '', dtX + 38, 43);

  // Divider
  doc.setDrawColor(180, 180, 180);
  doc.line(ML, 48, MR, 48);

  // Table header
  let y = 52;
  doc.setFillColor(240, 242, 246);
  doc.rect(ML, y - 4, MR - ML, 7, 'F');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
  const cols = { sr:ML, desc:22, sac:80, unit:95, rate:108, net:122, disc:136, taxable:148, tax:162, gross:185 };
  doc.text('SR', cols.sr, y);
  doc.text('DESCRIPTION', cols.desc, y);
  doc.text('SAC', cols.sac, y);
  doc.text('UNIT', cols.unit, y);
  doc.text('RATE', cols.rate, y);
  doc.text('NET AMT', cols.net, y);
  doc.text('DISC', cols.disc, y);
  doc.text('TAXABLE', cols.taxable, y);
  doc.text(intra ? 'CGST+SGST' : 'IGST', cols.tax, y);
  doc.text('GROSS', cols.gross, y);
  y += 7;

  // Row A
  doc.setFont('helvetica', 'normal');
  const net = inv.rate * (inv.qty || 1);
  const taxAmt = intra ? (inv.cgst + inv.sgst) : inv.igst;
  const rowGross = inv.gross - (inv.freight || 0) - (inv.other || 0);
  doc.text('A', cols.sr, y);
  const descLines = doc.splitTextToSize(inv.desc || '', 55);
  doc.text(descLines, cols.desc, y);
  doc.text(String(inv.sac || '998399'), cols.sac, y);
  doc.text(inv.unit || '', cols.unit, y);
  doc.text(fmt2(inv.rate), cols.rate, y);
  doc.text(fmt2(net), cols.net, y);
  doc.text(fmt2(inv.disc || 0), cols.disc, y);
  doc.text(fmt2(inv.taxable), cols.taxable, y);
  doc.text(fmt2(taxAmt), cols.tax, y);
  doc.text(fmt2(rowGross), cols.gross, y);
  y += Math.max(7, descLines.length * 4 + 3);

  // Row B (freight)
  if (inv.freight) {
    doc.text('B', cols.sr, y); doc.text('Freight charges', cols.desc, y);
    doc.text(fmt2(inv.freight), cols.taxable, y);
    doc.text(fmt2(inv.freight * (intra ? 0.18 : 0.18)), cols.tax, y);
    doc.text(fmt2(inv.freight * 1.18), cols.gross, y); y += 7;
  }
  // Row C (other)
  if (inv.other) {
    doc.text('C', cols.sr, y); doc.text('Other Taxable Charges', cols.desc, y);
    doc.text(fmt2(inv.other), cols.taxable, y);
    doc.text(fmt2(inv.other * 0.18), cols.tax, y);
    doc.text(fmt2(inv.other * 1.18), cols.gross, y); y += 7;
  }

  // Totals row
  doc.setFillColor(240, 242, 246);
  doc.rect(ML, y - 4, MR - ML, 7, 'F');
  doc.setFont('helvetica', 'bold');
  doc.text('TOTAL', cols.sr, y);
  doc.text(fmt2(inv.taxable), cols.taxable, y);
  doc.text(fmt2(taxAmt), cols.tax, y);
  doc.text(fmt2(inv.gross), cols.gross, y);
  y += 10;

  // Amount in words
  doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
  doc.text(`Gross Amount in Words: ${numberToWords(Math.round(inv.gross))}`, ML, y);
  y += 8;

  // Payment terms
  if (inv.paymentTerms) {
    doc.text(`Payment Terms: ${inv.paymentTerms}`, ML, y); y += 5;
  }
  doc.text('All payments by bank transfer/draft/cheque payable at Mumbai in favour of "Goru Training Pvt. Ltd."', ML, y); y += 4;
  doc.text(`Bank: ${GORU.bank}, Branch: ${GORU.branch}`, ML, y); y += 4;
  doc.text(`A/c No: ${GORU.acno}   IFSC: ${GORU.ifsc}`, ML, y); y += 10;

  // Authorised signatory
  doc.setFont('helvetica', 'bold');
  doc.text('FOR GORU TRAINING PRIVATE LIMITED', MR - 5, y, { align: 'right' }); y += 10;
  doc.text('AUTHORISED SIGNATORY', MR - 5, y, { align: 'right' });

  // Border around entire document
  doc.setDrawColor(180, 180, 180);
  doc.rect(ML, 8, MR - ML, y, 'S');

  const filename = `${inv.invNo.replace(/\//g, '-')}.pdf`;
  doc.save(filename);
  return filename;
}

function fmt2(n) {
  return (Number(n) || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/* ===========================================================
   IMPROVED EMAIL MODAL
   Downloads PDF, then opens Gmail compose with pre-filled
   subject, To, and body. User attaches the downloaded PDF.
   =========================================================== */

async function invEmail(invId, type) {
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;
  const cl = DB.invoicing.clients.find(c => c.id === inv.clientId) || {};
  const isPi = type === 'pi';
  const typePrefix = isPi ? 'PI' : 'TI';

  // Build subject: "PI/26-27/028 — Your order for 3 of EPP"
  const firstLine = (inv.lineItems && inv.lineItems.find(l => l._type !== 'freight')) || {};
  const qty = firstLine.qty || inv.qty || '';
  const shortName = firstLine.shortName || (() => {
    if (firstLine.desc) {
      const m = firstLine.desc.match(/\(([A-Z]{2,6})\)/);
      if (m) return m[1];
      return firstLine.desc.split(' ').slice(0, 2).join(' ');
    }
    return typePrefix;
  })();
  const defaultSubject = `${inv.invNo} — Your order for ${qty} of ${shortName}`;

  // Programme description for body (first non-freight line)
  const progDesc = firstLine.desc || inv.desc || '';
  const progQty  = qty;
  const progShort = shortName;
  const grossFmt = (inv.gross || 0).toLocaleString('en-IN', {minimumFractionDigits:2});
  const recvFmt  = (inv.receivableAmount || inv.gross || 0).toLocaleString('en-IN', {minimumFractionDigits:2});
  const tdsLine  = inv.tdsDeducted === 'yes'
    ? `\nTotal payable after TDS is ₹${recvFmt}.` : '';

  // Templates from Nirali's standard text
  const tplPI = `Dear ${cl.attn || ''},\n\nThank you for your order of ${progQty} ${progShort} kit${progQty > 1 ? 's' : ''}.\n\nPFA proforma invoice ${inv.invNo} for the same.${tdsLine}\n\nTax invoice will be raised once we receive full payment.\n\nMany thanks,\n\nRegards,\nNirali Bhakta`;

  const tplTI = `Dear ${cl.attn || ''},\n\nAcknowledging receipt of ₹${recvFmt} in our account. Many thanks.\n\nPFA tax invoice ${inv.invNo} for the same.\n\nThe kits will be dispatched today. Kindly acknowledge receipt.\n\nMany thanks,\n\nRegards,\nNirali Bhakta`;

  const defaultBody = isPi ? tplPI : tplTI;

  const body = `
    <div class="hint" style="background:#e8f0fb; margin-bottom:12px; font-size:12.5px;">
      <strong>Step 1</strong> — Download Word doc &nbsp;·&nbsp;
      <strong>Step 2</strong> — Open Gmail, attach the downloaded file, send.
    </div>
    <div class="field">
      <label>To</label>
      <input id="em-to" value="${escapeHtml(cl.email || '')}">
    </div>
    <div class="field">
      <label>Subject</label>
      <input id="em-subject" value="${escapeHtml(defaultSubject)}">
    </div>
    <div class="field">
      <label>Template</label>
      <select id="em-template" style="margin-bottom:8px;">
        <option value="standard">Standard (Nirali Bhakta)</option>
        <option value="custom">Custom — edit below</option>
      </select>
      <textarea id="em-body" rows="9" style="width:100%; padding:9px 11px; border:1px solid var(--line); border-radius:5px; font-size:13px; font-family:var(--sans); resize:vertical;">${escapeHtml(defaultBody)}</textarea>
    </div>`;

  openModal(`Email ${inv.invNo}`, body,
    `<button class="btn" id="em-cancel">Cancel</button>
     <button class="btn btn-primary" id="em-word">&#8659; Download Word doc</button>
     <button class="btn btn-primary" id="em-open-gmail">&#9993; Open Gmail</button>`);

  document.getElementById('em-cancel').onclick = closeModal;

  // Template switcher — live swap body text
  document.getElementById('em-template').onchange = e => {
    const ta = document.getElementById('em-body');
    if (e.target.value === 'standard') ta.value = isPi ? tplPI : tplTI;
  };

  document.getElementById('em-word').onclick = async () => {
    const btn = document.getElementById('em-word');
    btn.textContent = 'Generating…';
    btn.disabled = true;
    try {
      await invDownloadWord(inv.id, isPi ? 'pi' : 'ti');
      btn.textContent = '✓ Downloaded';
      toast(`${inv.invNo}.docx downloaded — attach in Gmail`);
    } catch(err) {
      btn.textContent = '⬇ Download Word doc';
      btn.disabled = false;
      toast('Word generation failed: ' + err.message);
    }
  };

  document.getElementById('em-open-gmail').onclick = () => {
    const to   = document.getElementById('em-to').value.trim();
    const subj = document.getElementById('em-subject').value.trim();
    const txt  = document.getElementById('em-body').value.trim();
    window.open(
      `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subj)}&body=${encodeURIComponent(txt)}`,
      '_blank'
    );
    closeModal();
    toast('Gmail opened — attach the Word doc before sending');
  };
}

/* ===========================================================
   PRODUCT MASTER
   Shared product list stored in DB.invoicing.products.
   Each product: { id, shortName, longName, rate, category }
   category: 'OFFLINE' | 'ONLINE' | 'OTHER' (auto or manual override)
   Auto-categorisation:
     ONLINE  — last word of shortName ends with 'O'
     OFFLINE — shortName ≤ 4 chars and NOT online
     OTHER   — everything else
   =========================================================== */

function productCategory(shortName) {
  if (!shortName) return 'OTHER';
  const words = shortName.trim().split(/\s+/);
  const last = words[words.length - 1];
  if (last.toUpperCase().endsWith('O')) return 'ONLINE';
  if (shortName.trim().length <= 4) return 'OFFLINE';
  return 'OTHER';
}

// ── Add / Edit Product ──────────────────────────────────────
function invOpenAddProduct(editId) {
  invInit();
  const existing = editId ? DB.invoicing.products.find(p => p.id === editId) : null;
  const v = existing || {};
  const cats = ['OFFLINE', 'ONLINE', 'OTHER'];
  const autocat = v.shortName ? productCategory(v.shortName) : 'OFFLINE';
  const body = `
    <div class="field-row">
      <div class="field"><label>Short name</label>
        <input id="pd-short" value="${escapeHtml(v.shortName||'')}" placeholder="e.g. EPP, ONTOP">
        <div style="font-size:11px; color:var(--ink-soft); margin-top:3px;" id="pd-cat-hint">Auto-category: <strong>${autocat}</strong></div>
      </div>
      <div class="field"><label>Rate (₹)</label>
        <input type="number" id="pd-rate" value="${v.rate||''}">
      </div>
    </div>
    <div class="field"><label>Long name / description</label>
      <input id="pd-long" value="${escapeHtml(v.longName||'')}" placeholder="e.g. Enhanced Performance Program">
    </div>
    <div class="field"><label>Category</label>
      <select id="pd-cat">
        ${cats.map(c => `<option value="${c}" ${(v.category||autocat)===c?'selected':''}>${c}</option>`).join('')}
      </select>
      <div style="font-size:11px; color:var(--ink-soft); margin-top:3px;">Auto-assigned from short name unless overridden here.</div>
    </div>`;

  openModal(existing ? 'Edit product' : 'Add product', body,
    `<button class="btn" id="pd-cancel">Cancel</button>
     <button class="btn btn-primary" id="pd-save">Save</button>`);

  // Auto-update category hint as user types
  document.getElementById('pd-short').addEventListener('input', e => {
    const cat = productCategory(e.target.value);
    document.getElementById('pd-cat-hint').innerHTML = `Auto-category: <strong>${cat}</strong>`;
    document.getElementById('pd-cat').value = cat;
  });

  document.getElementById('pd-cancel').onclick = closeModal;
  document.getElementById('pd-save').onclick = () => {
    const shortName = document.getElementById('pd-short').value.trim();
    const longName  = document.getElementById('pd-long').value.trim();
    const rate      = parseFloat(document.getElementById('pd-rate').value) || 0;
    const category  = document.getElementById('pd-cat').value;
    if (!shortName) { toast('Short name is required'); return; }
    if (!longName)  { toast('Long name is required'); return; }
    if (existing) {
      Object.assign(existing, { shortName, longName, rate, category });
    } else {
      DB.invoicing.products.push({ id: uid(), shortName, longName, rate, category });
    }
    saveDB(); closeModal();
    toast(`Product "${shortName}" saved`);
    invOpenProductList();
  };
}

// ── Product List ───────────────────────────────────────────
function invOpenProductList() {
  invInit();
  const products = DB.invoicing.products;
  const grouped = { OFFLINE: [], ONLINE: [], OTHER: [] };
  products.forEach(p => grouped[p.category || productCategory(p.shortName)].push(p));

  function sectionHtml(label, items, color) {
    if (!items.length) return '';
    return `
      <div style="font-size:11px; font-weight:700; letter-spacing:.08em; color:${color};
                  text-transform:uppercase; padding:10px 0 6px 0; border-bottom:2px solid ${color}22;
                  margin-top:14px;">${label}</div>
      ${items.map(p => `
        <div style="display:flex; align-items:center; gap:8px; padding:8px 0;
                    border-bottom:1px solid var(--line); font-size:13px;">
          <span style="font-weight:700; min-width:60px; color:var(--navy);">${escapeHtml(p.shortName)}</span>
          <span style="flex:1; color:var(--ink);">${escapeHtml(p.longName)}</span>
          <span style="font-family:var(--mono); min-width:80px; text-align:right;">${fmtMoney(p.rate)}</span>
          <span class="row-actions">
            <button data-pd-move="${p.id}" title="Move category">&#8644;</button>
            <button data-pd-edit="${p.id}" title="Edit">&#9998;</button>
            <button data-pd-del="${p.id}" title="Delete" style="color:var(--red);">&#10005;</button>
          </span>
        </div>`).join('')}`;
  }

  const body = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <span style="font-size:12px; color:var(--ink-soft);">${products.length} product${products.length!==1?'s':''} on file</span>
      <label class="btn btn-sm" style="cursor:pointer;">
        &#8659; Import CSV
        <input type="file" id="pd-csv-input" accept=".csv" style="display:none;">
      </label>
    </div>
    <div style="font-size:11px; color:var(--ink-soft); margin-bottom:10px;">
      CSV format: Short Name, Long Name, Rate, Category — new items added, existing updated by Short Name.
    </div>
    <div style="max-height:460px; overflow-y:auto;">
      ${sectionHtml('OFFLINE', grouped.OFFLINE, '#1e4f8a')}
      ${sectionHtml('ONLINE', grouped.ONLINE, '#1f7a4d')}
      ${sectionHtml('OTHER', grouped.OTHER, '#9a6b14')}
      ${!products.length ? '<div class="empty-note" style="padding:20px;">No products yet. Click + Add product or import a CSV.</div>' : ''}
    </div>`;

  openModal('Product list', body,
    `<button class="btn" id="pd-list-close">Close</button>
     <button class="btn btn-primary" id="pd-list-add">+ Add product</button>`);

  document.getElementById('pd-list-close').onclick = closeModal;
  document.getElementById('pd-list-add').onclick = () => { closeModal(); invOpenAddProduct(); };

  // Edit
  document.querySelectorAll('[data-pd-edit]').forEach(b =>
    b.onclick = () => { closeModal(); invOpenAddProduct(b.dataset.pdEdit); });

  // Delete with Y/N
  document.querySelectorAll('[data-pd-del]').forEach(b => {
    b.onclick = () => {
      const p = DB.invoicing.products.find(x => x.id === b.dataset.pdDel);
      if (!p) return;
      if (!confirm(`Delete product "${p.shortName} — ${p.longName}"?\nThis cannot be undone.`)) return;
      DB.invoicing.products = DB.invoicing.products.filter(x => x.id !== b.dataset.pdDel);
      saveDB(); closeModal(); invOpenProductList();
      toast(`Product "${p.shortName}" deleted`);
    };
  });

  // Move category
  document.querySelectorAll('[data-pd-move]').forEach(b => {
    b.onclick = () => {
      const p = DB.invoicing.products.find(x => x.id === b.dataset.pdMove);
      if (!p) return;
      const cats = ['OFFLINE','ONLINE','OTHER'];
      const movebody = `
        <div style="margin-bottom:14px; font-size:13px;">
          Moving: <strong>${escapeHtml(p.shortName)}</strong> — ${escapeHtml(p.longName)}<br>
          Current category: <strong>${p.category || productCategory(p.shortName)}</strong>
        </div>
        <div class="field"><label>Move to</label>
          <select id="pd-move-cat">
            ${cats.map(c => `<option value="${c}" ${(p.category||productCategory(p.shortName))===c?'selected':''}>${c}</option>`).join('')}
          </select>
        </div>`;
      openModal('Move product', movebody,
        `<button class="btn" id="pd-move-cancel">Cancel</button>
         <button class="btn btn-primary" id="pd-move-save">Move</button>`);
      document.getElementById('pd-move-cancel').onclick = () => { closeModal(); invOpenProductList(); };
      document.getElementById('pd-move-save').onclick = () => {
        p.category = document.getElementById('pd-move-cat').value;
        saveDB(); closeModal(); invOpenProductList();
        toast(`Moved "${p.shortName}" to ${p.category}`);
      };
    };
  });

  // CSV import
  document.getElementById('pd-csv-input').onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const lines = ev.target.result.split('\n').filter(l => l.trim());
      let added = 0, updated = 0, errors = 0;
      lines.forEach((line, idx) => {
        if (idx === 0 && line.toLowerCase().includes('short')) return; // skip header
        const parts = line.split(',').map(s => s.trim().replace(/^"|"$/g, ''));
        if (parts.length < 3) { errors++; return; }
        const [shortName, longName, rateStr, category] = parts;
        if (!shortName || !longName) { errors++; return; }
        const rate = parseFloat(rateStr) || 0;
        const cat = ['OFFLINE','ONLINE','OTHER'].includes((category||'').toUpperCase())
          ? category.toUpperCase() : productCategory(shortName);
        const existing2 = DB.invoicing.products.find(p =>
          p.shortName.toLowerCase() === shortName.toLowerCase());
        if (existing2) {
          existing2.longName = longName; existing2.rate = rate; existing2.category = cat;
          updated++;
        } else {
          DB.invoicing.products.push({ id: uid(), shortName, longName, rate, category: cat });
          added++;
        }
      });
      saveDB(); closeModal(); invOpenProductList();
      toast(`CSV imported: ${added} added, ${updated} updated${errors ? ', ' + errors + ' skipped' : ''}`);
    };
    reader.readAsText(file);
    e.target.value = '';
  };
}
