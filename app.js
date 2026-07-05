/* ===========================================================
   LMI Cashflow Manager — application logic
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
    DB.currentFY = sel.value;
    const months = monthsOfFY(DB.currentFY);
    let touched = [];
    if (!months.includes(DB.selectedMonth)) {
      touched = selectMonth(months[0]);
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
  if (!m) return false; // doesn't exist yet — handled by the isNew branch
  // Once a month has been provisioned, never re-provision it — even if the user
  // deleted items they didn't want. The _provisioned stamp is set by ensureMonthlyProvisions.
  if (m._provisioned) return false;
  // Only auto-heal months strictly after today
  if (mk <= todayMonthKey()) return false;
  const hasAnyRecurringTemplate = DB.recurringTemplate && DB.recurringTemplate.length > 0;
  const hasRecurringPayments = m.payments.some(p => p.recurring);
  const hasTdsLine = m.payments.some(p => /^TDS provisional/i.test(p.name));
  if (hasAnyRecurringTemplate && !hasRecurringPayments) return true;
  if (!hasTdsLine) return true;
  return false;
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
  wrap.innerHTML = list.map((r, i) => `
    <div class="recur-item">
      <span class="nm"><span class="star">&#9733;</span>${escapeHtml(r.name)}</span>
      <span class="amt">${fmtMoney(r.amount)}</span>
    </div>`).join('');
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
function renderRecurringEditor() {
  const list = DB.recurringTemplate;
  const body = `
    <div class="sub-list" id="rec-edit-list" style="max-height:280px;">
      ${list.length ? list.map((r, i) => `
        <div class="sub-list-item" style="cursor:default;">
          <span style="flex:1;">${escapeHtml(r.name)}</span>
          <input type="number" data-rec-idx="${i}" value="${r.amount}" style="width:110px; padding:5px 8px; border:1px solid var(--line); border-radius:4px; font-family:var(--mono); margin-left:8px;">
          <button data-rec-del="${i}" style="border:none;background:none;color:#aab2bd;cursor:pointer;margin-left:6px;">&#10005;</button>
        </div>`).join('') : '<div class="empty-note" style="padding:12px;">No recurring items yet.</div>'}
    </div>
    <div class="hint">Saving applies these amounts to ${monthLabel(DB.selectedMonth)} and every month after it this financial year. Earlier months are never changed.</div>
    <button class="btn btn-sm" id="rec-add-new">+ Add recurring item</button>`;
  openModal('Edit recurring payments', body, `<button class="btn" id="rec-cancel">Cancel</button><button class="btn btn-primary" id="rec-save">Save changes</button>`);
  document.getElementById('rec-cancel').onclick = closeModal;
  document.getElementById('rec-add-new').onclick = () => {
    const name = prompt('New recurring payment name:');
    if (!name) return;
    DB.recurringTemplate.push({ name: name.trim(), amount: 0, tds: false });
    renderRecurringEditor();
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
    // Also ensure TDS provisional / quarterly Waco fee are present in those same upcoming months.
    touched.forEach(mk => ensureMonthlyProvisions(mk));
    saveDB(touched); renderAll(); closeModal();
    toast(`Recurring payments updated from ${monthLabel(DB.selectedMonth)} onward`);
  };
}

// Pushes current recurringTemplate into the selected month and all later months in the
// same FY — never touches months before fromMk. Returns the touched month keys.
function applyRecurringForward(fromMk) {
  const months = monthsOfFY(fyLabelForMonth(fromMk)).filter(mk => mk >= fromMk);
  months.forEach(mk => {
    const m = ensureMonthExists(mk);
    DB.recurringTemplate.forEach(tpl => {
      let existing = m.payments.find(p => p.recurring && p.name.toLowerCase().startsWith(tpl.name.toLowerCase()));
      if (existing) {
        existing.amount = tpl.amount;
      } else {
        m.payments.push({ id: uid(), name: `${tpl.name} for ${monthLabel(mk)}`, amount: tpl.amount, status: 'planned', recurring: true, tds: !!tpl.tds });
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
  if (!confirm('Delete this receivable?')) return;
  const m = getMonth(DB.selectedMonth);
  m.receivables = m.receivables.filter(r => r.id !== id);
  if (DB.selectedMonth === todayMonthKey()) syncCarriedReceivables();
  saveDB([nextMonthKey(todayMonthKey())]); renderAll();
}


/* ===========================================================
   AUTOMATIC PROVISIONS
   - TDS provisional line each month (default est. 1,50,000), editable per month
   - TDS flagged on payments rolls into next month's TDS line automatically
   - Quarterly Waco marketing fee (Apr/Jul/Oct/Jan)
   =========================================================== */
function ensureMonthlyProvisions(mk) {
  const m = ensureMonthExists(mk);

  // Mark as provisioned immediately — this prevents re-provisioning if the user
  // later deletes items they don't want (e.g. recurring items carried from seed data
  // that were manually removed). Once stamped, this month is never auto-provisioned again.
  m._provisioned = true;

  // Recurring (starred) payments — only add if not already present
  DB.recurringTemplate.forEach(tpl => {
    const existing = m.payments.find(p => p.recurring && p.name.toLowerCase().startsWith(tpl.name.toLowerCase()));
    if (!existing) {
      m.payments.push({ id: uid(), name: `${tpl.name} for ${monthLabel(mk)}`, amount: tpl.amount, status: 'planned', recurring: true, tds: !!tpl.tds });
    }
  });

  // TDS provisional line (auto-added if not present)
  let tdsLine = m.payments.find(p => /^TDS provisional/i.test(p.name));
  if (!tdsLine) {
    tdsLine = { id: uid(), name: 'TDS provisional', amount: TDS_ESTIMATE, status: 'planned', recurring: false, tds: false, _auto: true };
    m.payments.push(tdsLine);
  }

  // Quarterly Waco marketing fee
  const mNum = monthNum(mk);
  if (QUARTERLY_MONTHS.includes(mNum)) {
    const has = m.payments.find(p => /waco marketing fee/i.test(p.name));
    if (!has) {
      m.payments.push({ id: uid(), name: 'Waco marketing fee', amount: WACO_FEE, status: 'planned', recurring: false, tds: false, _auto: true });
    }
  }
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
    tdsLine = { id: uid(), name: 'TDS provisional', amount: TDS_ESTIMATE, status: 'planned', recurring: false, tds: false, _auto: true };
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
      ensureMonthExists(DB.selectedMonth);
      ensureMonthlyProvisions(DB.selectedMonth);
      recalcTDSRollup(prevMonthKey(DB.selectedMonth));
      promoteCarriedReceivables(); // unlock carried entries now that their month is current
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
          <button data-inv-excel="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Download Excel">&#128229;</button>
          <button data-inv-email="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Email">&#9993;</button>
          ${inv.status !== 'cancelled' && inv.status !== 'converted' ? `<button data-inv-del="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Cancel/Delete" style="color:var(--red);">&#10005;</button>` : ''}
        </span>
      </td>
    </tr>`).join('');

  body.querySelectorAll('[data-inv-edit]').forEach(b =>
    b.onclick = () => b.dataset.invType === 'pi'
      ? invOpenPIForm(b.dataset.invEdit)
      : invOpenTIForm(null, b.dataset.invEdit));
  body.querySelectorAll('[data-inv-to-ti]').forEach(b =>
    b.onclick = () => invOpenTIForm(b.dataset.invToTi));
  body.querySelectorAll('[data-inv-excel]').forEach(b =>
    b.onclick = () => invDownloadExcel(b.dataset.invExcel, b.dataset.invType));
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
function invOpenPIForm(editId) {
  invInit();
  const existing = editId ? DB.invoicing.proformas.find(p => p.id === editId) : null;
  const v = existing || {};
  const clients = DB.invoicing.clients;
  const today = new Date().toISOString().slice(0, 10);

  const body = `
    <div class="field-row">
      <div class="field"><label>Client</label>
        <select id="pf-client">${clients.map(c =>
          `<option value="${c.id}" ${v.clientId===c.id?'selected':''}>${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</option>`
        ).join('')}</select>
      </div>
      <div class="field"><label>Invoice date</label>
        <input type="date" id="pf-date" value="${v.date || today}">
      </div>
    </div>
    <div class="field-row">
      <div class="field"><label>Date of supply</label>
        <input type="date" id="pf-supply" value="${v.supplyDate || today}">
      </div>
      <div class="field"><label>Payment terms</label>
        <input id="pf-terms" value="${escapeHtml(v.paymentTerms || 'Advance')}">
      </div>
    </div>
    <div class="field"><label>Row A — Description</label>
      <input id="pf-desc" value="${escapeHtml(v.desc || '')}"></div>
    <div class="field-row">
      <div class="field"><label>Row A — Unit</label><input id="pf-unit" value="${escapeHtml(v.unit || '')}"></div>
      <div class="field"><label>Row A — Rate</label><input type="number" id="pf-rate" value="${v.rate || ''}"></div>
      <div class="field"><label>Row A — Qty</label><input type="number" id="pf-qty" value="${v.qty || 1}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Row A — Discount</label><input type="number" id="pf-disc" value="${v.disc || 0}"></div>
      <div class="field"><label>Row A — SAC code</label><input id="pf-sac" value="${v.sac || '998399'}"></div>
    </div>
    <div class="field"><label>Row B — Freight charges (blank = 0)</label>
      <input type="number" id="pf-freight" value="${v.freight || ''}"></div>
    <div class="field"><label>Row C — Other taxable charges (blank = 0)</label>
      <input type="number" id="pf-other" value="${v.other || ''}"></div>
    <div id="pf-gst-preview" style="background:var(--paper); border-radius:6px; padding:10px 14px; font-size:12.5px; margin-top:4px;"></div>`;

  openModal(existing ? 'Edit Proforma Invoice' : 'New Proforma Invoice', body,
    `<button class="btn" id="pf-cancel">Cancel</button><button class="btn btn-primary" id="pf-save">${existing ? 'Save changes' : 'Generate PI'}</button>`);

  const preview = () => {
    const cid = document.getElementById('pf-client').value;
    const cl = clients.find(c => c.id === cid);
    const rate = parseFloat(document.getElementById('pf-rate').value) || 0;
    const qty = parseFloat(document.getElementById('pf-qty').value) || 1;
    const disc = parseFloat(document.getElementById('pf-disc').value) || 0;
    const freight = parseFloat(document.getElementById('pf-freight').value) || 0;
    const other = parseFloat(document.getElementById('pf-other').value) || 0;
    const net = rate * qty;
    const taxable = net - disc + freight + other;
    const type = cl ? gstType(cl.state) : 'igst';
    const cgst = type === 'intra' ? taxable * 0.09 : 0;
    const sgst = type === 'intra' ? taxable * 0.09 : 0;
    const igst = type === 'igst' ? taxable * 0.18 : 0;
    const gross = taxable + cgst + sgst + igst;
    const gstLabel = type === 'intra' ? 'CGST 9% + SGST 9%' : 'IGST 18%';
    document.getElementById('pf-gst-preview').innerHTML =
      `<b>Preview:</b> Net ${fmtMoney(net)} 2212 Disc ${fmtMoney(disc)} = Taxable ${fmtMoney(taxable)} + ${gstLabel} = <b>Gross ${fmtMoney(gross)}</b>`;
    return { net, disc, taxable, freight, other, cgst, sgst, igst, gross, type };
  };

  ['pf-client','pf-rate','pf-qty','pf-disc','pf-freight','pf-other'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', preview);
    if (el && el.tagName === 'SELECT') el.addEventListener('change', preview);
  });
  preview();

  document.getElementById('pf-cancel').onclick = closeModal;
  document.getElementById('pf-save').onclick = () => {
    const cid = document.getElementById('pf-client').value;
    const cl = clients.find(c => c.id === cid);
    if (!cl) { toast('Select a client'); return; }
    const desc = document.getElementById('pf-desc').value.trim();
    if (!desc) { toast('Description is required'); return; }
    const calc = preview();
    const rec = {
      id: existing ? existing.id : uid(),
      invNo: existing ? existing.invNo : nextPINumber(),
      date: document.getElementById('pf-date').value,
      supplyDate: document.getElementById('pf-supply').value,
      paymentTerms: document.getElementById('pf-terms').value,
      clientId: cid, clientName: cl.companyName, clientShort: cl.shortName,
      desc, unit: document.getElementById('pf-unit').value,
      rate: parseFloat(document.getElementById('pf-rate').value) || 0,
      qty: parseFloat(document.getElementById('pf-qty').value) || 1,
      sac: document.getElementById('pf-sac').value || '998399',
      disc: calc.disc, freight: calc.freight, other: calc.other,
      taxable: calc.taxable, cgst: calc.cgst, sgst: calc.sgst,
      igst: calc.igst, gross: calc.gross, gstType: calc.type,
      status: existing ? existing.status : 'draft',
      piNo: null, // for TIs — the PI it was raised from
    };
    if (existing) {
      Object.assign(existing, rec);
    } else {
      DB.invoicing.proformas.push(rec);
      // Auto-add to cashflow receivables for current month
      invAddToReceivables(rec, 'PI');
    }
    DB.invoicing.piSequence[currentInvFY()] = parseInt(rec.invNo.split('/').pop(), 10);
    saveDB(); closeModal(); invRenderRegister();
    toast(`${rec.invNo} ${existing ? 'updated' : 'generated'}`);
  };
}

// ── Invoice form (TI) ──────────────────────────────────────
function invOpenTIForm(fromPiId, editTiId) {
  invInit();
  const existingTI = editTiId ? DB.invoicing.taxInvoices.find(t => t.id === editTiId) : null;
  const sourcePi = fromPiId ? DB.invoicing.proformas.find(p => p.id === fromPiId) : null;
  const v = existingTI || sourcePi || {};
  const clients = DB.invoicing.clients;
  const today = new Date().toISOString().slice(0, 10);

  // If converting from PI, pre-select that client and pre-fill
  const preClient = v.clientId || '';
  const body = `
    ${sourcePi ? `<div class="hint" style="background:var(--blue-bg); margin-bottom:12px;">Converting from ${escapeHtml(sourcePi.invNo)}. Review details and confirm or change below.</div>` : ''}
    <div class="field-row">
      <div class="field"><label>Client</label>
        <select id="ti-client">${clients.map(c =>
          `<option value="${c.id}" ${preClient===c.id?'selected':''}>${escapeHtml(c.shortName)} — ${escapeHtml(c.companyName)}</option>`
        ).join('')}</select>
      </div>
      <div class="field"><label>Invoice date</label>
        <input type="date" id="ti-date" value="${existingTI ? existingTI.date : today}">
      </div>
    </div>
    <div class="field-row">
      <div class="field"><label>Date of supply</label>
        <input type="date" id="ti-supply" value="${v.supplyDate || today}">
      </div>
      <div class="field"><label>Payment terms</label>
        <input id="ti-terms" value="${escapeHtml(v.paymentTerms || 'Advance')}">
      </div>
    </div>
    <div class="field"><label>Row A — Description</label>
      <input id="ti-desc" value="${escapeHtml(v.desc || '')}"></div>
    <div class="field-row">
      <div class="field"><label>Row A — Unit</label><input id="ti-unit" value="${escapeHtml(v.unit || '')}"></div>
      <div class="field"><label>Row A — Rate</label><input type="number" id="ti-rate" value="${v.rate || ''}"></div>
      <div class="field"><label>Row A — Qty</label><input type="number" id="ti-qty" value="${v.qty || 1}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Row A — Discount</label><input type="number" id="ti-disc" value="${v.disc || 0}"></div>
      <div class="field"><label>Row A — SAC code</label><input id="ti-sac" value="${v.sac || '998399'}"></div>
    </div>
    <div class="field"><label>Row B — Freight charges</label>
      <input type="number" id="ti-freight" value="${v.freight || ''}"></div>
    <div class="field"><label>Row C — Other taxable charges</label>
      <input type="number" id="ti-other" value="${v.other || ''}"></div>
    <div id="ti-gst-preview" style="background:var(--paper); border-radius:6px; padding:10px 14px; font-size:12.5px; margin-top:4px;"></div>`;

  openModal(existingTI ? 'Edit Tax Invoice' : 'New Tax Invoice', body,
    `<button class="btn" id="ti-cancel">Cancel</button><button class="btn btn-primary" id="ti-save">${existingTI ? 'Save changes' : 'Generate TI'}</button>`);

  const preview = () => {
    const cid = document.getElementById('ti-client').value;
    const cl = clients.find(c => c.id === cid);
    const rate = parseFloat(document.getElementById('ti-rate').value) || 0;
    const qty = parseFloat(document.getElementById('ti-qty').value) || 1;
    const disc = parseFloat(document.getElementById('ti-disc').value) || 0;
    const freight = parseFloat(document.getElementById('ti-freight').value) || 0;
    const other = parseFloat(document.getElementById('ti-other').value) || 0;
    const net = rate * qty;
    const taxable = net - disc + freight + other;
    const type = cl ? gstType(cl.state) : 'igst';
    const cgst = type === 'intra' ? taxable * 0.09 : 0;
    const sgst = type === 'intra' ? taxable * 0.09 : 0;
    const igst = type === 'igst' ? taxable * 0.18 : 0;
    const gross = taxable + cgst + sgst + igst;
    const gstLabel = type === 'intra' ? 'CGST 9% + SGST 9%' : 'IGST 18%';
    document.getElementById('ti-gst-preview').innerHTML =
      `<b>Preview:</b> Net ${fmtMoney(net)} 2212 Disc ${fmtMoney(disc)} = Taxable ${fmtMoney(taxable)} + ${gstLabel} = <b>Gross ${fmtMoney(gross)}</b>`;
    return { net, disc, taxable, freight, other, cgst, sgst, igst, gross, type };
  };
  ['ti-client','ti-rate','ti-qty','ti-disc','ti-freight','ti-other'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', preview);
    if (el && el.tagName === 'SELECT') el.addEventListener('change', preview);
  });
  preview();

  document.getElementById('ti-cancel').onclick = closeModal;
  document.getElementById('ti-save').onclick = () => {
    const cid = document.getElementById('ti-client').value;
    const cl = clients.find(c => c.id === cid);
    if (!cl) { toast('Select a client'); return; }
    const desc = document.getElementById('ti-desc').value.trim();
    if (!desc) { toast('Description is required'); return; }
    const calc = preview();
    const rec = {
      id: existingTI ? existingTI.id : uid(),
      invNo: existingTI ? existingTI.invNo : nextTINumber(),
      date: document.getElementById('ti-date').value,
      supplyDate: document.getElementById('ti-supply').value,
      paymentTerms: document.getElementById('ti-terms').value,
      clientId: cid, clientName: cl.companyName, clientShort: cl.shortName,
      desc, unit: document.getElementById('ti-unit').value,
      rate: parseFloat(document.getElementById('ti-rate').value) || 0,
      qty: parseFloat(document.getElementById('ti-qty').value) || 1,
      sac: document.getElementById('ti-sac').value || '998399',
      disc: calc.disc, freight: calc.freight, other: calc.other,
      taxable: calc.taxable, cgst: calc.cgst, sgst: calc.sgst,
      igst: calc.igst, gross: calc.gross, gstType: calc.type,
      status: existingTI ? existingTI.status : 'draft',
      fromPiId: sourcePi ? sourcePi.id : (existingTI ? existingTI.fromPiId : null),
      fromPiNo: sourcePi ? sourcePi.invNo : (existingTI ? existingTI.fromPiNo : null),
    };
    if (existingTI) {
      Object.assign(existingTI, rec);
    } else {
      DB.invoicing.taxInvoices.push(rec);
      // If converting from PI: mark PI as converted, remove PI receivable, add TI receivable
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
function invAddToReceivables(inv, type) {
  const mk = inv.date ? inv.date.slice(0, 7) : todayMonthKey();
  const m = ensureMonthExists(mk);
  m.receivables.push({
    id: uid(),
    name: `${inv.invNo} — ${inv.clientShort || inv.clientName}`,
    amount: inv.gross,
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
  if (!confirm('Cancel this invoice? This will also remove its receivable from the cashflow.')) return;
  invInit();
  if (type === 'pi') {
    const inv = DB.invoicing.proformas.find(p => p.id === invId);
    if (inv) { invRemoveFromReceivables(inv.invNo); inv.status = 'cancelled'; }
  } else {
    const inv = DB.invoicing.taxInvoices.find(t => t.id === invId);
    if (inv) { invRemoveFromReceivables(inv.invNo); inv.status = 'cancelled'; }
  }
  saveDB(); invRenderRegister();
  toast('Invoice cancelled and receivable removed');
}

// ── Email (mailto) ─────────────────────────────────────────
function invEmail(invId, type) {
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;
  const cl = DB.invoicing.clients.find(c => c.id === inv.clientId);
  const to = cl ? cl.email : '';
  const subject = `${inv.invNo} — ${type === 'pi' ? 'Proforma Invoice' : 'Tax Invoice'} from Goru Training Pvt. Ltd.`;
  const body = `Dear ${cl ? cl.attn : 'Sir/Madam'},\n\nPlease find attached ${inv.invNo} dated ${inv.date} for ${inv.desc}.\n\nGross Amount: ₹${inv.gross.toLocaleString('en-IN')}\n\nKindly arrange payment at your earliest convenience.\n\nRegards,\nGoru Training Pvt. Ltd.`;
  window.open(`mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`);
  toast('Email client opened — attach the downloaded Excel before sending');
}

// ── Excel generation ───────────────────────────────────────
async function invDownloadExcel(invId, type) {
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;
  const cl = DB.invoicing.clients.find(c => c.id === inv.clientId) || {};
  toast('Generating Excel…');
  try {
    const XLSX = window.XLSX;
    if (!XLSX) { toast('Excel library not loaded — try refreshing'); return; }
    const wb = XLSX.utils.book_new();
    const isPi = type === 'pi';
    const rows = invBuildSheetData(inv, cl, isPi);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      {wch:4},{wch:30},{wch:10},{wch:8},{wch:10},{wch:12},{wch:8},
      {wch:14},{wch:5},{wch:12},{wch:5},{wch:12},{wch:5},{wch:12},{wch:14}
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Invoice');
    XLSX.writeFile(wb, `${inv.invNo.replace(/\//g,'-')}.xlsx`);
    toast(`${inv.invNo} downloaded`);
  } catch (e) {
    console.error(e);
    toast('Excel generation failed — ' + e.message);
  }
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
  const typeLabel = isPi ? 'Proforma Invoice' : 'Tax Invoice';

  const defaultBody = `Dear ${cl.attn || 'Sir/Madam'},

Please find attached ${inv.invNo} dated ${inv.date} from Goru Training Pvt. Ltd.

Description: ${inv.desc}
Gross Amount: ₹${(inv.gross || 0).toLocaleString('en-IN')}

${isPi ? 'Kindly review and confirm the order.' : 'Kindly arrange payment at your earliest convenience.'}

Regards,
Goru Training Pvt. Ltd.`;

  const subject = `${inv.invNo} — ${typeLabel} from Goru Training Pvt. Ltd.`;

  const body = `
    <div class="hint" style="background:var(--blue-bg); margin-bottom:14px;">
      Step 1 — Click <strong>Download PDF</strong> to save the invoice.<br>
      Step 2 — Click <strong>Open Gmail</strong> to compose the email, then attach the downloaded PDF.
    </div>
    <div class="field">
      <label>To (email)</label>
      <input id="em-to" value="${escapeHtml(cl.email || '')}">
    </div>
    <div class="field">
      <label>Subject</label>
      <input id="em-subject" value="${escapeHtml(subject)}">
    </div>
    <div class="field">
      <label>Email body template</label>
      <select id="em-template" style="margin-bottom:8px;">
        <option value="custom">Custom (type below)</option>
        <option value="pi_standard">PI — Standard text (coming soon)</option>
        <option value="ti_standard">TI — Standard text (coming soon)</option>
      </select>
      <textarea id="em-body" rows="8" style="width:100%; padding:9px 11px; border:1px solid var(--line); border-radius:5px; font-size:13px; font-family:var(--sans); resize:vertical;">${escapeHtml(defaultBody)}</textarea>
    </div>`;

  openModal(`Email ${inv.invNo}`, body,
    `<button class="btn" id="em-cancel">Cancel</button>
     <button class="btn btn-primary" id="em-pdf">&#8659; Download PDF first</button>
     <button class="btn btn-primary" id="em-open-gmail">&#9993; Open Gmail</button>`);

  document.getElementById('em-cancel').onclick = closeModal;

  document.getElementById('em-template').onchange = e => {
    if (e.target.value !== 'custom') {
      toast('Standard templates coming soon — type your message below for now');
      e.target.value = 'custom';
    }
  };

  document.getElementById('em-pdf').onclick = async () => {
    const btn = document.getElementById('em-pdf');
    btn.textContent = 'Generating PDF…';
    btn.disabled = true;
    try {
      await invGeneratePDF(inv, cl, isPi);
      btn.textContent = '✓ PDF downloaded';
      toast(`${inv.invNo}.pdf downloaded — attach it in Gmail`);
    } catch (e) {
      btn.textContent = '⬇ Download PDF first';
      btn.disabled = false;
      toast('PDF generation failed: ' + e.message);
    }
  };

  document.getElementById('em-open-gmail').onclick = () => {
    const to = document.getElementById('em-to').value.trim();
    const subj = document.getElementById('em-subject').value.trim();
    const bodyText = document.getElementById('em-body').value.trim();
    const gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subj)}&body=${encodeURIComponent(bodyText)}`;
    window.open(gmailUrl, '_blank');
    closeModal();
    toast('Gmail opened — attach the downloaded PDF before sending');
  };
}
