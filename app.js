/* ===========================================================
   LMI Cashflow Manager — application logic
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
    .filter(r => (Number(r.amount) || 0) > 0)
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
  // Only auto-heal months that are strictly after today. Past months and today's
  // month are never healed on navigation — new recurring items are only pushed
  // to them if you explicitly open Edit Recurring from that month and save.
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
  wrap.innerHTML = months.map(mk => {
    const active = mk === DB.selectedMonth ? 'active' : '';
    const isCurrent = mk === tmk ? 'is-current' : '';
    const hasData = !!DB.months[mk];
    return `<button class="month-tab ${active} ${isCurrent} ${hasData ? '' : 'future'}" data-mk="${mk}">${monthLabel(mk)}${isCurrent ? '<span class="dot"></span>' : ''}</button>`;
  }).join('');
  wrap.querySelectorAll('.month-tab').forEach(btn => {
    btn.onclick = () => {
      const touched = selectMonth(btn.dataset.mk);
      saveDB(touched);
      renderAll();
    };
  });
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
  const foot = `<button class="btn" id="ob-cancel">Cancel</button><button class="btn btn-primary" id="ob-save">Save</button>`;
  openModal('Update opening balance', body, foot);
  document.getElementById('ob-cancel').onclick = closeModal;

  document.getElementById('ob-month').addEventListener('change', e => {
    const mk = e.target.value;
    document.getElementById('ob-amount').value = getOpening(mk) || '';
    document.getElementById('ob-hint').textContent = openingHintText(mk);
  });

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

  // Recurring (starred) payments — Salaries, Reimbursement, etc. — were previously only
  // populated via the "Edit recurring" save action, so any month nobody had opened that
  // editor for after stayed empty. Apply the standing recurring template here too, so every
  // month gets these automatically the first time it's created.
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
  syncCarriedReceivables();
  saveDB([nextMonthKey(todayMonthKey())]);
  wireActionBar();
  wireStaticButtons();
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
