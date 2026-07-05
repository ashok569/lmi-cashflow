/* ===========================================================
   LMI Cashflow Manager — application logic
   VERSION 2.3.7 — fix: Word download uses Packer.toBlob (browser-compatible) instead of Packer.toBuffer (Node-only); fix dataset.invWord reference; invBuildWordDoc returns Document not Buffer. — edit PI/TI syncs cashflow receivable amount; cancel vs permanent delete modal; invUpdateReceivableAmount() — header updated to match actual LMI India letterhead (LMI INDIA branding, Apeejay House address, CIN, email/web/tel, logo placeholder), footer updated.
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
          <button data-inv-word="${inv.id}" data-inv-type="${isPi ? 'pi' : 'ti'}" title="Download Word">&#128229;</button>
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
      const oldGross = existing.gross;
      Object.assign(existing, rec);
      // If amount changed, update the linked cashflow receivable
      if (oldGross !== rec.gross) {
        invUpdateReceivableAmount(rec.invNo, rec.gross);
      }
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
  // If called from dashboard, open the invoicing module first
  if (fromDashboard) {
    const overlay = document.getElementById('invoicingOverlay');
    if (overlay && overlay.style.display === 'none') openInvoicingModule();
  }
  const existingTI = editTiId ? DB.invoicing.taxInvoices.find(t => t.id === editTiId) : null;
  const sourcePi = fromPiId ? DB.invoicing.proformas.find(p => p.id === fromPiId) : null;
  const v = existingTI || sourcePi || {};
  const clients = DB.invoicing.clients;
  const today = new Date().toISOString().slice(0, 10);
  const fy = currentInvFY();
  const nextTiNum = `TI/${fy}/${String((DB.invoicing.tiSequence[fy]||0)+1).padStart(3,'0')}`;

  if (!clients.length) {
    toast('No clients found — please import or add clients first (Invoicing tab → Clients)');
    return;
  }

  // If converting from PI, pre-select that client and pre-fill
  const preClient = v.clientId || '';
  const body = `
    <div class="hint" style="background:#e8f0fb; margin-bottom:12px;">
      ${sourcePi ? `Converting from ${escapeHtml(sourcePi.invNo)} → ` : ''}
      ${existingTI ? `Editing ${escapeHtml(existingTI.invNo)}` : `Next TI number will be: <strong>${nextTiNum}</strong>`}
    </div>
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
      const oldGross = existingTI.gross;
      Object.assign(existingTI, rec);
      // If amount changed, update the linked cashflow receivable
      if (oldGross !== rec.gross) {
        invUpdateReceivableAmount(rec.invNo, rec.gross);
      }
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
  invInit();
  const inv = type === 'pi'
    ? DB.invoicing.proformas.find(p => p.id === invId)
    : DB.invoicing.taxInvoices.find(t => t.id === invId);
  if (!inv) return;

  const body = `
    <div class="field">
      <label>What would you like to do with ${escapeHtml(inv.invNo)}?</label>
    </div>
    <div class="hint" style="margin-top:8px;">
      <strong>Cancel</strong> — marks the invoice as cancelled and removes its receivable from the cashflow, but keeps it in the register for audit trail.<br><br>
      <strong>Permanently delete</strong> — removes the invoice from the register entirely. Use only for test entries or duplicates.
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
    if (type === 'pi') {
      DB.invoicing.proformas = DB.invoicing.proformas.filter(p => p.id !== invId);
    } else {
      DB.invoicing.taxInvoices = DB.invoicing.taxInvoices.filter(t => t.id !== invId);
    }
    saveDB(); closeModal(); invRenderRegister();
    toast(`${inv.invNo} permanently deleted`);
  };
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
  const invLabel = isPi ? 'Proforma Invoice No :-' : 'Tax Invoice No :-';
  const dateLabel = isPi ? 'Proforma Invoice Date :-' : 'Tax Invoice Date :-';

  // Column widths (DXA) matching your 15-column template
  const COLS = [270, 1742, 614, 500, 614, 728, 500, 840, 332, 728, 332, 728, 332, 728, 838];
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

  // ── From/To header row ──
  const fromW = COLS[0]+COLS[1]+COLS[2]+COLS[3]+COLS[4];
  const toW   = COLS[5]+COLS[6]+COLS[7]+COLS[8]+COLS[9]+COLS[10];
  const lbW   = COLS[11]+COLS[12]+COLS[13];

  const bL = (t=false,b=false,l=true,r=false) => ({
    top: t?thinBorder:noBorder, bottom: b?thinBorder:noBorder,
    left: l?thinBorder:noBorder, right: r?thinBorder:noBorder
  });

  function addrRow2(fromTxt, toTxt, lbTxt, valTxt, fromBold=false) {
    return new TableRow({ children: [
      new TableCell({ width:{size:fromW,type:WidthType.DXA}, columnSpan:5,
        borders:bL(false,false,true,false), margins:{top:30,bottom:30,left:80,right:80},
        children:[para(run2(fromTxt||'',{bold:fromBold,size:18}))] }),
      new TableCell({ width:{size:toW,type:WidthType.DXA}, columnSpan:6,
        borders:bL(false,false,true,false), margins:{top:30,bottom:30,left:80,right:80},
        children:[para(run2(toTxt||'',{size:18}))] }),
      new TableCell({ width:{size:lbW,type:WidthType.DXA}, columnSpan:3,
        borders:bL(false,false,true,false), margins:{top:30,bottom:30,left:80,right:80},
        children:[para(run2(lbTxt||'',{bold:true,size:17}))] }),
      new TableCell({ width:{size:COLS[14],type:WidthType.DXA},
        borders:bL(false,false,true,true), margins:{top:30,bottom:30,left:80,right:80},
        children:[para(run2(valTxt||'',{size:17}))] }),
    ]});
  }

  function lvRow(fl,fv,tl,tv,rl,rv) {
    function lvCell(lbl,val,w,cs) {
      return new TableCell({ width:{size:w,type:WidthType.DXA}, columnSpan:cs||1,
        borders:bL(false,false,true,false), margins:{top:30,bottom:30,left:80,right:80},
        children:[para([run2(lbl,{bold:true,size:17}),run2(' '+(val||''),{size:17})])] });
    }
    return new TableRow({ children:[
      lvCell(fl,fv,fromW,5), lvCell(tl,tv,toW,6), lvCell(rl,rv,lbW,3),
      new TableCell({ width:{size:COLS[14],type:WidthType.DXA},
        borders:bL(false,false,true,true),
        children:[new Paragraph({children:[]})] })
    ]});
  }

  // ── From/To first row with inv number ──
  const fromHdrRow = new TableRow({ children:[
    new TableCell({ width:{size:fromW,type:WidthType.DXA}, columnSpan:5,
      borders:bL(true,false,true,false), margins:{top:60,bottom:20,left:80,right:80},
      children:[para(run2('From,',{bold:true,size:19}))] }),
    new TableCell({ width:{size:toW,type:WidthType.DXA}, columnSpan:6,
      borders:bL(true,false,true,false), margins:{top:60,bottom:20,left:80,right:80},
      children:[para(run2('To,',{bold:true,size:19}))] }),
    new TableCell({ width:{size:lbW,type:WidthType.DXA}, columnSpan:3,
      borders:bL(true,false,true,false), margins:{top:60,bottom:20,left:80,right:80},
      children:[para(run2(invLabel,{bold:true,size:17}))] }),
    new TableCell({ width:{size:COLS[14],type:WidthType.DXA},
      borders:bL(true,false,true,true), margins:{top:60,bottom:20,left:80,right:80},
      children:[para(run2(inv.invNo,{bold:true,size:17,color:NAVY}))] }),
  ]});

  // ── Line item rows ──
  const lineHdr1 = new TableRow({ children:[
    hdrCell('SR',COLS[0]), hdrCell('DESCRIPTION',COLS[1]),
    hdrCell('HNS/SAC CODE',COLS[2]), hdrCell('UNIT',COLS[3]),
    hdrCell('RATE',COLS[4]), hdrCell('NET AMOUNT',COLS[5]),
    hdrCell('DISC',COLS[6]), hdrCell('TAXABLE VALUE',COLS[7]),
    hdrCell('CGST',COLS[8]+COLS[9],2), hdrCell('SGST',COLS[10]+COLS[11],2),
    hdrCell('IGST',COLS[12]+COLS[13],2), hdrCell('GROSS AMOUNT',COLS[14]),
  ]});

  function navyEmpty(w) {
    return new TableCell({ width:{size:w,type:WidthType.DXA},
      shading:{fill:NAVY,type:ShadingType.CLEAR,color:NAVY},
      borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:thinBorder},
      children:[new Paragraph({children:[]})] });
  }

  const lineHdr2 = new TableRow({ children:[
    ...COLS.slice(0,8).map(w => navyEmpty(w)),
    hdrCell('%',COLS[8]), hdrCell('AMOUNT',COLS[9]),
    hdrCell('%',COLS[10]), hdrCell('AMOUNT',COLS[11]),
    hdrCell('%',COLS[12]), hdrCell('AMOUNT',COLS[13]),
    navyEmpty(COLS[14]),
  ]});

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
  const rowA = lineRow2('A',inv.desc,inv.sac,inv.unit,inv.rate,net,inv.disc||0,inv.taxable,
    intra?0.09:0, inv.cgst||0, intra?0.09:0, inv.sgst||0,
    intra?0:0.18, inv.igst||0,
    inv.gross - (inv.freight||0)*(intra?1:1.18) - (inv.other||0)*(intra?1:1.18));

  const rowB = lineRow2('B','Freight charges','','','','','',(inv.freight||0),
    0,0,0,0, intra?0:0.18, (inv.freight||0)*0.18, (inv.freight||0)*(intra?1:1.18));

  const rowC = lineRow2('C','Other Taxable Charges','','','','','',(inv.other||0),
    0,0,0,0, intra?0:0.18, (inv.other||0)*0.18, (inv.other||0)*(intra?1:1.18));

  const totalRow2 = lineRow2('TOTAL','','','','','','',(inv.taxable||0),
    0,inv.cgst||0, 0,inv.sgst||0, 0,inv.igst||0, inv.gross||0, LIGHT);

  // ── Words row ──
  const wordsRow2 = new TableRow({ children:[new TableCell({
    columnSpan:15, width:{size:TOTAL,type:WidthType.DXA},
    borders:{top:thinBorder,bottom:noBorder,left:thinBorder,right:thinBorder},
    margins:{top:80,bottom:80,left:100,right:100},
    children:[para([run2('Gross Amount in Words :- ',{bold:true,size:18}),
      run2(numberToWords(Math.round(inv.gross||0)),{size:18,italics:true})])]
  })]});

  function wideRow2(text, bold=false, bBot=false) {
    return new TableRow({ children:[new TableCell({
      columnSpan:15, width:{size:TOTAL,type:WidthType.DXA},
      borders:{top:noBorder,bottom:bBot?thinBorder:noBorder,left:thinBorder,right:thinBorder},
      margins:{top:30,bottom:30,left:100,right:100},
      children:[para(run2(text,{bold,size:17}))]
    })]});
  }

  function splitRow2(leftTxt, rightTxt) {
    const lw = TOTAL - COLS[11]-COLS[12]-COLS[13]-COLS[14];
    const rw = COLS[11]+COLS[12]+COLS[13]+COLS[14];
    return new TableRow({ children:[
      new TableCell({ width:{size:lw,type:WidthType.DXA}, columnSpan:11,
        borders:{top:noBorder,bottom:noBorder,left:thinBorder,right:noBorder},
        margins:{top:60,bottom:60,left:100,right:100},
        children:[para(run2(leftTxt,{size:17}))] }),
      new TableCell({ width:{size:rw,type:WidthType.DXA}, columnSpan:4,
        borders:{top:noBorder,bottom:noBorder,left:noBorder,right:thinBorder},
        margins:{top:60,bottom:60,left:80,right:80},
        children:[para(run2(rightTxt,{bold:true,size:18,color:NAVY}),AlignmentType.RIGHT)] }),
    ]});
  }

  const sigRow2 = new TableRow({ children:[
    new TableCell({ width:{size:TOTAL-COLS[11]-COLS[12]-COLS[13]-COLS[14],type:WidthType.DXA}, columnSpan:11,
      borders:{top:noBorder,bottom:thinBorder,left:thinBorder,right:noBorder},
      margins:{top:120,bottom:120,left:100,right:100},
      children:[new Paragraph({children:[]})] }),
    new TableCell({ width:{size:COLS[11]+COLS[12]+COLS[13]+COLS[14],type:WidthType.DXA}, columnSpan:4,
      borders:{top:noBorder,bottom:thinBorder,left:noBorder,right:thinBorder},
      margins:{top:120,bottom:120,left:80,right:80},
      children:[para(run2('AUTHORISED SIGNATORY',{bold:true,size:18,color:NAVY}),AlignmentType.RIGHT)] }),
  ]});

  const table = new Table({
    width:{size:TOTAL,type:WidthType.DXA}, columnWidths:COLS,
    rows:[
      titleRow, emptyRow(), fromHdrRow,
      addrRow2(GORU.name, cl.companyName||'', dateLabel, inv.date||'', true),
      addrRow2(GORU.addr1, cl.addr1||'', 'Date of Supply :-', inv.supplyDate||''),
      addrRow2(GORU.addr2, cl.addr2||'', 'Place of Supply :-', cl.state||''),
      lvRow('State :-',GORU.state,'State :-',cl.state||'','Kind Attn:',cl.attn||''),
      lvRow('GSTIN :-',GORU.gstin,'GSTIN :-',cl.gstin||'','Mob :-',cl.mobile||''),
      lvRow('PAN :-',GORU.pan,'PAN :-',cl.pan||'','Email:-',cl.email||''),
      lvRow('TAN :-',GORU.tan,'TAN :-',cl.tan||'','',''),
      emptyRow(), lineHdr1, lineHdr2,
      rowA, rowB, rowC, totalRow2,
      emptyRow(), wordsRow2, emptyRow(), emptyRow(),
      splitRow2(`Payment Terms :  ${inv.paymentTerms||''}`, 'FOR GORU TRAINING PRIVATE LIMITED'),
      wideRow2('All payments by bank transfer/draft/cheque payable at Mumbai in favour of "Goru Training Pvt. Ltd."'),
      wideRow2(`Bank Name: ${GORU.bank}    Branch: ${GORU.branch}`),
      wideRow2(`A/c. No.: ${GORU.acno};   IFSC Code: ${GORU.ifsc};`),
      emptyRow(), sigRow2,
    ]
  });

  const { Header, Footer } = D;

  // ── Header: company letterhead ──
  // ── Header: LMI India letterhead (matching uploaded sample) ──
  const LMI_BLUE = '1F5FA6';
  // LMI logo embedded as base64 — auto-appears in every generated invoice header
  const LMI_LOGO_B64 = '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6b4j6w9rd69okM0sluXQoUeVCTGSWIwccnFfYFFAGbouiaZ4V0e303SNPtdJ0q0TbBZ2MCwwwqOcIigBR14Ar48+EXwD1r4+eIfHPxj1Tx18QPh1eeKdXktdOsPD91DYn+ybMmC1MqTQO25issmMgfvAcc5P2rRQB+b3xI8AeIPhT+zX+0p8DoYNd8R2NjHba74Z1C5gM897a3c8TTxbo0Cs8c6SEhVH+sziuw+PngHxN/wuiHwHo+j3s/hP4ywaOut30ETGKx+wSKb0yHGFM1oqR89Spr7xooA+J/8AhGdUj139t0JpV4sWoadbR2G23bFzjSZV2w8fPhiF+XPJxXJ/BPx9+zb4f+DPgLS/FPwpWXxNY6BYW2qyXHw0nnka7S3jWctJ9lO9jIGy2TnrzX6CUUAfGfxp8N+Gfjj8Tv2VWs/CTar8PZLrWFk0690d4re3gWwKxLLA6DylDxqFVlA+VcVe8afCLSv2R/i54e+KPw78I29t4N1aWPQ/F+jaTp4draORtsGowKqlk8tyFkVOGVs4yN1fX1FAHzBZ+FJ9c/bq+IbXljc/2FqPw5srBrvymEMha6mDor427trdOvIrlvgX8edN/Zd+GemfCr4q6dr2ja94UVtMsby00O7vbTWrZGb7PNbSQRuNzR7QUYghgQa+yKKAPjL4kfEL4sfEv9nrWLC98O3PhvUviP4iXw74bsFs3W+0zR5yqyXV9yQjmJZmxgbd6A5PSn8Xv2HdR8NfB77R4R+KPxE13V/BUcOseG9D1jUbaayFxZjdDGIktkb7itGu1h9/uMg/bNFAHN/Dvxenj/wH4f8AEiWs9iNVsYbw2tzG0csDOgZo3UjIZSSDn0rpKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAK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';
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
                  run2('LMI INDIA', { bold: true, size: 34, color: LMI_BLUE }),
                  run2('  -  ', { size: 24, color: LMI_BLUE }),
                  run2('A Division of Goru Training Pvt. Ltd.', { bold: true, italics: true, size: 24, color: LMI_BLUE }),
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

  const doc = new Document({
    sections:[{
      properties:{ page:{ size:{width:11906,height:16838},
        margin:{top:1800,right:850,bottom:1200,left:850} } },
      headers: { default: docHeader },
      footers: { default: docFooter },
      children:[table]
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
      Step 1 — Click <strong>Download Word doc</strong> to save the invoice.<br>
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
     <button class="btn btn-primary" id="em-pdf">&#8659; Download Word doc first</button>
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
    btn.textContent = 'Generating Word doc…';
    btn.disabled = true;
    try {
      await invDownloadWord(inv.id, isPi ? 'pi' : 'ti');
      btn.textContent = '✓ Word doc downloaded';
      toast(`${inv.invNo}.docx downloaded — attach it in Gmail`);
    } catch (e) {
      btn.textContent = '⬇ Download Word doc first';
      btn.disabled = false;
      toast('Word generation failed: ' + e.message);
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
