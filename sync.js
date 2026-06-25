/* ===========================================================
   LMI Cashflow Manager — Supabase sync layer  VERSION 2.2.1
   Handles auth, cloud load/save, and realtime multi-device sync.
   Loaded BEFORE app.js. Exposes window.Cloud.
   =========================================================== */

const SUPABASE_URL = window.__SUPABASE_URL__ || '';
const SUPABASE_ANON_KEY = window.__SUPABASE_ANON_KEY__ || '';

let sb = null;                  // supabase client
let currentUser = null;         // { id, email }
let realtimeChannel = null;
let lastLocalWriteAt = {};      // month_key -> timestamp, to ignore our own realtime echoes briefly

function cloudConfigured() {
  return !!(SUPABASE_URL && SUPABASE_ANON_KEY);
}

function initSupabaseClient() {
  if (!cloudConfigured()) return null;
  sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  return sb;
}

/* ---------- Auth ---------- */

async function signIn(email, password) {
  const { data, error } = await sb.auth.signInWithPassword({ email, password });
  if (error) throw error;
  await ensureAppUserRow(data.user);
  return data.user;
}

async function signUpFirstUser(email, password) {
  const { data, error } = await sb.auth.signUp({ email, password });
  if (error) throw error;
  if (data.user && data.session) await ensureAppUserRow(data.user);
  return { user: data.user, session: data.session };
}

// Used by "Add user" — creates a teammate's account. Supabase will sign the
// CURRENT browser session over to the new account on signUp, so we restore
// Ashok's session immediately after.
async function addTeammate(email, password, displayName) {
  const { data: { session: adminSession } } = await sb.auth.getSession();
  const { data, error } = await sb.auth.signUp({ email, password, options: { data: { display_name: displayName || '' } } });
  if (error) throw error;
  const hadSession = !!data.session;
  if (adminSession) {
    await sb.auth.setSession({ access_token: adminSession.access_token, refresh_token: adminSession.refresh_token });
  }
  if (data.user) {
    try { await ensureAppUserRow(data.user); } catch (e) { /* RLS may block this until they first sign in themselves; non-fatal */ }
  }
  return { user: data.user, hadSession };
}

async function ensureAppUserRow(user) {
  if (!user) return;
  try {
    const { data: existing, error: selErr } = await sb.from('app_users').select('id').eq('id', user.id).maybeSingle();
    if (selErr) { console.warn('app_users lookup failed (non-fatal):', selErr); return; }
    if (!existing) {
      const { error: insErr } = await sb.from('app_users').insert({
        id: user.id,
        email: user.email,
        display_name: (user.user_metadata && user.user_metadata.display_name) || user.email.split('@')[0],
      });
      if (insErr) console.warn('app_users insert failed (non-fatal):', insErr);
    }
  } catch (err) {
    console.warn('ensureAppUserRow failed (non-fatal):', err);
  }
}

async function signOut() {
  await sb.auth.signOut();
  currentUser = null;
}

function currentUserLabel() {
  if (!currentUser) return '';
  return (currentUser.displayName || currentUser.email || '').split('@')[0];
}

/* ---------- Cloud load ---------- */

async function cloudLoadAll() {
  const { data: monthsRows, error: e1 } = await sb.from('months').select('*');
  if (e1) throw e1;
  const { data: wsRows, error: e2 } = await sb.from('workspace').select('*').limit(1);
  if (e2) throw e2;

  const ws = wsRows && wsRows[0];
  const db = {
    months: {},
    fixedDeposits: (ws && ws.fixed_deposits) || [],
    bankGuarantees: (ws && ws.bank_guarantees) || [],
    standardVendors: (ws && ws.standard_vendors) || [],
    recurringTemplate: (ws && ws.recurring_template) || [],
    pendingActions: (ws && ws.pending_actions) || { NIRALI: [], ASHOK: [], SANDEEP: [], COMPLETED: [] },
    currentFY: (ws && ws.current_fy) || '26-27',
    selectedMonth: null,
    _workspaceId: ws && ws.id,
  };
  (monthsRows || []).forEach(row => {
    db.months[row.month_key] = rowToMonth(row);
  });
  return db;
}

function rowToMonth(row) {
  const payments = row.payments || [];
  return {
    opening: Number(row.opening) || 0,
    openingManual: !!row.opening_manual,
    hdfc: Number(row.hdfc) || 0,
    yesbank: Number(row.yesbank) || 0,
    receipts: row.receipts || [],
    payments,
    receivables: row.receivables || [],
    imports: row.imports || [],
    // Any month already in the database with payments is treated as provisioned —
    // prevents the self-healer from re-adding items the user deliberately deleted.
    _provisioned: payments.length > 0,
  };
}

function monthToRow(mk, m) {
  return {
    month_key: mk,
    opening: m.opening || 0,
    opening_manual: !!m.openingManual,
    hdfc: m.hdfc || 0,
    yesbank: m.yesbank || 0,
    receipts: m.receipts || [],
    payments: m.payments || [],
    receivables: m.receivables || [],
    imports: m.imports || [],
    updated_at: new Date().toISOString(),
    updated_by: currentUserLabel(),
  };
}

/* ---------- Cloud save (debounced per-month) ---------- */

const _pendingMonths = new Set();
let _pendingWorkspace = false;
let _saveTimer = null;

function queueCloudSave(touchedMonthKey) {
  if (touchedMonthKey) _pendingMonths.add(touchedMonthKey);
  else _pendingWorkspace = true;
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(flushCloudSave, 450);
}

async function flushCloudSave() {
  if (!sb || !currentUser) return;
  const months = Array.from(_pendingMonths);
  _pendingMonths.clear();
  const doWorkspace = _pendingWorkspace;
  _pendingWorkspace = false;

  try {
    for (const mk of months) {
      if (!DB.months[mk]) continue;
      lastLocalWriteAt[mk] = Date.now();
      const row = monthToRow(mk, DB.months[mk]);
      const { error } = await sb.from('months').upsert(row, { onConflict: 'month_key' });
      if (error) console.error('Month save failed', mk, error);
    }
    if (doWorkspace) {
      lastLocalWriteAt['__workspace__'] = Date.now();
      const payload = {
        fixed_deposits: DB.fixedDeposits,
        bank_guarantees: DB.bankGuarantees,
        standard_vendors: DB.standardVendors,
        recurring_template: DB.recurringTemplate,
        pending_actions: DB.pendingActions || {},
        current_fy: DB.currentFY,
        updated_at: new Date().toISOString(),
        updated_by: currentUserLabel(),
      };
      if (DB._workspaceId) {
        const { error } = await sb.from('workspace').update(payload).eq('id', DB._workspaceId);
        if (error) console.error('Workspace save failed', error);
      }
    }
    setSyncStatus('synced');
  } catch (err) {
    console.error('Cloud save error', err);
    setSyncStatus('error');
  }
}

/* ---------- Realtime ---------- */

function startRealtime() {
  if (!sb) return;
  realtimeChannel = sb.channel('cashflow-sync')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'months' }, handleMonthChange)
    .on('postgres_changes', { event: '*', schema: 'public', table: 'workspace' }, handleWorkspaceChange)
    .subscribe(status => {
      if (status === 'SUBSCRIBED') setSyncStatus('synced');
      if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') setSyncStatus('error');
      if (status === 'CLOSED') setSyncStatus('offline');
    });
}

function handleMonthChange(payload) {
  const row = payload.new;
  if (!row || !row.month_key) return;
  // ignore the echo of our own very-recent write
  const since = Date.now() - (lastLocalWriteAt[row.month_key] || 0);
  if (since < 1200) return;
  DB.months[row.month_key] = rowToMonth(row);
  renderAll();
  setSyncStatus('synced');
}

function handleWorkspaceChange(payload) {
  const row = payload.new;
  if (!row) return;
  const since = Date.now() - (lastLocalWriteAt['__workspace__'] || 0);
  if (since < 1200) return;
  DB.fixedDeposits = row.fixed_deposits || [];
  DB.bankGuarantees = row.bank_guarantees || [];
  DB.standardVendors = row.standard_vendors || [];
  DB.recurringTemplate = row.recurring_template || [];
  DB.pendingActions = row.pending_actions || { NIRALI: [], ASHOK: [], SANDEEP: [], COMPLETED: [] };
  DB.currentFY = row.current_fy || DB.currentFY;
  renderAll();
  // If Pending Actions board is open, re-render it too
  if (document.getElementById('pendingActionsOverlay') &&
      document.getElementById('pendingActionsOverlay').style.display !== 'none') {
    if (typeof paRender === 'function') paRender();
  }
  setSyncStatus('synced');
}

function setSyncStatus(status) {
  const el = document.getElementById('syncStatus');
  if (!el) return;
  const map = {
    synced: ['&#9679;', 'Synced', 'ok'],
    saving: ['&#9679;', 'Saving&hellip;', 'busy'],
    error: ['&#9679;', 'Sync error', 'err'],
    offline: ['&#9679;', 'Offline', 'err'],
  };
  const [dot, label, cls] = map[status] || map.offline;
  el.innerHTML = `${dot} ${label}`;
  el.className = `sync-status ${cls}`;
}

window.Cloud = {
  cloudConfigured, initSupabaseClient, signIn, signUpFirstUser, addTeammate, signOut,
  cloudLoadAll, queueCloudSave, startRealtime, setSyncStatus,
  get currentUser() { return currentUser; },
  set currentUser(u) { currentUser = u; },
  get client() { return sb; },
};
