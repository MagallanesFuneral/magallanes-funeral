// ============================================================
//  supabase.js — Magallanes Funeral
//  Replaces all localStorage data calls with Supabase.
//  Handles login/logout and exposes window.DB for app.js.
// ============================================================

const SUPABASE_URL = "https://bhpmxlzpuwkfqsqmtczl.supabase.co";
const SUPABASE_KEY = "sb_publishable_Rga_JA13WFbEuDptORwArQ_EcEzh0-G";

const _sb = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
window._sb = _sb;

// ── Table name map ────────────────────────────────────────────
const TABLE = {
  contracts:    "contracts",
  cashReceived: "cash_received",
  cashExpense:  "cash_expense",
  bankReceived: "bank_received",
  bankExpense:  "bank_expense",
  pnbDeposit:   "pnb_deposit",
  settings:     "settings",
};

// ── Field name translators (JS camelCase ↔ DB snake_case) ────

function contractToDb(r) {
  return {
    id: r.id || undefined,
    date: r.date || null,
    contract: r.contract || null,
    deceased: r.deceased || null,
    casket: r.casket || null,
    address: r.address || null,
    amount: Number(r.amount) || 0,
    inhaus: Number(r.inhaus) || 0,
    bai: Number(r.bai) || 0,
    gl: Number(r.gl) || 0,
    gcash: Number(r.gcash) || 0,
    cash: Number(r.cash) || 0,
    discount: Number(r.discount) || 0,
    last_payment: r.lastPayment || r.last_payment || "—",
  };
}
function contractFromDb(r) {
  return {
    id: r.id, date: r.date, contract: r.contract,
    deceased: r.deceased, casket: r.casket, address: r.address,
    amount: Number(r.amount) || 0, inhaus: Number(r.inhaus) || 0,
    bai: Number(r.bai) || 0, gl: Number(r.gl) || 0,
    gcash: Number(r.gcash) || 0, cash: Number(r.cash) || 0,
    discount: Number(r.discount) || 0,
    lastPayment: r.last_payment || "—",
  };
}

function cashToDb(r) {
  return {
    id: r.id || undefined,
    date: r.date || null, contract: r.contract || null,
    receipt: r.receipt || null, client: r.client || null,
    particular: r.particular || null, amount: Number(r.amount) || 0,
  };
}
function cashFromDb(r) {
  return { id: r.id, date: r.date, contract: r.contract,
    receipt: r.receipt, client: r.client,
    particular: r.particular, amount: Number(r.amount) || 0 };
}

function cashExpToDb(r) {
  return { id: r.id || undefined, date: r.date || null,
    particular: r.particular || null, amount: Number(r.amount) || 0 };
}
function cashExpFromDb(r) {
  return { id: r.id, date: r.date,
    particular: r.particular, amount: Number(r.amount) || 0 };
}

function bankToDb(r) {
  return { id: r.id || undefined, date: r.date || null,
    contract: r.contract || null, type: r.type || null,
    client: r.client || null, amount: Number(r.amount) || 0 };
}
function bankFromDb(r) {
  return { id: r.id, date: r.date, contract: r.contract,
    type: r.type, client: r.client, amount: Number(r.amount) || 0 };
}

function bankExpToDb(r) {
  return { id: r.id || undefined, date: r.date || null,
    cv: r.cv || null, check_no: r.check || r.check_no || null,
    particular: r.particular || null, withdraw: Number(r.withdraw) || 0 };
}
function bankExpFromDb(r) {
  return { id: r.id, date: r.date, cv: r.cv,
    check: r.check_no, particular: r.particular,
    withdraw: Number(r.withdraw) || 0 };
}

function pnbToDb(r) {
  return { id: r.id || undefined, date: r.date || null,
    amount: Number(r.amount) || 0 };
}
function pnbFromDb(r) {
  return { id: r.id, date: r.date, amount: Number(r.amount) || 0 };
}

// ── Generic DB helpers ────────────────────────────────────────

async function dbGetAll(table, fromDb) {
  const { data, error } = await _sb.from(table).select("*").order("created_at", { ascending: true });
  if (error) { console.error("dbGetAll", table, error); return []; }
  return (data || []).map(fromDb);
}

async function dbUpsert(table, row, toDb) {
  const payload = toDb(row);
  if (!payload.id) delete payload.id;
  const { data, error } = await _sb.from(table).upsert(payload, { onConflict: "id" }).select().single();
  if (error) { console.error("dbUpsert", table, error); return null; }
  return data;
}

async function dbDelete(table, id) {
  const { error } = await _sb.from(table).delete().eq("id", id);
  if (error) console.error("dbDelete", table, error);
}

// ── Public DB API exposed to app.js ──────────────────────────
window.DB = {
  // Contracts
  async getContracts()          { return dbGetAll(TABLE.contracts, contractFromDb); },
  async saveContract(r)         { return dbUpsert(TABLE.contracts, r, contractToDb); },
  async deleteContract(id)      { return dbDelete(TABLE.contracts, id); },

  // Cash Received
  async getCashReceived()       { return dbGetAll(TABLE.cashReceived, cashFromDb); },
  async saveCashReceived(r)     { return dbUpsert(TABLE.cashReceived, r, cashToDb); },
  async deleteCashReceived(id)  { return dbDelete(TABLE.cashReceived, id); },

  // Cash Expense
  async getCashExpense()        { return dbGetAll(TABLE.cashExpense, cashExpFromDb); },
  async saveCashExpense(r)      { return dbUpsert(TABLE.cashExpense, r, cashExpToDb); },
  async deleteCashExpense(id)   { return dbDelete(TABLE.cashExpense, id); },

  // Bank Received
  async getBankReceived()       { return dbGetAll(TABLE.bankReceived, bankFromDb); },
  async saveBankReceived(r)     { return dbUpsert(TABLE.bankReceived, r, bankToDb); },
  async deleteBankReceived(id)  { return dbDelete(TABLE.bankReceived, id); },

  // Bank Expense
  async getBankExpense()        { return dbGetAll(TABLE.bankExpense, bankExpFromDb); },
  async saveBankExpense(r)      { return dbUpsert(TABLE.bankExpense, r, bankExpToDb); },
  async deleteBankExpense(id)   { return dbDelete(TABLE.bankExpense, id); },

  // PNB Deposit
  async getPnbDeposit()         { return dbGetAll(TABLE.pnbDeposit, pnbFromDb); },
  async savePnbDeposit(r)       { return dbUpsert(TABLE.pnbDeposit, r, pnbToDb); },
  async deletePnbDeposit(id)    { return dbDelete(TABLE.pnbDeposit, id); },

  // Settings (shared — single row, no user_id filter needed with RLS)
  async getSettings() {
    const { data, error } = await _sb.from(TABLE.settings).select("*").limit(1).maybeSingle();
    if (error) { console.error("getSettings", error); return null; }
    return data;
  },
  async saveSettings(obj) {
    // Get existing row id if any
    const existing = await this.getSettings();
    const payload = {
      id: existing?.id || undefined,
      cash_balance:    Number(obj.cashBalance)    || 0,
      bank_balance:    Number(obj.bankBalance)    || 0,
      finance_clerk:   obj.financeClerk   || "Jennifer F. Landicho",
      accountant:      obj.accountant     || "Ranni V. Dalisay",
      finance_manager: obj.financeManager || "June Lizette M. Quizon",
      updated_at:      new Date().toISOString(),
    };
    if (!payload.id) delete payload.id;
    const { error } = await _sb.from(TABLE.settings).upsert(payload, { onConflict: "id" });
    if (error) console.error("saveSettings", error);
  },
};

// ── Auth UI ───────────────────────────────────────────────────

const loginOverlay   = document.getElementById("loginOverlay");
const loginEmail     = document.getElementById("loginEmail");
const loginPassword  = document.getElementById("loginPassword");
const loginError     = document.getElementById("loginError");
const loginSpinner   = document.getElementById("loginSpinner");
const btnLogin       = document.getElementById("btnLogin");
const btnSignOut     = document.getElementById("btnSignOut");
const userBadge      = document.getElementById("userBadge");
const userEmailEl    = document.getElementById("userEmail");

function showLoginError(msg) {
  loginError.textContent = msg;
  loginError.style.display = "block";
}
function hideLoginError() {
  loginError.style.display = "none";
}

function showApp(user) {
  loginOverlay.style.display = "none";
  userBadge.style.display    = "flex";
  userEmailEl.textContent    = user.email;
}

function showLogin() {
  loginOverlay.style.display = "flex";
  userBadge.style.display    = "none";
}

// Handle login button
async function doLogin() {
  const email    = loginEmail.value.trim();
  const password = loginPassword.value;
  if (!email || !password) { showLoginError("Please enter your email and password."); return; }

  hideLoginError();
  btnLogin.disabled    = true;
  loginSpinner.style.display = "block";

  const { data, error } = await _sb.auth.signInWithPassword({ email, password });

  btnLogin.disabled    = false;
  loginSpinner.style.display = "none";

  if (error) { showLoginError(error.message); return; }
  showApp(data.user);
}

btnLogin?.addEventListener("click", doLogin);

// Allow Enter key to submit
loginPassword?.addEventListener("keydown", e => { if (e.key === "Enter") doLogin(); });
loginEmail?.addEventListener("keydown",    e => { if (e.key === "Enter") loginPassword?.focus(); });

// Sign out
btnSignOut?.addEventListener("click", async () => {
  await _sb.auth.signOut();
  showLogin();
  loginEmail.value    = "";
  loginPassword.value = "";
});

// On load — check if already logged in
(async () => {
  const { data: { session } } = await _sb.auth.getSession();
  if (session?.user) {
    showApp(session.user);
  } else {
    showLogin();
  }
})();
