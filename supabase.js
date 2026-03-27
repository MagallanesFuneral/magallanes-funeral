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
  pnbSavings:   "pnb_savings",
  landbank:     "landbank",
  dswd:         "dswd",
  bai:          "bai",
  settings:     "settings",
  contractForms: "contract_forms",
  refundLog:     "refund_log",
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

function pnbSavingsToDb(r) {
  return { id: r.id || undefined, date: r.date || null, amount: Number(r.amount) || 0 };
}
function pnbSavingsFromDb(r) {
  return { id: r.id, date: r.date, amount: Number(r.amount) || 0 };
}
function landbankToDb(r) {
  return { id: r.id || undefined, date: r.date || null, amount: Number(r.amount) || 0 };
}
function landbankFromDb(r) {
  return { id: r.id, date: r.date, amount: Number(r.amount) || 0 };
}

function baiToDb(r) {
  return {
    id: r.id || undefined,
    date_applied:    r.dateApplied    || null,
    contract:        r.contract       || null,
    amount:          Number(r.amount) || 0,
    date_completed:  r.dateCompleted  || null,
    status:          r.status         || "Pending",
  };
}
function baiFromDb(r) {
  return {
    id:            r.id,
    dateApplied:   r.date_applied   || "",
    contract:      r.contract       || "",
    amount:        Number(r.amount) || 0,
    dateCompleted: r.date_completed || "",
    status:        r.status         || "Pending",
  };
}

function dswdToDb(r) {
  return {
    id: r.id || undefined,
    date: r.date || null,
    contract: r.contract || null,
    deceased: r.deceased || null,
    contract_amt: Number(r.contractAmt) || 0,
    payment: Number(r.payment) || 0,
    balance: Number(r.balance) || 0,
    dswd_refund: Number(r.dswdRefund) || 0,
    after_tax: Number(r.afterTax) || 0,
    date_received: r.dateReceived || null,
    payable: Number(r.payable) || 0,
    date_release: r.dateRelease || null,
    beneficiary: r.beneficiary || null,
    dswd_discount: Number(r.dswdDiscount) || 0,
    status: r.status || "Waiting",
  };
}
function dswdFromDb(r) {
  return {
    id: r.id,
    date: r.date || "",
    contract: r.contract || "",
    deceased: r.deceased || "",
    contractAmt: Number(r.contract_amt) || 0,
    payment: Number(r.payment) || 0,
    balance: Number(r.balance) || 0,
    dswdRefund: Number(r.dswd_refund) || 0,
    afterTax: Number(r.after_tax) || 0,
    dateReceived: r.date_received || "",
    payable: Number(r.payable) || 0,
    dateRelease: r.date_release || "",
    beneficiary: r.beneficiary || "",
    dswdDiscount: Number(r.dswd_discount) || 0,
    status: r.status || "Waiting",
  };
}

function contractFormToDb(r) {
  return {
    id:           r.id || undefined,
    form_no:      r.formNo      || null,
    form_date:    r.formDate    || null,
    client_name:  r.clientName  || null,
    address:      r.address     || null,
    contact:      r.contact     || null,
    fb_account:   r.fbAccount   || null,
    deceased1:    r.deceased1   || null,
    deceased2:    r.deceased2   || null,
    svc1:         r.svc1        || null,  svc2: r.svc2 || null,
    svc3:         r.svc3        || null,  svc4: r.svc4 || null,
    svc5:         r.svc5        || null,  svc6: r.svc6 || null,
    svc7:         r.svc7        || null,  svc8: r.svc8 || null,
    amt1:         r.amt1        || null,  amt2:  r.amt2  || null,
    amt3:         r.amt3        || null,  amt4:  r.amt4  || null,
    amt5:         r.amt5        || null,  amt6:  r.amt6  || null,
    amt7:         r.amt7        || null,  amt8:  r.amt8  || null,
    amt9:         r.amt9        || null,  amt10: r.amt10 || null,
    amt11:        r.amt11       || null,  amt12: r.amt12 || null,
    total:        r.total       || null,
    relation:     r.relation    || null,
  };
}
function contractFormFromDb(r) {
  return {
    id:         r.id,
    formNo:     r.form_no     || "",
    formDate:   r.form_date   || "",
    clientName: r.client_name || "",
    address:    r.address     || "",
    contact:    r.contact     || "",
    fbAccount:  r.fb_account  || "",
    deceased1:  r.deceased1   || "",
    deceased2:  r.deceased2   || "",
    svc1:  r.svc1  || "", svc2:  r.svc2  || "",
    svc3:  r.svc3  || "", svc4:  r.svc4  || "",
    svc5:  r.svc5  || "", svc6:  r.svc6  || "",
    svc7:  r.svc7  || "", svc8:  r.svc8  || "",
    amt1:  r.amt1  || "", amt2:  r.amt2  || "",
    amt3:  r.amt3  || "", amt4:  r.amt4  || "",
    amt5:  r.amt5  || "", amt6:  r.amt6  || "",
    amt7:  r.amt7  || "", amt8:  r.amt8  || "",
    amt9:  r.amt9  || "", amt10: r.amt10 || "",
    amt11: r.amt11 || "", amt12: r.amt12 || "",
    total:    r.total    || "",
    relation: r.relation || "",
  };
}

function refundLogToDb(r) {
  return {
    id:           r.id || undefined,
    contract_no:  r.contractNo   || null,
    deceased:     r.deceased     || null,
    amount:       Number(r.amount) || 0,
    date_refunded:r.dateRefunded || null,
    notes:        r.notes        || null,
  };
}
function refundLogFromDb(r) {
  return {
    id:           r.id,
    contractNo:   r.contract_no  || "",
    deceased:     r.deceased     || "",
    amount:       Number(r.amount) || 0,
    dateRefunded: r.date_refunded || "",
    notes:        r.notes        || "",
  };
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

  // DSWD
  async getDswd()               { return dbGetAll(TABLE.dswd, dswdFromDb); },
  async saveDswd(r)             { return dbUpsert(TABLE.dswd, r, dswdToDb); },
  async deleteDswd(id)          { return dbDelete(TABLE.dswd, id); },

  // BAI
  async getBai()                { return dbGetAll(TABLE.bai, baiFromDb); },
  async saveBai(r)              { return dbUpsert(TABLE.bai, r, baiToDb); },
  async deleteBai(id)           { return dbDelete(TABLE.bai, id); },

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
      cash_balance:           Number(obj.cashBalance)          || 0,
      bank_balance:           Number(obj.bankBalance)          || 0,
      pnb_savings_balance:    Number(obj.pnbSavingsBalance)    || 0,
      landbank_balance:       Number(obj.landbankBalance)      || 0,
      finance_clerk:   obj.financeClerk   || "Jennifer F. Landicho",
      accountant:      obj.accountant     || "Ranni V. Dalisay",
      finance_manager: obj.financeManager || "June Lizette M. Quizon",
      updated_at:      new Date().toISOString(),
    };
    if (!payload.id) delete payload.id;
    const { error } = await _sb.from(TABLE.settings).upsert(payload, { onConflict: "id" });
    if (error) console.error("saveSettings", error);
  },

  // Contract Forms
  async getContractForms()        { return dbGetAll(TABLE.contractForms, contractFormFromDb); },
  async saveContractForm(r)       { return dbUpsert(TABLE.contractForms, r, contractFormToDb); },
  async deleteContractForm(id)    { return dbDelete(TABLE.contractForms, id); },

  // PNB Savings
  async getPnbSavings()          { return dbGetAll(TABLE.pnbSavings, pnbSavingsFromDb); },
  async savePnbSavings(r)        { return dbUpsert(TABLE.pnbSavings, r, pnbSavingsToDb); },
  async deletePnbSavings(id)     { return dbDelete(TABLE.pnbSavings, id); },

  // Landbank
  async getLandbank()            { return dbGetAll(TABLE.landbank, landbankFromDb); },
  async saveLandbank(r)          { return dbUpsert(TABLE.landbank, r, landbankToDb); },
  async deleteLandbank(id)       { return dbDelete(TABLE.landbank, id); },

  // Refund Log
  async getRefundLog()           { return dbGetAll(TABLE.refundLog, refundLogFromDb); },
  async saveRefundLog(r)         { return dbUpsert(TABLE.refundLog, r, refundLogToDb); },
  async deleteRefundLog(id)      { return dbDelete(TABLE.refundLog, id); },

  // ── Delete All (Fresh Restart) ────────────────────────────
  async deleteAllContracts()    { const { error } = await _sb.from(TABLE.contracts).delete().gte("id", "00000000-0000-0000-0000-000000000000");    if (error) throw error; },
  async deleteAllCashReceived() { const { error } = await _sb.from(TABLE.cashReceived).delete().gte("id", "00000000-0000-0000-0000-000000000000"); if (error) throw error; },
  async deleteAllCashExpense()  { const { error } = await _sb.from(TABLE.cashExpense).delete().gte("id", "00000000-0000-0000-0000-000000000000");  if (error) throw error; },
  async deleteAllBankReceived() { const { error } = await _sb.from(TABLE.bankReceived).delete().gte("id", "00000000-0000-0000-0000-000000000000"); if (error) throw error; },
  async deleteAllBankExpense()  { const { error } = await _sb.from(TABLE.bankExpense).delete().gte("id", "00000000-0000-0000-0000-000000000000");  if (error) throw error; },
  async deleteAllPnbDeposit()   { const { error } = await _sb.from(TABLE.pnbDeposit).delete().gte("id", "00000000-0000-0000-0000-000000000000");   if (error) throw error; },
  async deleteAllDswd()         { const { error } = await _sb.from(TABLE.dswd).delete().gte("id", "00000000-0000-0000-0000-000000000000");          if (error) throw error; },
  async deleteAllBai()          { const { error } = await _sb.from(TABLE.bai).delete().gte("id", "00000000-0000-0000-0000-000000000000");           if (error) throw error; },
  async deleteAllSettings()     { const { error } = await _sb.from(TABLE.settings).delete().gte("id", "00000000-0000-0000-0000-000000000000");     if (error) throw error; },
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
