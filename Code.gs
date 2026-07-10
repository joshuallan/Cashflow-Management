/**
 * CASHFLOW INTELLIGENCE — FINANCIAL COMMAND CENTER (Code.gs)  v8
 * ----------------------------------------------------------------
 * SHEET LAYOUT (this version): Headers on ROW 1, data starts ROW 2,
 * on both the Ledger and the Backlog sheet. All indexing and the
 * running-balance formula anchor derive from HEADER_ROWS below.
 *
 * NOTE ON THE OPENING BALANCE: the running-balance formula anchors the
 * FIRST data row to cell P1 (R1C16), mirroring your original formula's
 * P$1 anchor. Keep your opening balance in P1 (it can share the header
 * cell via a note, or move the constant below if you relocate it).
 *
 * Engines in this version:
 *   • 3-State Lifecycle: Planned (TX=F,OK=F) → Inputted (TX=T,OK=F)
 *     → Settled (TX=T,OK=T, numeric P)
 *   • Chronological Re-Stitching + global column P formula repair
 *   • Shadow IDs (Ledger col Q) — minted ONLY when a row is created by
 *     the app or when a Backlog item hits the calendar; manual Backlog
 *     rows stay ID-less
 *   • Backlog Engine: read, add, EDIT, drag-reorder, bidirectional
 *     basket (Backlog→Ledger deletes from Backlog; Ledger→Backlog
 *     re-adds it)
 *   • Bidirectional Smart Memory quads {beneficiary, recipient,
 *     destination, account}
 *   • High-Velocity trace prefix intelligence per Expense Category
 *
 * Ledger columns:
 *   A TX  B OK  C Day  D Date  E Trace  F Beneficiary  G Recipient
 *   H Destination  I Account#  J Expense Category  K Details  L Type
 *   M Amount  N Withdrawal  O Deposit  P Balance (formula)  Q Shadow UID
 * Backlog sheet: identical A–Q, plus R Priority (Low/Med/High).
  *
 * SECURITY (v9): every readable/mutating RPC takes an `auth` object
 * ({user, hash}) as its FIRST parameter and is gated server-side by
 * verifySession(auth.user, auth.hash) — defined in Auth.gs — before
 * touching any sheet. The client UI is never trusted: a failed or
 * missing session throws immediately and nothing is read or written.
 * Editor-run maintenance (restitchLedger, backfillUids) and private
 * helpers (readBacklog_, readChecks_, findCheckRow_, findBacklogRow_,
 * findRowByUid) remain un-gated: they are invoked from the script
 * editor or from functions that have already been verified.
 */

const SHEET_NAME = '2026';
const BACKLOG_SHEET = 'Backlog';
const HEADER_ROWS = 2;                       // ← headers on row 2, data from row 3
const FIRST_DATA_ROW = HEADER_ROWS + 1;

const COL = {
  TX: 1, OK: 2, DAY: 3, DATE: 4, TRACE: 5,
  ACTUAL_PAYEE: 6, PAYEE: 7, DEST: 8, ACCT: 9,
  EXPENSE: 10, DETAILS: 11, TYPE: 12,
  AMOUNT: 13, WITHDRAWAL: 14, DEPOSIT: 15, BALANCE: 16,
  UID: 17, PRIORITY: 18
};
const NUM_COLS = 17;   // Ledger width
const BL_COLS = 18;    // Backlog width (adds Priority)


/* ---- Check Registry (PDC) module ----
   The Checks sheet keeps its own layout: headers on ROW 1, data ROW 2. */
const CHECKS_SHEET = 'Checks';
const CHK_HEADER_ROWS = 1;
const CHK_FIRST_DATA_ROW = 2;
const CHK_COL = {
  DATE_ISSUED: 1, BANK: 2, ACCT_NUM: 3, CHECK_NUM: 4,
  BATCH: 5, PAYEE: 6, CATEGORY: 7, MEMO: 8,
  AMOUNT: 9, STATUS: 10, DAYS_DUE: 11,
  TARGET_DATE: 12, DATE_CLEARED: 13
};
const CHK_NUM_COLS = 13;
const CHK_WINDOW_DAYS = 30;   // pending checks due within this horizon (plus overdue)


/** Your running-balance formula in R1C1, with the anchor row derived
 *  from FIRST_DATA_ROW so a future layout change is a one-line edit. */
const BALANCE_FORMULA_R1C1 =
  '=IF(RC2=FALSE,"",IF(AND(ISBLANK(RC14),ISBLANK(RC15)),"",IF(ROW()=' + FIRST_DATA_ROW + ',R1C16,R[-1]C) - RC14 + RC15))';

/* ------------------------------------------------------------------ */
/* doGet stays open: it serves the app shell containing the LOGIN GATE.
   No financial data ships with the page — the first byte of ledger data
   only leaves the server after verifySession passes inside getData. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cashflow Intelligence — Financial Command Center')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------ */
/*  READ — Ledger + Backlog + all intelligence maps                    */
/* ------------------------------------------------------------------ */
function getData(auth) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + SHEET_NAME + '" not found. Rename your ledger tab or edit SHEET_NAME in Code.gs.');

  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = sh.getLastRow();

  const out = {
    transactions: [],
    backlog: [],
    checks: [],           // Check Registry: pending PDCs due within 30 days (incl. overdue)
    checkBatches: [],     // unique Batch Codes among those pending checks
    categories: [], beneficiaries: [], recipients: [], destinations: [], accounts: [],
    beneficiaryMap: {},   // beneficiary → most recent {destination, account, recipient}
    combos: [],           // Smart Memory quads {b, r, d, a}, most recent first
    categoryPrefixMap: {},// Expense Category → most common trace prefix
    currentBalance: 0,
    balanceAsOf: null,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    threshold: Number(PropertiesService.getUserProperties().getProperty('CASH_THRESHOLD')) || 0
  };
  out.userDisplayName = displayName;   // personalizes the frontend greeting

  const catSet = {}, benSet = {}, recSet = {}, destSet = {}, acctSet = {};
  const comboSeen = {};
  const prefCounts = {};
  const traceSet = {};   // normalized Ledger traces (col E) for check-duplicate detection

  if (lastRow >= FIRST_DATA_ROW) {
    const values = sh.getRange(FIRST_DATA_ROW, 1, lastRow - HEADER_ROWS, NUM_COLS).getValues();

    values.forEach(function (r, i) {
      const rawDate = r[COL.DATE - 1];
      if (!(rawDate instanceof Date)) return;
      const withdrawal = toNum(r[COL.WITHDRAWAL - 1]);
      const deposit = toNum(r[COL.DEPOSIT - 1]);
      if (!withdrawal && !deposit) return;

      const balRaw = r[COL.BALANCE - 1];
      const settled = (balRaw !== '' && balRaw !== null && !isNaN(Number(balRaw)));
      const tx = r[COL.TX - 1] === true;

      /* 3-state lifecycle:
         planned  (TX=false, OK=false) — not yet in the bank app
         inputted (TX=true,  OK=false) — in the bank, awaiting approval
         settled  (TX=true,  OK=true, numeric P) — cleared, in the chain */
      const state = settled ? 'settled' : (tx ? 'inputted' : 'planned');

      const txn = {
        row: FIRST_DATA_ROW + i,
        uid: String(r[COL.UID - 1] || ''),
        date: Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd'),
        beneficiary: String(r[COL.ACTUAL_PAYEE - 1] || ''),
        recipient: String(r[COL.PAYEE - 1] || ''),
        destination: String(r[COL.DEST - 1] || ''),
        account: String(r[COL.ACCT - 1] || ''),
        category: String(r[COL.EXPENSE - 1] || ''),
        details: String(r[COL.DETAILS - 1] || ''),
        type: String(r[COL.TYPE - 1] || ''),
        trace: String(r[COL.TRACE - 1] || ''),
        withdrawal: withdrawal, deposit: deposit,
        settled: settled, state: state,
        balance: settled ? Number(balRaw) : null
      };
      out.transactions.push(txn);

      if (settled) { out.currentBalance = Number(balRaw); out.balanceAsOf = txn.date; }
      if (txn.category) catSet[txn.category] = true;
      if (txn.beneficiary) benSet[txn.beneficiary] = true;
      if (txn.recipient) recSet[txn.recipient] = true;
      if (txn.destination) destSet[txn.destination] = true;
      if (txn.account) acctSet[txn.account] = true;
      if (txn.trace) traceSet[txn.trace.trim().toLowerCase()] = true;

      // Beneficiary memory — most recent wins
      if (txn.beneficiary && (txn.destination || txn.account || txn.recipient)) {
        out.beneficiaryMap[txn.beneficiary] = {
          destination: txn.destination, account: txn.account, recipient: txn.recipient
        };
      }

      // Smart Memory quads — dedupe on lowercased quad; most recent first.
      // Beneficiary is included so typing in ANY of the four fields can
      // filter the other three.
      if (txn.beneficiary || txn.recipient || txn.destination || txn.account) {
        const key = [txn.beneficiary, txn.recipient, txn.destination, txn.account]
          .map(function (x) { return String(x).trim().toLowerCase(); }).join('|');
        if (comboSeen[key] !== undefined) out.combos.splice(comboSeen[key], 1);
        out.combos.unshift({ b: txn.beneficiary.trim(), r: txn.recipient.trim(), d: txn.destination.trim(), a: txn.account.trim() });
        out.combos.forEach(function (c, idx) {
          comboSeen[[c.b, c.r, c.d, c.a].map(function (x) { return x.toLowerCase(); }).join('|')] = idx;
        });
      }

      // Trace intelligence per category ('20260630_LogisticsTraviz01' → 'LogisticsTraviz')
      if (txn.category && txn.trace && txn.trace.indexOf('PLANNED_') !== 0) {
        const stripped = txn.trace.replace(/^\d{8}_/, '');
        const m = stripped.match(/^[A-Za-z]+/);
        if (m) {
          const c = prefCounts[txn.category] || (prefCounts[txn.category] = {});
          c[m[0]] = (c[m[0]] || 0) + 1;
        }
      }
    });
  }

  out.categories = Object.keys(catSet).sort();
  out.beneficiaries = Object.keys(benSet).sort();
  out.recipients = Object.keys(recSet).sort();
  out.destinations = Object.keys(destSet).sort();
  out.accounts = Object.keys(acctSet).sort();

  Object.keys(prefCounts).forEach(function (cat) {
    let best = '', n = 0;
    Object.keys(prefCounts[cat]).forEach(function (p) {
      if (prefCounts[cat][p] > n) { best = p; n = prefCounts[cat][p]; }
    });
    if (best) out.categoryPrefixMap[cat] = best;
  });

  readBacklog_(ss, tz, out);
  readChecks_(ss, tz, out, traceSet);
  return out;
}

/* ---- Backlog reader (READ-ONLY) ----
   Manually typed Backlog rows deliberately get NO UID here. A Shadow ID
   is minted only when an item is created via the app's 'Add Backlog
   Entry' form or when it hits the calendar (moves to the Ledger).
   UID-less rows are identified by their sheet row plus a
   beneficiary+amount fingerprint. */
function readBacklog_(ss, tz, out) {
  const sh = ss.getSheetByName(BACKLOG_SHEET);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return;

  const values = sh.getRange(FIRST_DATA_ROW, 1, lastRow - HEADER_ROWS, BL_COLS).getValues();

  values.forEach(function (r, i) {
    const withdrawal = toNum(r[COL.WITHDRAWAL - 1]);
    const deposit = toNum(r[COL.DEPOSIT - 1]);
    const hasContent = withdrawal || deposit || String(r[COL.ACTUAL_PAYEE - 1] || '').trim();
    if (!hasContent) return;

    const rawDate = r[COL.DATE - 1];
    out.backlog.push({
      row: FIRST_DATA_ROW + i,
      uid: String(r[COL.UID - 1] || ''),   // '' for hand-typed rows — by design
      date: (rawDate instanceof Date) ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd') : '',
      beneficiary: String(r[COL.ACTUAL_PAYEE - 1] || ''),
      recipient: String(r[COL.PAYEE - 1] || ''),
      destination: String(r[COL.DEST - 1] || ''),
      account: String(r[COL.ACCT - 1] || ''),
      category: String(r[COL.EXPENSE - 1] || ''),
      details: String(r[COL.DETAILS - 1] || ''),
      type: String(r[COL.TYPE - 1] || ''),
      trace: String(r[COL.TRACE - 1] || ''),
      withdrawal: withdrawal, deposit: deposit,
      priority: String(r[COL.PRIORITY - 1] || 'Med')
    });
  });
}

/* ---- Check Registry reader (READ-ONLY) ----
   Streams pending PDCs due within CHK_WINDOW_DAYS (overdue included so
   the drawer can flag them red). A check whose Check Number matches any
   Ledger trace (col E) is flagged isDuplicate — it's already recorded. */
function readChecks_(ss, tz, out, traceSet) {
  const sh = ss.getSheetByName(CHECKS_SHEET);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < CHK_FIRST_DATA_ROW) return;

  const todayStr = out.today;
  const horizon = new Date(todayStr + 'T00:00:00');
  horizon.setDate(horizon.getDate() + CHK_WINDOW_DAYS);
  const horizonStr = Utilities.formatDate(horizon, tz, 'yyyy-MM-dd');

  const values = sh.getRange(CHK_FIRST_DATA_ROW, 1, lastRow - CHK_HEADER_ROWS, CHK_NUM_COLS).getValues();
  const batchSet = {};

  const iso = function (v) {
    return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();
  };

  values.forEach(function (r, i) {
    const status = String(r[CHK_COL.STATUS - 1] || '').trim();
    // OLD:  if (status.toLowerCase() !== 'pending') return;
    const statusLc = status.toLowerCase();
    if (statusLc !== 'pending' && statusLc !== 'bounced') return;   // Bounced stays visible for management

    const targetDate = iso(r[CHK_COL.TARGET_DATE - 1]);
    // Window: due within the next 30 days, PLUS anything already overdue
    if (targetDate && targetDate > horizonStr) return;

    const checkNum = String(r[CHK_COL.CHECK_NUM - 1] || '').trim();
    const batch = String(r[CHK_COL.BATCH - 1] || '').trim();
    if (batch) batchSet[batch] = true;

    out.checks.push({
      row: CHK_FIRST_DATA_ROW + i,
      dateIssued: iso(r[CHK_COL.DATE_ISSUED - 1]),
      bank: String(r[CHK_COL.BANK - 1] || ''),
      acctNum: String(r[CHK_COL.ACCT_NUM - 1] || ''),
      checkNum: checkNum,
      batch: batch,
      payee: String(r[CHK_COL.PAYEE - 1] || ''),
      category: String(r[CHK_COL.CATEGORY - 1] || ''),
      memo: String(r[CHK_COL.MEMO - 1] || ''),
      amount: toNum(r[CHK_COL.AMOUNT - 1]),
      status: status,
      daysDue: String(r[CHK_COL.DAYS_DUE - 1] || ''),
      targetDate: targetDate,
      dateCleared: iso(r[CHK_COL.DATE_CLEARED - 1]),
      overdue: !!(targetDate && targetDate < todayStr),
      isDuplicate: !!(checkNum && traceSet[checkNum.toLowerCase()])
    });
  });

  out.checkBatches = Object.keys(batchSet).sort();
}

/* Locate a check by Check Number, verified against an amount+payee
   fingerprint; falls back to number-only when it's unique. */
function findCheckRow_(sh, checkNum, amount, payee) {
  const lastRow = sh.getLastRow();
  if (lastRow < CHK_FIRST_DATA_ROW) return null;
  const values = sh.getRange(CHK_FIRST_DATA_ROW, 1, lastRow - CHK_HEADER_ROWS, CHK_NUM_COLS).getValues();
  const targetNum = String(checkNum || '').trim().toLowerCase();
  if (!targetNum) return null;
  const wantAmt = Number(amount) || 0;
  const wantPayee = String(payee || '').trim().toLowerCase();
  let numOnlyMatch = null, numOnlyCount = 0;
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (String(r[CHK_COL.CHECK_NUM - 1] || '').trim().toLowerCase() !== targetNum) continue;
    numOnlyCount++;
    if (numOnlyMatch === null) numOnlyMatch = CHK_FIRST_DATA_ROW + i;
    const rowAmt = toNum(r[CHK_COL.AMOUNT - 1]);
    const rowPayee = String(r[CHK_COL.PAYEE - 1] || '').trim().toLowerCase();
    if ((!wantAmt || Math.abs(rowAmt - wantAmt) < 0.005) &&
        (!wantPayee || rowPayee === wantPayee)) {
      return CHK_FIRST_DATA_ROW + i;
    }
  }
  return numOnlyCount === 1 ? numOnlyMatch : null;
}

/* Defer Target Deposit Date (col L) and/or set Status (col J).
   Deferring a Bounced check auto-resets it to Pending. */
function updateCheckStatus(auth, checkNum, amount, payee, newDate, newStatus) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!checkNum) throw new Error('Missing check number.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const csh = ss.getSheetByName(CHECKS_SHEET);
    if (!csh) throw new Error('Sheet "' + CHECKS_SHEET + '" not found.');
    const row = findCheckRow_(csh, checkNum, amount, payee);
    if (!row) throw new Error('Could not find check #' + checkNum + ' — refresh and try again.');
    const curStatus = String(csh.getRange(row, CHK_COL.STATUS).getValue() || '').trim().toLowerCase();
    if (newDate) {
      const d = new Date(newDate + 'T00:00:00');
      if (isNaN(d.getTime())) throw new Error('Invalid date.');
      csh.getRange(row, CHK_COL.TARGET_DATE).setValue(d);
      if (curStatus === 'bounced' && !newStatus) newStatus = 'Pending';
    }
    if (newStatus) csh.getRange(row, CHK_COL.STATUS).setValue(newStatus);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* Record Bounce: statement mirroring via double-entry.
   Row 1 settled withdrawal ('[Payee] Check Bounced (#N)') +
   Row 2 settled deposit ('Bounced Check Credited Back'), inserted
   chronologically into the settled block with full P-formula repair,
   then the check's Status (col J) flips to 'Bounced'. */
function handleCheckBounce(auth, chk) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!chk || !chk.checkNum) throw new Error('Missing check data.');
  const amount = Number(chk.amount);
  if (!(amount > 0)) throw new Error('Check amount must be a positive number.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    const csh = ss.getSheetByName(CHECKS_SHEET);
    if (!csh) throw new Error('Sheet "' + CHECKS_SHEET + '" not found.');
    const tz = ss.getSpreadsheetTimeZone();
    const crow = findCheckRow_(csh, chk.checkNum, amount, chk.payee);
    if (!crow) throw new Error('Could not find check #' + chk.checkNum + ' in the Checks sheet.');
    const curStatus = String(csh.getRange(crow, CHK_COL.STATUS).getValue() || '').trim().toLowerCase();
    if (curStatus === 'bounced') throw new Error('Check #' + chk.checkNum + ' is already recorded as bounced.');

    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const today = new Date(todayStr + 'T00:00:00');
    const lastRow = sh.getLastRow();

    let lastSettled = HEADER_ROWS;
    if (lastRow >= FIRST_DATA_ROW) {
      const balances = sh.getRange(FIRST_DATA_ROW, COL.BALANCE, lastRow - HEADER_ROWS, 1).getValues();
      for (let i = balances.length - 1; i >= 0; i--) {
        const v = balances[i][0];
        if (v !== '' && v !== null && !isNaN(Number(v))) { lastSettled = FIRST_DATA_ROW + i; break; }
      }
    }
    let insertAfter = HEADER_ROWS;
    if (lastSettled >= FIRST_DATA_ROW) {
      const dates = sh.getRange(FIRST_DATA_ROW, COL.DATE, lastSettled - HEADER_ROWS, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        const dv = dates[i][0];
        if (dv instanceof Date && dv.getTime() <= today.getTime()) insertAfter = FIRST_DATA_ROW + i;
      }
    }

    sh.insertRowsAfter(Math.max(insertAfter, HEADER_ROWS), 2);
    const r1 = insertAfter + 1;

    const dayStr = Utilities.formatDate(today, tz, 'EEE');
    const payee = String(chk.payee || '').trim();
    const category = String(chk.category || '').trim();

    function mkRow(traceV, detailsV, typeV, isDeposit) {
      const row = new Array(NUM_COLS).fill('');
      row[COL.TX - 1] = true;
      row[COL.OK - 1] = true;
      row[COL.DAY - 1] = dayStr;
      row[COL.DATE - 1] = today;
      row[COL.TRACE - 1] = traceV;
      row[COL.ACTUAL_PAYEE - 1] = payee;
      row[COL.DEST - 1] = 'Check';
      row[COL.EXPENSE - 1] = category;
      row[COL.DETAILS - 1] = detailsV;
      row[COL.TYPE - 1] = typeV;
      row[COL.AMOUNT - 1] = amount;
      if (isDeposit) row[COL.DEPOSIT - 1] = amount;
      else row[COL.WITHDRAWAL - 1] = amount;
      row[COL.UID - 1] = Utilities.getUuid();
      return row;
    }

    const row1 = mkRow(String(chk.checkNum), payee + ' Check Bounced (#' + chk.checkNum + ')', 'Direct', false);
    const row2 = mkRow(String(chk.checkNum) + '_RETURN', 'Bounced Check Credited Back', 'Deposit', true);
    sh.getRange(r1, 1, 2, NUM_COLS).setValues([row1, row2]);

    const blockEnd = Math.max(lastSettled + 2, r1 + 1);
    sh.getRange(FIRST_DATA_ROW, COL.BALANCE, blockEnd - HEADER_ROWS, 1).setFormulaR1C1(BALANCE_FORMULA_R1C1);
    SpreadsheetApp.flush();

    const v1 = sh.getRange(r1, COL.BALANCE).getValue();
    const v2 = sh.getRange(r1 + 1, COL.BALANCE).getValue();
    const vEnd = sh.getRange(blockEnd, COL.BALANCE).getValue();
    if ([v1, v2, vEnd].some(function (v) { return v === '' || v === null || isNaN(Number(v)); })) {
      throw new Error('Bounce recorded, but the balance chain did not fully resolve — check rows ' + r1 + '–' + blockEnd + '.');
    }

    csh.getRange(crow, CHK_COL.STATUS).setValue('Bounced');
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* Locate a Backlog row: by UID when present, otherwise by row number
   verified against a beneficiary+amount fingerprint so a concurrently
   edited sheet never touches the wrong row. Returns row or null. */
function findBacklogRow_(bsh, ref) {
  if (ref && ref.uid) return findRowByUid(bsh, ref.uid);
  if (!ref || !ref.row) return null;
  const lastRow = bsh.getLastRow();
  if (ref.row < FIRST_DATA_ROW || ref.row > lastRow) return null;
  const vals = bsh.getRange(ref.row, 1, 1, BL_COLS).getValues()[0];
  const ben = String(vals[COL.ACTUAL_PAYEE - 1] || '').trim();
  const amt = toNum(vals[COL.WITHDRAWAL - 1]) || toNum(vals[COL.DEPOSIT - 1]);
  if (ben === String(ref.beneficiary || '').trim() &&
      Math.abs(amt - Number(ref.amount || 0)) < 0.005) return ref.row;
  return null;
}

/* ------------------------------------------------------------------ */
/*  ADD planned (single) — optionally consumes a Backlog item.         */
/*  This is where a Backlog item earns its Shadow ID: the fresh UID is */
/*  written to the new Ledger row, and the Backlog source is deleted.  */
/* ------------------------------------------------------------------ */
function addTransaction(auth, t) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!t || !t.date) throw new Error('Date is required.');
  if (!String(t.beneficiary || '').trim() && !String(t.recipient || '').trim()) {
    throw new Error('Enter at least a Beneficiary or a Recipient.');
  }
  const amount = Number(t.amount);
  if (!(amount > 0)) throw new Error('Amount must be a positive number.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    const tz = ss.getSpreadsheetTimeZone();

    const d = new Date(t.date + 'T00:00:00');
    if (isNaN(d.getTime())) throw new Error('Invalid date.');

    const row = new Array(NUM_COLS).fill('');
    row[COL.TX - 1] = false;   // State 1: Planned
    row[COL.OK - 1] = false;
    row[COL.DAY - 1] = Utilities.formatDate(d, tz, 'EEE');
    row[COL.DATE - 1] = d;
    row[COL.TRACE - 1] = String(t.trace || '').trim() ||
      ('PLANNED_' + Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss'));
    row[COL.ACTUAL_PAYEE - 1] = String(t.beneficiary || '').trim();
    row[COL.PAYEE - 1] = String(t.recipient || '').trim();
    row[COL.DEST - 1] = String(t.destination || '').trim();
    row[COL.ACCT - 1] = String(t.account || '').trim();
    row[COL.EXPENSE - 1] = String(t.category || '').trim();
    row[COL.DETAILS - 1] = String(t.details || '').trim();
    row[COL.TYPE - 1] = String(t.type || '').trim();
    row[COL.AMOUNT - 1] = amount;
    if (t.direction === 'in') row[COL.DEPOSIT - 1] = amount;
    else row[COL.WITHDRAWAL - 1] = amount;
    row[COL.UID - 1] = Utilities.getUuid();   // Shadow ID minted on Ledger entry

    const newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, COL.ACCT).setNumberFormat('@');   // leading zeros survive
    sh.getRange(newRow, 1, 1, NUM_COLS).setValues([row]);

    // Bidirectional sync: item came from the Backlog → delete it there
    if (t.backlogUid || t.backlogRow) {
      const bsh = ss.getSheetByName(BACKLOG_SHEET);
      if (bsh) {
        const br = findBacklogRow_(bsh, {
          uid: t.backlogUid, row: Number(t.backlogRow) || 0,
          beneficiary: t.backlogBen, amount: t.backlogAmt
        });
        if (br) bsh.deleteRow(br);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  BULK ENTRY — batch of planned rows (traces are WYSIWYG)            */
/* ------------------------------------------------------------------ */
function addBulkTransactions(auth, batch) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!batch || !batch.date) throw new Error('Transaction Date is required.');
  if (!String(batch.tag || '').trim()) throw new Error('Group Tag is required — it builds the trace numbers.');

  const rows = (batch.rows || []).filter(function (r) {
    return r && (String(r.beneficiary || '').trim() || Number(r.amount) || String(r.details || '').trim());
  });
  if (!rows.length) throw new Error('Add at least one entry row.');
  rows.forEach(function (r, i) {
    if (!String(r.beneficiary || '').trim()) throw new Error('Row ' + (i + 1) + ': Beneficiary is required.');
    if (!(Number(r.amount) > 0)) throw new Error('Row ' + (i + 1) + ': Amount must be a positive number.');
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    const tz = ss.getSpreadsheetTimeZone();

    const d = new Date(batch.date + 'T00:00:00');
    if (isNaN(d.getTime())) throw new Error('Invalid Transaction Date.');
    const refD = batch.refDate ? new Date(batch.refDate + 'T00:00:00') : d;
    if (isNaN(refD.getTime())) throw new Error('Invalid Reference Date.');

    const refStr = Utilities.formatDate(refD, tz, 'yyyyMMdd');
    const tagClean = String(batch.tag).replace(/\s+/g, '');
    const rangeStr = buildRangeString_(batch.rangeStart, batch.rangeEnd);
    const dayStr = Utilities.formatDate(d, tz, 'EEE');
    const defaultRecipient = String(batch.recipient || '').trim();
    const defaultDestination = String(batch.destination || '').trim();
    const defaultAccount = String(batch.account || '').trim();
    const category = String(batch.category || '').trim();
    const type = String(batch.type || '').trim();
    const isIn = batch.direction === 'in';

    const matrix = rows.map(function (r, i) {
      const seq = padSeq_(i + 1);
      const trace = String(r.trace || '').trim() || (rangeStr
        ? refStr + '_' + tagClean + '_' + rangeStr + '_' + seq
        : refStr + '_' + tagClean + seq);
      const amount = Number(r.amount);
      const recipient = String(r.recipient || '').trim() || defaultRecipient;
      const destination = String(r.destination || '').trim() || defaultDestination;
      const account = String(r.account || '').trim() || defaultAccount;

      const row = new Array(NUM_COLS).fill('');
      row[COL.TX - 1] = false;   // State 1: Planned
      row[COL.OK - 1] = false;
      row[COL.DAY - 1] = dayStr;
      row[COL.DATE - 1] = d;
      row[COL.TRACE - 1] = trace;
      row[COL.ACTUAL_PAYEE - 1] = String(r.beneficiary).trim();
      row[COL.PAYEE - 1] = recipient;
      row[COL.DEST - 1] = destination;
      row[COL.ACCT - 1] = account;
      row[COL.EXPENSE - 1] = category;
      row[COL.DETAILS - 1] = String(r.details || '').trim();
      row[COL.TYPE - 1] = type;
      row[COL.AMOUNT - 1] = amount;
      if (isIn) row[COL.DEPOSIT - 1] = amount;
      else row[COL.WITHDRAWAL - 1] = amount;
      row[COL.UID - 1] = Utilities.getUuid();
      return row;
    });

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, COL.ACCT, matrix.length, 1).setNumberFormat('@');
    sh.getRange(startRow, 1, matrix.length, NUM_COLS).setValues(matrix);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

function buildRangeString_(startISO, endISO) {
  if (!startISO || !endISO) return '';
  const s = new Date(startISO + 'T00:00:00'), e = new Date(endISO + 'T00:00:00');
  if (isNaN(s.getTime()) || isNaN(e.getTime())) return '';
  const mm = function (dt) { return ('0' + (dt.getMonth() + 1)).slice(-2); };
  if (s.getMonth() === e.getMonth() && s.getFullYear() === e.getFullYear()) {
    return mm(s) + '_' + s.getDate() + '-' + e.getDate();
  }
  return mm(s) + '_' + s.getDate() + '-' + mm(e) + '_' + e.getDate();
}

function padSeq_(n) { return ('0' + n).slice(-2); }

/* ------------------------------------------------------------------ */
/*  MARK INPUTTED — State 1 → State 2 (TX=true, OK stays false)        */
/* ------------------------------------------------------------------ */
function markInputted(auth, uid) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!uid) throw new Error('Missing transaction ID.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const row = findRowByUid(sh, uid);
    if (!row) throw new Error('Could not find that transaction — refresh the dashboard.');
    const balRaw = sh.getRange(row, COL.BALANCE).getValue();
    if (balRaw !== '' && balRaw !== null && !isNaN(Number(balRaw))) {
      throw new Error('That row is already settled.');
    }
    sh.getRange(row, COL.TX).setValue(true);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  SETTLE — Chronological Re-Stitching + global P formula repair      */
/* ------------------------------------------------------------------ */
function settleTransaction(auth, payload) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!payload || !payload.uid) throw new Error('Missing transaction ID.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    const tz = ss.getSpreadsheetTimeZone();
    const lastRow = sh.getLastRow();

    const srcRow = findRowByUid(sh, payload.uid);
    if (!srcRow) throw new Error('Could not find that transaction. It may have been settled or deleted already — refresh the dashboard.');

    const rowVals = sh.getRange(srcRow, 1, 1, NUM_COLS).getValues()[0];
    const balRaw = rowVals[COL.BALANCE - 1];
    if (balRaw !== '' && balRaw !== null && !isNaN(Number(balRaw))) {
      throw new Error('That row is already settled. Refresh the dashboard.');
    }

    const wasWithdrawal = toNum(rowVals[COL.WITHDRAWAL - 1]) > 0;
    const plannedAmount = wasWithdrawal
      ? toNum(rowVals[COL.WITHDRAWAL - 1]) : toNum(rowVals[COL.DEPOSIT - 1]);

    const finalAmount = (payload.amount !== undefined && payload.amount !== null && payload.amount !== '')
      ? Number(payload.amount) : plannedAmount;
    if (!(finalAmount > 0)) throw new Error('Final amount must be a positive number.');

    const settleDate = payload.date
      ? new Date(payload.date + 'T00:00:00')
      : new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + 'T00:00:00');
    if (isNaN(settleDate.getTime())) throw new Error('Invalid settle date.');

    // Type: explicit payload > type already on the row > direction default
    const existingType = String(rowVals[COL.TYPE - 1] || '').trim();
    const finalType = String(payload.type || '').trim() || existingType || (wasWithdrawal ? 'Direct' : 'Deposit');

    // Trace cleanup: real trace if given, otherwise clear only the PLANNED_ tag
    const oldTrace = String(rowVals[COL.TRACE - 1] || '');
    let finalTrace = String(payload.trace || '').trim();
    if (!finalTrace) finalTrace = oldTrace.indexOf('PLANNED_') === 0 ? '' : oldTrace;

    /* ---- UPDATE in place ---- */
    sh.getRange(srcRow, COL.TX).setValue(true);
    sh.getRange(srcRow, COL.OK).setValue(true);
    sh.getRange(srcRow, COL.DAY).setValue(Utilities.formatDate(settleDate, tz, 'EEE'));
    sh.getRange(srcRow, COL.DATE).setValue(settleDate);
    sh.getRange(srcRow, COL.TRACE).setValue(finalTrace);
    if (payload.beneficiary !== undefined) sh.getRange(srcRow, COL.ACTUAL_PAYEE).setValue(String(payload.beneficiary).trim());
    if (payload.recipient !== undefined)   sh.getRange(srcRow, COL.PAYEE).setValue(String(payload.recipient).trim());
    if (payload.category !== undefined)    sh.getRange(srcRow, COL.EXPENSE).setValue(String(payload.category).trim());
    if (payload.details !== undefined)     sh.getRange(srcRow, COL.DETAILS).setValue(String(payload.details).trim());
    sh.getRange(srcRow, COL.TYPE).setValue(finalType);
    sh.getRange(srcRow, COL.AMOUNT).setValue(finalAmount);
    sh.getRange(srcRow, wasWithdrawal ? COL.WITHDRAWAL : COL.DEPOSIT).setValue(finalAmount);

    /* ---- CHRONOLOGICAL RE-STITCHING ----
       Find the settled block (rows with numeric P, excluding this row),
       insert this row after the last settled row whose date ≤ the settle
       date, then re-apply the balance formula to the WHOLE block so a
       back-dated settlement re-flows every subsequent balance. */
    let lastSettled = HEADER_ROWS;
    if (lastRow >= FIRST_DATA_ROW) {
      const balances = sh.getRange(FIRST_DATA_ROW, COL.BALANCE, lastRow - HEADER_ROWS, 1).getValues();
      for (let i = balances.length - 1; i >= 0; i--) {
        const r = FIRST_DATA_ROW + i;
        const v = balances[i][0];
        if (r !== srcRow && v !== '' && v !== null && !isNaN(Number(v))) { lastSettled = r; break; }
      }
    }

    let insertAfter = HEADER_ROWS;
    if (lastSettled >= FIRST_DATA_ROW) {
      const dates = sh.getRange(FIRST_DATA_ROW, COL.DATE, lastSettled - HEADER_ROWS, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        const r = FIRST_DATA_ROW + i;
        if (r === srcRow) continue;
        const dv = dates[i][0];
        if (dv instanceof Date && dv.getTime() <= settleDate.getTime()) insertAfter = r;
      }
    }

    let destRow = srcRow;
    const target = insertAfter + 1;
    if (srcRow !== target) {
      sh.moveRows(sh.getRange(srcRow + ':' + srcRow), target);
      destRow = (srcRow > target) ? target : target - 1;
    }

    /* ---- GLOBAL FORMULA REPAIR across the settled block ---- */
    const blockEnd = Math.max(lastSettled + 1, destRow);
    sh.getRange(FIRST_DATA_ROW, COL.BALANCE, blockEnd - HEADER_ROWS, 1)
      .setFormulaR1C1(BALANCE_FORMULA_R1C1);
    SpreadsheetApp.flush();

    const movedVal = sh.getRange(destRow, COL.BALANCE).getValue();
    const chainVal = sh.getRange(blockEnd, COL.BALANCE).getValue();
    if (movedVal === '' || movedVal === null || isNaN(Number(movedVal)) ||
        chainVal === '' || chainVal === null || isNaN(Number(chainVal))) {
      throw new Error('Settled, but the balance chain did not fully resolve — check rows ' + destRow + '–' + blockEnd + ' in the sheet.');
    }
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  DELETE a pending row (by Shadow UID)                               */
/* ------------------------------------------------------------------ */
function deletePlanned(auth, uid) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!uid) throw new Error('Missing transaction ID.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const row = findRowByUid(sh, uid);
    if (!row) throw new Error('Could not find that transaction — refresh the dashboard.');
    const balRaw = sh.getRange(row, COL.BALANCE).getValue();
    if (balRaw !== '' && balRaw !== null && !isNaN(Number(balRaw))) {
      throw new Error('Refused: that row is already settled. Delete settled rows in the sheet itself.');
    }
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  BASKET: Ledger → Backlog (re-adds the row to the Backlog sheet)    */
/* ------------------------------------------------------------------ */
function moveToBacklog(auth, uid) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!uid) throw new Error('Missing transaction ID.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    const bsh = ss.getSheetByName(BACKLOG_SHEET);
    if (!bsh) throw new Error('Sheet "' + BACKLOG_SHEET + '" not found — create it (Ledger layout + col R Priority).');

    const row = findRowByUid(sh, uid);
    if (!row) throw new Error('Could not find that transaction — refresh the dashboard.');
    const vals = sh.getRange(row, 1, 1, NUM_COLS).getValues()[0];
    const balRaw = vals[COL.BALANCE - 1];
    if (balRaw !== '' && balRaw !== null && !isNaN(Number(balRaw))) {
      throw new Error('Only pending (unsettled) items can move to the Backlog.');
    }

    const blRow = vals.slice(0, NUM_COLS);
    blRow[COL.TX - 1] = false;       // reset lifecycle to pre-planned
    blRow[COL.OK - 1] = false;
    blRow[COL.BALANCE - 1] = '';
    blRow.push('Med');               // R — default priority

    const dest = bsh.getLastRow() + 1;
    bsh.getRange(dest, COL.ACCT).setNumberFormat('@');
    bsh.getRange(dest, 1, 1, BL_COLS).setValues([blRow]);
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  BACKLOG REORDER — rewrite rows in the order of the given ORIGINAL  */
/*  sheet row numbers (from the last getData read). Row-based identity */
/*  means drag-and-drop works even for hand-typed rows without a UID.  */
/* ------------------------------------------------------------------ */
function reorderBacklog(auth, rowOrder) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!rowOrder || !rowOrder.length) return getData(auth);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const bsh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKLOG_SHEET);
    if (!bsh) throw new Error('Sheet "' + BACKLOG_SHEET + '" not found.');
    const lastRow = bsh.getLastRow();
    if (lastRow < FIRST_DATA_ROW) return getData(auth);

    const n = lastRow - HEADER_ROWS;
    if (rowOrder.length !== n) {
      // Sheet changed since the drawer loaded (row added/removed) — refuse
      // to guess and just hand back fresh data.
      return getData(auth);
    }
    const rng = bsh.getRange(FIRST_DATA_ROW, 1, n, BL_COLS);
    const rows = rng.getValues();

    const byRow = {};
    rows.forEach(function (r, i) { byRow[FIRST_DATA_ROW + i] = r; });

    const ordered = [], taken = {};
    rowOrder.forEach(function (rn) {
      rn = Number(rn);
      if (byRow[rn] && !taken[rn]) { ordered.push(byRow[rn]); taken[rn] = true; }
    });
    rows.forEach(function (r, i) {   // safety: anything unreferenced keeps a slot at the end
      if (!taken[FIRST_DATA_ROW + i]) ordered.push(r);
    });

    bsh.getRange(FIRST_DATA_ROW, COL.ACCT, n, 1).setNumberFormat('@');
    rng.setValues(ordered);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  ADD BACKLOG ENTRY — created via the app, so it DOES get a UID      */
/* ------------------------------------------------------------------ */
function addBacklogEntry(auth, t) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!t) throw new Error('Missing entry.');
  if (!String(t.beneficiary || '').trim()) throw new Error('Beneficiary is required.');
  const amount = Number(t.amount);
  if (!(amount > 0)) throw new Error('Amount must be a positive number.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bsh = ss.getSheetByName(BACKLOG_SHEET);
    if (!bsh) throw new Error('Sheet "' + BACKLOG_SHEET + '" not found — create it (Ledger layout + col R Priority).');
    const tz = ss.getSpreadsheetTimeZone();

    let d = null;
    if (t.date) {
      d = new Date(t.date + 'T00:00:00');
      if (isNaN(d.getTime())) throw new Error('Invalid date.');
    }

    const row = new Array(BL_COLS).fill('');
    row[COL.TX - 1] = false;
    row[COL.OK - 1] = false;
    if (d) { row[COL.DAY - 1] = Utilities.formatDate(d, tz, 'EEE'); row[COL.DATE - 1] = d; }
    row[COL.ACTUAL_PAYEE - 1] = String(t.beneficiary || '').trim();
    row[COL.PAYEE - 1] = String(t.recipient || '').trim();
    row[COL.EXPENSE - 1] = String(t.category || '').trim();
    row[COL.DETAILS - 1] = String(t.details || '').trim();
    row[COL.TYPE - 1] = String(t.type || '').trim();
    row[COL.AMOUNT - 1] = amount;
    if (t.direction === 'in') row[COL.DEPOSIT - 1] = amount;
    else row[COL.WITHDRAWAL - 1] = amount;
    row[COL.UID - 1] = Utilities.getUuid();   // app-created entries get a Shadow ID
    row[COL.PRIORITY - 1] = ['Low', 'Med', 'High'].indexOf(String(t.priority)) !== -1 ? String(t.priority) : 'Med';

    const dest = bsh.getLastRow() + 1;
    bsh.getRange(dest, COL.ACCT).setNumberFormat('@');
    bsh.getRange(dest, 1, 1, BL_COLS).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  EDIT BACKLOG ENTRY — updates fields in place. Deliberately does    */
/*  NOT mint a UID: a manual row stays ID-less until it hits the       */
/*  calendar. Located by UID, or row + fingerprint for manual rows.    */
/* ------------------------------------------------------------------ */
function updateBacklogEntry(auth, payload) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  if (!payload) throw new Error('Missing entry.');
  if (!String(payload.beneficiary || '').trim()) throw new Error('Beneficiary is required.');
  const amount = Number(payload.amount);
  if (!(amount > 0)) throw new Error('Amount must be a positive number.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bsh = ss.getSheetByName(BACKLOG_SHEET);
    if (!bsh) throw new Error('Sheet "' + BACKLOG_SHEET + '" not found.');
    const tz = ss.getSpreadsheetTimeZone();

    const row = findBacklogRow_(bsh, {
      uid: payload.refUid, row: Number(payload.refRow) || 0,
      beneficiary: payload.refBen, amount: payload.refAmt
    });
    if (!row) throw new Error('Could not find that Backlog item — it may have changed. Refresh and try again.');

    let d = null;
    if (payload.date) {
      d = new Date(payload.date + 'T00:00:00');
      if (isNaN(d.getTime())) throw new Error('Invalid date.');
    }

    bsh.getRange(row, COL.DAY).setValue(d ? Utilities.formatDate(d, tz, 'EEE') : '');
    bsh.getRange(row, COL.DATE).setValue(d || '');
    bsh.getRange(row, COL.ACTUAL_PAYEE).setValue(String(payload.beneficiary).trim());
    bsh.getRange(row, COL.PAYEE).setValue(String(payload.recipient || '').trim());
    bsh.getRange(row, COL.EXPENSE).setValue(String(payload.category || '').trim());
    bsh.getRange(row, COL.DETAILS).setValue(String(payload.details || '').trim());
    bsh.getRange(row, COL.AMOUNT).setValue(amount);
    if (payload.direction === 'in') {
      bsh.getRange(row, COL.WITHDRAWAL).setValue('');
      bsh.getRange(row, COL.DEPOSIT).setValue(amount);
    } else {
      bsh.getRange(row, COL.WITHDRAWAL).setValue(amount);
      bsh.getRange(row, COL.DEPOSIT).setValue('');
    }
    const prio = ['Low', 'Med', 'High'].indexOf(String(payload.priority)) !== -1 ? String(payload.priority) : 'Med';
    bsh.getRange(row, COL.PRIORITY).setValue(prio);
  } finally {
    lock.releaseLock();
  }
  return getData(auth);
}

/* ------------------------------------------------------------------ */
/*  One-time repairs (run from the Apps Script editor)                 */
/* ------------------------------------------------------------------ */
function restitchLedger() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lastRow = sh.getLastRow();
    if (lastRow < FIRST_DATA_ROW) return;

    const balances = sh.getRange(FIRST_DATA_ROW, COL.BALANCE, lastRow - HEADER_ROWS, 1).getValues();
    let lastSettled = HEADER_ROWS;
    for (let i = balances.length - 1; i >= 0; i--) {
      const v = balances[i][0];
      if (v !== '' && v !== null && !isNaN(Number(v))) { lastSettled = FIRST_DATA_ROW + i; break; }
    }
    if (lastSettled < FIRST_DATA_ROW) { Logger.log('No settled rows found.'); return; }

    const n = lastSettled - HEADER_ROWS;
    const rng = sh.getRange(FIRST_DATA_ROW, 1, n, NUM_COLS);
    const rows = rng.getValues();

    const indexed = rows.map(function (r, i) { return { r: r, i: i }; });
    indexed.sort(function (a, b) {
      const da = a.r[COL.DATE - 1], db = b.r[COL.DATE - 1];
      const ta = (da instanceof Date) ? da.getTime() : Number.MAX_SAFE_INTEGER;
      const tb = (db instanceof Date) ? db.getTime() : Number.MAX_SAFE_INTEGER;
      return (ta - tb) || (a.i - b.i);   // stable date sort
    });

    const sorted = indexed.map(function (x) { return x.r; });
    sh.getRange(FIRST_DATA_ROW, 1, n, COL.BALANCE - 1)
      .setValues(sorted.map(function (r) { return r.slice(0, COL.BALANCE - 1); }));
    sh.getRange(FIRST_DATA_ROW, COL.UID, n, 1)
      .setValues(sorted.map(function (r) { return [r[COL.UID - 1]]; }));
    sh.getRange(FIRST_DATA_ROW, COL.BALANCE, n, 1).setFormulaR1C1(BALANCE_FORMULA_R1C1);

    SpreadsheetApp.flush();
    Logger.log('Re-stitched ' + n + ' rows. Final balance: ' + sh.getRange(lastSettled, COL.BALANCE).getValue());
  } finally {
    lock.releaseLock();
  }
}

function backfillUids() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return;
  const vals = sh.getRange(FIRST_DATA_ROW, 1, lastRow - HEADER_ROWS, NUM_COLS).getValues();
  let n = 0;
  vals.forEach(function (r, i) {
    const hasMove = toNum(r[COL.WITHDRAWAL - 1]) || toNum(r[COL.DEPOSIT - 1]);
    const pending = r[COL.BALANCE - 1] === '' || r[COL.BALANCE - 1] === null;
    if (hasMove && pending && !r[COL.UID - 1]) {
      sh.getRange(FIRST_DATA_ROW + i, COL.UID).setValue(Utilities.getUuid());
      n++;
    }
  });
  Logger.log('Assigned UIDs to ' + n + ' pending Ledger row(s).');
}

/* ------------------------------------------------------------------ */
function findRowByUid(sh, uid) {
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return null;
  const uids = sh.getRange(FIRST_DATA_ROW, COL.UID, lastRow - HEADER_ROWS, 1).getValues();
  for (let i = 0; i < uids.length; i++) {
    if (String(uids[i][0]) === String(uid)) return FIRST_DATA_ROW + i;
  }
  return null;
}

function saveThreshold(auth, value) {
  // ---- SERVER-SIDE GATEKEEPER (defined in Auth.gs) ----
  if (!auth || !auth.user || !auth.hash) throw new Error('Not signed in — please log in again.');
  const displayName = verifySession(auth.user, auth.hash);

  const v = Number(value) || 0;
  PropertiesService.getUserProperties().setProperty('CASH_THRESHOLD', String(v));
  return getData(auth);
}

function toNum(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
