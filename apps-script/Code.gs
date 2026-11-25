/**
 * PDD Handover Letter – shared reference counter.
 *
 * Lives in a Google Sheet you own (Extensions > Apps Script).
 * Sheet "Counters": one row per bank (key, next number). Edit "next" to change where a bank's count starts.
 * Sheet "Register": every reference number ever issued.
 * A lock makes sure two people can never receive the same number.
 */

const COUNTERS = 'Counters';
const REGISTER = 'Register';

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'peek') return json_({ ok: true, next: peek_(key_(p.key)) });
    if (p.action === 'issue') return json_(Object.assign({ ok: true }, issue_(p)));
    if (p.action === 'recent') return json_({ ok: true, rows: recent_() });
    throw new Error('Unknown action');
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function key_(k) {
  k = String(k || '');
  if (!/^[a-z0-9-]{1,120}$/.test(k)) throw new Error('Invalid bank key');
  return k;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').slice(0, max || 200);
}

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function counters_() { return sheet_(COUNTERS, ['Bank key', 'Next number']); }
function register_() {
  return sheet_(REGISTER, ['Issued at', 'Ref No', 'Bank', 'Serial', 'User No', 'Prepared by', 'Borrower', 'Loan No', 'Letter date']);
}

function findRow_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}

function peek_(key) {
  const sh = counters_();
  const row = findRow_(sh, key);
  return row ? (Number(sh.getRange(row, 2).getValue()) || 1) : 1;
}

function issue_(p) {
  const key = key_(p.key);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = counters_();
    let row = findRow_(sh, key);
    let n;
    if (row) {
      n = Number(sh.getRange(row, 2).getValue()) || 1;
      sh.getRange(row, 2).setValue(n + 1);
    } else {
      n = 1;
      sh.appendRow([key, 2]);
    }
    const parts = [clean_(p.ini, 20)];
    if (p.user) parts.push(clean_(p.user, 10));
    parts.push(String(n), clean_(p.mon, 3), clean_(p.yr, 4));
    const ref = parts.join('/');
    register_().appendRow([new Date(), ref, clean_(p.bank), n, clean_(p.user, 10), clean_(p.by), clean_(p.borrower), clean_(p.loan, 60), clean_(p.date, 12)]);
    SpreadsheetApp.flush();
    return { serial: n, ref: ref };
  } finally {
    lock.releaseLock();
  }
}

function recent_() {
  const sh = register_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const count = Math.min(20, last - 1);
  const vals = sh.getRange(last - count + 1, 1, count, 9).getValues();
  return vals.reverse().map(r => ({
    ref: r[1], bank: r[2], userNo: String(r[4]), by: r[5], borrower: r[6], loan: r[7], date: r[8]
  }));
}
