/**
 * PDD Handover Letter – shared reference counter (version 2).
 *
 * Works whether this script was opened from a Google Sheet (Extensions > Apps Script)
 * or created on its own at script.google.com. If there is no sheet, it creates one
 * called "PDD Register" in your Drive the first time setup() runs.
 *
 * Tabs:
 *   Counters – one row per bank: key and next number. Edit "Next number" to change where a bank starts.
 *   Register – every reference number ever issued.
 * A lock makes sure two people can never receive the same number.
 *
 * ONE-TIME SETUP
 *   1. Paste this whole file over the old code and Save.
 *   2. Choose the function "setup" in the toolbar and click Run. Allow the permissions.
 *      The log shows the sheet's address.
 *   3. Deploy > Manage deployments > Edit (pencil) > Version: New version >
 *      Execute as: Me, Who has access: Anyone > Deploy.
 */

const COUNTERS = 'Counters';
const REGISTER = 'Register';

function setup() {
  const ss = ss_();
  counters_();
  register_();
  Logger.log('Counter is using this sheet: ' + ss.getUrl());
  return ss.getUrl();
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.action === 'ping') out = { ok: true, sheet: ss_().getName() };
    else if (p.action === 'peek') out = { ok: true, next: peek_(key_(p.key)) };
    else if (p.action === 'issue') out = Object.assign({ ok: true }, issue_(p));
    else if (p.action === 'recent') out = { ok: true, rows: recent_() };
    else throw new Error('Unknown action');
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return reply_(out, p.callback);
}

function doPost(e) {
  return doGet(e);
}

// JSONP reply when the page asks for one (works from any website without CORS problems).
function reply_(obj, callback) {
  const body = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]{0,60}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* deleted or no access: make a new one */ }
  }
  ss = SpreadsheetApp.create('PDD Register');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
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
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    // Reuse an empty default "Sheet1" for the first tab instead of leaving it lying around.
    const blank = ss.getSheets().filter(s => s.getLastRow() === 0 && /^Sheet\d*$/.test(s.getName()))[0];
    sh = blank ? blank.setName(name) : ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
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
  const keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === key) return i + 2;
  return 0;
}

function peek_(key) {
  const sh = counters_();
  const row = findRow_(sh, key);
  return row ? (parseInt(sh.getRange(row, 2).getValue(), 10) || 1) : 1;
}

function issue_(p) {
  const key = key_(p.key);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Counter is busy, please try again');
  try {
    const sh = counters_();
    const row = findRow_(sh, key);
    let n;
    if (row) {
      n = parseInt(sh.getRange(row, 2).getValue(), 10) || 1;
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
  const vals = sh.getRange(last - count + 1, 1, count, 9).getDisplayValues();
  return vals.reverse().map(r => ({
    ref: r[1], bank: r[2], userNo: r[4], by: r[5], borrower: r[6], loan: r[7], date: r[8]
  }));
}
