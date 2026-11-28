/**
 * PDD Handover Letter – shared reference counter (version 4).
 *
 * Works whether this script was opened from a Google Sheet (Extensions > Apps Script)
 * or created on its own at script.google.com. If there is no sheet, it creates one
 * called "PDD Register" in your Drive the first time setup() runs.
 *
 * Tabs:
 *   Counters – one row per bank: key and next number. Edit "Next number" to change where a bank starts.
 *   Register – every reference number ever issued, with the full letter (so it can be opened again).
 *   Users    – every person's user number. No two people get the same number: if the number
 *              from the first 3 letters is taken, the first two letters stay and the third
 *              is swapped for the next letter of the name until a free number is found.
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
const USERS = 'Users';

function setup() {
  const ss = ss_();
  counters_();
  register_();
  users_();
  Logger.log('Counter is using this sheet: ' + ss.getUrl());
  return ss.getUrl();
}

// Run this once from the toolbar to start every bank's count again at 1.
// (Old letters stay in the Register tab and can still be found.)
function resetCounters() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = counters_();
    const last = sh.getLastRow();
    if (last >= 2) sh.deleteRows(2, last - 1);
    SpreadsheetApp.flush();
    Logger.log('All counters reset: every bank starts again at 1.');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.action === 'ping') out = { ok: true, sheet: ss_().getName() };
    else if (p.action === 'peek') out = { ok: true, next: peek_(key_(p.key)) };
    else if (p.action === 'issue') out = Object.assign({ ok: true }, issue_(p));
    else if (p.action === 'recent') out = { ok: true, rows: recent_() };
    else if (p.action === 'user') out = Object.assign({ ok: true }, user_(p.name));
    else if (p.action === 'save') out = Object.assign({ ok: true }, save_(p.ref, p.data));
    else if (p.action === 'find') out = { ok: true, rows: find_(p.q) };
    else throw new Error('Unknown action');
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return reply_(out, p.callback);
}

// The website posts the full letter here (too long for a web address).
function doPost(e) {
  let p = (e && e.parameter) || {};
  try {
    const body = e && e.postData && e.postData.contents;
    if (body) p = Object.assign({}, p, JSON.parse(body));
  } catch (err) { /* not JSON: use the parameters */ }
  return doGet({ parameter: p });
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
  const sh = sheet_(REGISTER, ['Issued at', 'Ref No', 'Bank', 'Serial', 'User No', 'Prepared by', 'Borrower', 'Loan No', 'Letter date', 'Letter data']);
  if (!sh.getRange(1, 10).getValue()) sh.getRange(1, 10).setValue('Letter data').setFontWeight('bold');
  return sh;
}

// Store the full letter next to its reference number.
function save_(ref, data) {
  ref = clean_(ref, 60);
  if (!ref) throw new Error('Missing reference number');
  const json = typeof data === 'string' ? data : JSON.stringify(data || {});
  if (json.length > 45000) throw new Error('Letter too large to store');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Counter is busy, please try again');
  try {
    const sh = register_();
    const last = sh.getLastRow();
    if (last >= 2) {
      const refs = sh.getRange(2, 2, last - 1, 1).getDisplayValues();
      for (let i = refs.length - 1; i >= 0; i--) {
        if (refs[i][0] === ref) { sh.getRange(i + 2, 10).setValue(json); return { saved: ref }; }
      }
    }
    throw new Error('Reference number not found in the register');
  } finally {
    lock.releaseLock();
  }
}

// Search old letters by reference number, borrower, loan no. or bank (newest first).
function find_(q) {
  q = String(q || '').toLowerCase().trim();
  if (q.length < 2) throw new Error('Type at least 2 characters');
  const sh = register_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 10).getDisplayValues();
  const out = [];
  for (let i = vals.length - 1; i >= 0 && out.length < 25; i--) {
    const r = vals[i];
    const hay = [r[1], r[2], r[5], r[6], r[7]].join(' ').toLowerCase();
    if (hay.indexOf(q) === -1) continue;
    let data = null;
    try { data = r[9] ? JSON.parse(r[9]) : null; } catch (e) { data = null; }
    out.push({ ref: r[1], bank: r[2], by: r[5], borrower: r[6], loan: r[7], date: r[8], data: data });
  }
  return out;
}

function users_() { return sheet_(USERS, ['Name key', 'Name', 'User No', 'First seen']); }

// Letters A=1 … Z=26 of the first two letters plus the letter at position "third", plus any numbers in the name.
function userCandidate_(name, third) {
  const letters = (String(name).match(/[a-z]/gi) || []);
  const nums = (String(name).match(/\d+/g) || []).reduce((a, n) => a + parseInt(n, 10), 0);
  const win = letters.slice(0, 2).concat(letters.slice(third, third + 1));
  if (!win.length && !nums) return 0;
  return win.reduce((a, c) => a + c.toUpperCase().charCodeAt(0) - 64, 0) + nums;
}

function nameKey_(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().slice(0, 120);
}

function user_(name) {
  const key = nameKey_(name);
  if (!key || !/[a-z]/.test(key)) throw new Error('Enter your name');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Counter is busy, please try again');
  try {
    const sh = users_();
    const last = sh.getLastRow();
    const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 3).getDisplayValues();
    const mine = rows.find(r => String(r[0]).trim() === key);
    if (mine) return { userNo: String(mine[2]), name: mine[1] };
    const taken = new Set(rows.map(r => String(r[2]).trim()));
    const letters = (String(name).match(/[a-z]/gi) || []).length;
    let no = 0;
    // Keep the first two letters; swap the third for the 4th, 5th … letter until a number is free.
    for (let third = 2; third === 2 || third < letters; third++) {
      const c = userCandidate_(name, third);
      if (c && !taken.has(String(c))) { no = c; break; }
    }
    // Name too short to shift further: take the next free number after the first one.
    if (!no) { no = userCandidate_(name, 2) || 1; while (taken.has(String(no))) no++; }
    sh.appendRow([key, clean_(name, 80), no, new Date()]);
    SpreadsheetApp.flush();
    return { userNo: String(no), name: clean_(name, 80) };
  } finally {
    lock.releaseLock();
  }
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
