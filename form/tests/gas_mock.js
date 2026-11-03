// Имитация Google Таблицы и сервисов Apps Script для тестов apps_script.gs (node).
const fs = require('fs'), path = require('path');
const pad = n => String(n).padStart(2, '0');
function fmtDate(d, f) {
  const m = { yyyy: d.getFullYear(), MM: pad(d.getMonth() + 1), dd: pad(d.getDate()), HH: pad(d.getHours()), mm: pad(d.getMinutes()) };
  if (f === 'd') return String(d.getDate());
  if (f === 'M') return String(d.getMonth() + 1);
  return f.replace(/yyyy|MM|dd|HH|mm/g, k => m[k]);
}
function makeSheet(ss, name) {
  const sh = { name, rows: [],
    getName: () => sh.name, setName: n => { delete ss.sheets[sh.name]; sh.name = n; ss.sheets[n] = sh; },
    getLastRow: () => sh.rows.length, getLastColumn: () => Math.max(0, ...sh.rows.map(r => r.length)),
    appendRow: r => { sh.rows.push(r.slice()); }, deleteRow: i => { sh.rows.splice(i - 1, 1); },
    clear: () => { sh.rows = []; }, setFrozenRows() {}, autoResizeColumns() {},
    getDataRange: () => sh.getRange(1, 1, Math.max(1, sh.rows.length), Math.max(1, sh.getLastColumn())),
    getRange(r, c, nr, nc) {
      if (typeof r === 'string') return { setNumberFormat() {}, setFontWeight() {} };
      nr = nr || 1; nc = nc || 1;
      return { setNumberFormat() {}, setFontWeight() {},
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = (sh.rows[r - 1 + i] || [])[c - 1 + j]; return v === undefined ? '' : v; })),
        setValues: v => v.forEach((row, i) => { const t = sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; row.forEach((x, j) => { t[c - 1 + j] = x; }); }) };
    } };
  return sh;
}
function load(opts = {}) {
  const ss = { sheets: {}, toasts: [] };
  const api = {
    getSheetByName: n => ss.sheets[n] || null, insertSheet: n => (ss.sheets[n] = makeSheet(ss, n)),
    getSpreadsheetTimeZone: () => 'Europe/Moscow', toast: t => ss.toasts.push(t) };
  const g = {
    now: opts.now || new Date(2026, 10, 22, 12, 0), mail: [], cache: {}, quota: 100,
  };
  const RealDate = Date;
  const ctx = {
    SpreadsheetApp: { getActive: () => api, getUi: () => ({}) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: { MimeType: { JSON: 1 }, createTextOutput: t => ({ t, setMimeType() { return this; } }) },
    Utilities: { formatDate: (d, tz, f) => fmtDate(d, f) },
    CacheService: { getScriptCache: () => ({ get: k => g.cache[k] || null, put: (k, v) => { g.cache[k] = v; }, remove: k => { delete g.cache[k]; } }) },
    MailApp: { getRemainingDailyQuota: () => g.quota, sendEmail: (to, s, b, o) => g.mail.push({ to, s, b, o }) },
    console,
    Date: class extends RealDate { constructor(...a) { if (a.length) super(...a); else super(g.now); } },
  };
  const code = fs.readFileSync(path.join(__dirname, '..', 'apps_script.gs'), 'utf8');
  const fn = new Function(...Object.keys(ctx), code + '; return { doGet, doPost, buildFor1C, buildNotSubmitted, setup, settings_ };');
  const G = fn(...Object.values(ctx));
  const csv = f => fs.readFileSync(f, 'utf8').replace(/^﻿/, '').trim().split('\n').map(l => l.split(';'));
  G.ss = ss; G.g = g;
  G.loadCsv = (sheet, file) => { (ss.sheets[sheet] || api.insertSheet(sheet)).rows = csv(file); };
  G.get = params => JSON.parse(G.doGet({ parameter: params || {} }).t);
  G.post = body => JSON.parse(G.doPost({ postData: { contents: JSON.stringify(body) } }).t);
  G.setSetting = (label, value) => { const r = ss.sheets['Настройки'].rows.find(r => r[0] === label); r[1] = value; };
  return G;
}
module.exports = { load };
