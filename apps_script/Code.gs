/**
 * Puss Cue Book — data service (Google Apps Script, bound to a private Google Sheet on Gary's account).
 *
 * The page on GitHub Pages talks to this script; nobody needs an account, only the shared passcode.
 * Nothing is ever lost:
 *   - "cues"  sheet: the current cue list, one row per cue (readable by eye).
 *   - "log"   sheet: APPEND-ONLY, every change ever made (who, when, field, before, after). The cue list can be
 *                    rebuilt from it.
 *   - Google Sheets' own File > Version history.
 *   - a copy of the whole spreadsheet every 6 hours, "Puss Cue Book backup <date>" in My Drive (installBackups()).
 *   - Gary's repo copy: Tools/CueSheet/cue_sync.py webpull pulls everything into git.
 *
 * Writes are safe to retry (each carries an op id; a repeat is acknowledged, not re-applied) and never silently
 * overwrite: a field that someone else changed since you opened it comes back as a conflict for you to decide.
 *
 * Setup: see Tools/CueBook/README.md.
 */

var FIELDS = ['klara', 'gary', 'cameron', 'light', 'soundTech', 'soundCue', 'action', 'videoCue', 'beats', 'notes'];
var META = ['key', 'ord', 'version', 'cut', 'num', 'origin', 'director', 'after', 'updatedBy', 'updatedAt'];
var CUE_COLUMNS = META.concat(FIELDS);
var LOG_COLUMNS = ['at', 'who', 'client', 'op', 'key', 'field', 'before', 'after', 'version'];
var MAX_TEXT = 5000;

function doGet() {
  return json_({ ok: true, service: 'puss-cue-book', hint: 'POST only' });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, code: 'bad_request', message: 'Not JSON' });
  }
  var props = PropertiesService.getScriptProperties();
  var pass = props.getProperty('PASSCODE');
  if (!pass || req.pass !== pass) {
    Utilities.sleep(800);
    return json_({ ok: false, code: 'wrong_passcode', message: 'Wrong passcode' });
  }
  try {
    if (req.action === 'list') return json_(list_());
    if (req.action === 'log') return json_(readLog_(Number(req.limit) || 60));
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (req.action === 'save') return json_(save_(req));
      if (req.action === 'add') return json_(add_(req));
      if (req.action === 'seed') return json_(seed_(req, props));
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: false, code: 'bad_request', message: 'Unknown action ' + req.action });
  } catch (err) {
    return json_({ ok: false, code: 'server_error', message: String(err && err.message || err) });
  }
}

// ------------------------------------------------------------------ reads
function list_() {
  var rows = readCues_().rows;
  rows.sort(function (a, b) { return a.ord - b.ord; });
  return { ok: true, cues: rows, serverTime: Date.now() };
}

function readLog_(limit) {
  var sh = sheet_('log', LOG_COLUMNS);
  var n = sh.getLastRow() - 1;
  if (n <= 0) return { ok: true, log: [] };
  var take = Math.min(limit, n);
  var values = sh.getRange(sh.getLastRow() - take + 1, 1, take, LOG_COLUMNS.length).getValues();
  var out = values.map(function (v) { return toObj_(LOG_COLUMNS, v); }).reverse();
  return { ok: true, log: out };
}

// ------------------------------------------------------------------ writes
function save_(req) {
  var who = clean_(req.who, 60), client = clean_(req.client, 60), op = clean_(req.op, 80);
  if (op && opSeen_(op)) return { ok: true, duplicate: true, cue: findCue_(req.key) };
  var table = readCues_();
  var idx = table.index[req.key];
  if (idx == null) return { ok: false, code: 'not_found', message: 'No cue ' + req.key };
  var cur = table.rows[idx];
  var patch = req.patch || {}, base = req.base || {};
  var conflicts = [], changes = [];
  Object.keys(patch).forEach(function (f) {
    if (FIELDS.indexOf(f) < 0 && f !== 'cut') return;
    var next = f === 'cut' ? !!patch[f] : clean_(patch[f], MAX_TEXT);
    var was = f === 'cut' ? !!cur[f] : String(cur[f] || '');
    var expected = f === 'cut' ? !!base[f] : String(base[f] == null ? '' : base[f]);
    if (was === next) return;                              // already so (another device got there first)
    if (was !== expected) { conflicts.push({ field: f, theirs: was, yours: next, base: expected }); return; }
    changes.push({ field: f, before: was, after: next });
  });
  if (conflicts.length) return { ok: false, code: 'conflict', conflicts: conflicts, cue: cur };
  if (!changes.length) return { ok: true, unchanged: true, cue: cur };
  var now = Date.now();
  changes.forEach(function (c) { cur[c.field] = c.after; });
  cur.version = Number(cur.version || 0) + 1;
  cur.updatedBy = who;
  cur.updatedAt = now;
  writeCue_(table.sheet, idx, cur);
  appendLog_(changes.map(function (c) {
    return { at: now, who: who, client: client, op: op, key: cur.key, field: c.field, before: c.before, after: c.after, version: cur.version };
  }));
  return { ok: true, cue: cur };
}

function add_(req) {
  var who = clean_(req.who, 60), client = clean_(req.client, 60), op = clean_(req.op, 80);
  if (op && opSeen_(op)) return { ok: true, duplicate: true, cue: findCue_(clean_(req.key, 40)) };
  var table = readCues_();
  var key = clean_(req.key, 40);
  if (!/^W[A-Z0-9]{4,12}$/.test(key) || table.index[key] != null) return { ok: false, code: 'bad_request', message: 'Bad new key' };
  var after = table.index[req.after] != null ? table.rows[table.index[req.after]] : null;
  var ord;
  if (after) {
    var sorted = table.rows.slice().sort(function (a, b) { return a.ord - b.ord; });
    var i = sorted.indexOf(after);
    ord = i + 1 < sorted.length ? (after.ord + sorted[i + 1].ord) / 2 : after.ord + 10;
  } else {
    ord = Number(req.ord) || (table.rows.length + 1) * 10;
  }
  var now = Date.now();
  var cue = { key: key, ord: ord, version: 1, cut: false, num: '', origin: 'web', director: '', after: after ? after.key : '',
    updatedBy: who, updatedAt: now };
  FIELDS.forEach(function (f) { cue[f] = clean_((req.fields || {})[f], MAX_TEXT); });
  var sh = table.sheet;
  sh.getRange(sh.getLastRow() + 1, 1, 1, CUE_COLUMNS.length).setValues([CUE_COLUMNS.map(function (c) { return cell_(cue[c]); })]);
  appendLog_([{ at: now, who: who, client: client, op: op, key: key, field: '_new', before: '', after: cue.after, version: 1 }]);
  return { ok: true, cue: cue };
}

/** One-off: load the cue list into an EMPTY cues sheet (refuses otherwise). */
function seed_(req, props) {
  var table = readCues_();
  if (table.rows.length && !req.append) return { ok: false, code: 'not_empty', message: 'cues sheet already has ' + table.rows.length + ' rows' };
  var have = table.index;
  var rows = (req.cues || []).filter(function (c) { return c && c.key && have[c.key] == null; }).map(function (c) {
    var cue = { key: clean_(c.key, 40), ord: Number(c.ord) || 0, version: 1, cut: !!c.cut, num: clean_(c.num, 20),
      origin: clean_(c.origin, 20), director: clean_(c.director, MAX_TEXT), after: '', updatedBy: '', updatedAt: '' };
    FIELDS.forEach(function (f) { cue[f] = clean_(c[f], MAX_TEXT); });
    return CUE_COLUMNS.map(function (col) { return cell_(cue[col]); });
  });
  if (rows.length) table.sheet.getRange(table.sheet.getLastRow() + 1, 1, rows.length, CUE_COLUMNS.length).setValues(rows);
  appendLog_([{ at: Date.now(), who: 'seed', client: '', op: clean_(req.op, 80), key: '', field: '_seed', before: '', after: String(rows.length), version: '' }]);
  return { ok: true, seeded: rows.length };
}

// ------------------------------------------------------------------ backups
/** Run once from the editor: copies the spreadsheet every 6 hours ("Puss Cue Book backup <date>" in My Drive). */
function installBackups() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().everyHours(6).create();
  dailyBackup();
}

// Uses only the Sheets permission (no Drive access, which Google blocks for unverified scripts on some accounts):
// each copy lands in My Drive as "Puss Cue Book backup <date>".
function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stamp = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd HHmm');
  ss.copy('Puss Cue Book backup ' + stamp);
}

// ------------------------------------------------------------------ sheet helpers
function sheet_(name, columns) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readCues_() {
  var sh = sheet_('cues', CUE_COLUMNS);
  var n = sh.getLastRow() - 1;
  var rows = [], index = {};
  if (n > 0) {
    var values = sh.getRange(2, 1, n, CUE_COLUMNS.length).getValues();
    values.forEach(function (v, i) {
      var o = toObj_(CUE_COLUMNS, v);
      o.ord = Number(o.ord) || 0;
      o.version = Number(o.version) || 1;
      o.cut = o.cut === true || o.cut === 'TRUE' || o.cut === 'true';
      o.updatedAt = o.updatedAt ? Number(o.updatedAt) : '';
      FIELDS.concat(['key', 'num', 'origin', 'director', 'after', 'updatedBy']).forEach(function (f) { o[f] = String(o[f] == null ? '' : o[f]); });
      rows.push(o);
      index[o.key] = i;
    });
  }
  return { sheet: sh, rows: rows, index: index };
}

function findCue_(key) {
  var t = readCues_();
  return t.index[key] != null ? t.rows[t.index[key]] : null;
}

function writeCue_(sh, idx, cue) {
  sh.getRange(idx + 2, 1, 1, CUE_COLUMNS.length).setValues([CUE_COLUMNS.map(function (c) { return cell_(cue[c]); })]);
}

function appendLog_(entries) {
  var sh = sheet_('log', LOG_COLUMNS);
  var rows = entries.map(function (e) { return LOG_COLUMNS.map(function (c) { return cell_(e[c]); }); });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_COLUMNS.length).setValues(rows);
}

function opSeen_(op) {
  var sh = sheet_('log', LOG_COLUMNS);
  var n = sh.getLastRow() - 1;
  if (n <= 0) return false;
  var take = Math.min(n, 2000);                           // retries arrive within minutes; the tail is enough
  var col = LOG_COLUMNS.indexOf('op') + 1;
  var ops = sh.getRange(sh.getLastRow() - take + 1, col, take, 1).getValues();
  for (var i = 0; i < ops.length; i++) if (ops[i][0] === op) return true;
  return false;
}

function toObj_(cols, v) {
  var o = {};
  cols.forEach(function (c, i) { o[c] = v[i]; });
  return o;
}

// Text is stored as literal text (the leading apostrophe is Sheets' "this is text" mark and is not part of the
// value): otherwise "=..." becomes a formula and "007" or "1-2" become numbers or dates.
function cell_(v) {
  if (v === true || v === false || typeof v === 'number') return v;
  var s = String(v == null ? '' : v);
  return s ? "'" + s : '';
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/\r\n/g, '\n').slice(0, max);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
