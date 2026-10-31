/**
 * DGM Property Map — Google Sheets backend
 *
 * Paste this whole file into Extensions → Apps Script of the
 * "DGM Property Tracker" spreadsheet, then Deploy → New deployment →
 * Web app (Execute as: Me, Who has access: Anyone).
 *
 * Every request must carry the passphrase stored in Settings!B1.
 * Do not rename the tabs or the header row of the data tabs.
 */

var TAB = { props: 'Properties', ten: 'Tenants', sales: 'Sales', notes: 'Notes', settings: 'Settings' };

// Properties columns <-> record fields used by the map page
var PROP_MAP = [
  ['ParcelID', 'pid'], ['Address', 'addr'], ['Owner', 'owner'], ['Stage', 'stage'], ['Market', 'market'],
  ['Asking', 'asking', 'num'], ['ValueEst', 'valueEst', 'num'], ['LeasableSF', 'sfLease', 'num'],
  ['LeaseRate', 'rate', 'num'], ['LeaseType', 'leaseType'], ['Insurance', 'ins', 'num'],
  ['Score_Price', 'score.price', 'num'], ['Score_Condition', 'score.cond', 'num'],
  ['Score_Affordability', 'score.afford', 'num'], ['Score_Owner', 'score.owner', 'num'],
  ['Score_Impact', 'score.impact', 'num'],
  ['Why_Price', 'scoreNotes.price'], ['Why_Condition', 'scoreNotes.cond'],
  ['Why_Affordability', 'scoreNotes.afford'], ['Why_Owner', 'scoreNotes.owner'], ['Why_Impact', 'scoreNotes.impact'],
  ['ContactName', 'contact.name'], ['ContactRole', 'contact.role'], ['ContactPhone', 'contact.phone'],
  ['ContactEmail', 'contact.email'], ['ContactNotes', 'contact.note'],
  ['CalcDown', 'calc.down', 'num'], ['CalcRate', 'calc.rate', 'num'], ['CalcYears', 'calc.yrs', 'num'],
  ['UpdatedBy', 'updatedBy'], ['UpdatedAt', 'updatedAt', 'num']
];
var TEN_MAP = [['Business', 'name'], ['SqFt', 'sf'], ['RatePerSfMo', 'rate'], ['LeaseEnds', 'ends', 'date'], ['Notes', 'note']];
var SALE_MAP = [['Date', 'date', 'date'], ['Price', 'price'], ['Notes', 'note']];

function doGet() {
  return ContentService.createTextOutput('DGM Property Map API is running. The map page talks to this URL.');
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  var pass = String(sheet_(TAB.settings).getRange('B1').getValue()).trim();
  if (!pass || String(req.pass || '').trim() !== pass) return out_({ ok: false, error: 'bad_passphrase' });
  try {
    switch (req.action) {
      case 'list':   return out_({ ok: true, data: listAll_() });
      case 'save':   return out_(locked_(function () { return save_(req.record, req.base, req.by); }));
      case 'delete': return out_(locked_(function () { return remove_(req.pid); }));
      case 'note':   return out_(locked_(function () { return note_(req.pid, req.text, req.by); }));
    }
    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: 'server', message: String(err) });
  }
}

// ---------- actions ----------
function listAll_() {
  var props = {}, t = table_(TAB.props);
  t.rows.forEach(function (r) { if (r.ParcelID) props[r.ParcelID] = toRecord_(r); });
  attachChildren_(props);
  var notes = table_(TAB.notes).rows.filter(function (r) { return r.ParcelID && r.Text; }).map(function (r) {
    return { pid: String(r.ParcelID), by: String(r.Author || ''), at: toMs_(r.Date), text: String(r.Text) };
  });
  return { props: props, notes: notes };
}

function save_(rec, base, by) {
  if (!rec || !rec.pid) return { ok: false, error: 'bad_request' };
  var t = table_(TAB.props), pid = String(rec.pid);
  var existing = t.rows.filter(function (r) { return String(r.ParcelID) === pid; })[0];
  if (existing) {
    var cur = Number(existing.UpdatedAt) || 0;
    if (base == null || Number(base) !== cur) {
      var current = toRecord_(existing), one = {}; one[pid] = current; attachChildren_(one);
      return { ok: false, error: 'conflict', current: current };
    }
  }
  rec.updatedBy = String(by || rec.updatedBy || '').slice(0, 80);
  rec.updatedAt = Date.now();
  var rowVals = t.headers.map(function (h) {
    if (h === 'Overall') return null; // formula, filled below
    var m = PROP_MAP.filter(function (x) { return x[0] === h; })[0];
    if (!m) return existing ? existing[h] : '';
    var v = getPath_(rec, m[1]);
    return v == null ? '' : v;
  });
  var rowNum = existing ? existing._row : t.sh.getLastRow() + 1;
  t.sh.getRange(rowNum, 1, 1, rowVals.length).setValues([rowVals.map(function (v) { return v == null ? '' : v; })]);
  var oc = t.headers.indexOf('Overall'), s1 = t.headers.indexOf('Score_Price'), s5 = t.headers.indexOf('Score_Impact');
  if (oc >= 0 && s1 >= 0 && s5 >= 0) {
    t.sh.getRange(rowNum, oc + 1).setFormula('=IFERROR(AVERAGE(' + col_(s1 + 1) + rowNum + ':' + col_(s5 + 1) + rowNum + '),"")');
  }
  replaceChildren_(TAB.ten, pid, rec.tenants || [], TEN_MAP);
  replaceChildren_(TAB.sales, pid, rec.sales || [], SALE_MAP);
  var saved = {}; saved[pid] = toRecord_(table_(TAB.props).rows.filter(function (r) { return String(r.ParcelID) === pid; })[0]);
  attachChildren_(saved);
  return { ok: true, record: saved[pid] };
}

function remove_(pid) {
  pid = String(pid || '');
  [TAB.props, TAB.ten, TAB.sales].forEach(function (name) { deleteWhere_(name, pid); });
  return { ok: true };
}

function note_(pid, text, by) {
  text = String(text || '').trim();
  if (!pid || !text) return { ok: false, error: 'bad_request' };
  var t = table_(TAB.notes), now = Date.now();
  var row = t.headers.map(function (h) {
    return h === 'ParcelID' ? String(pid) : h === 'Author' ? String(by || '').slice(0, 80) : h === 'Date' ? new Date(now) : h === 'Text' ? text.slice(0, 5000) : '';
  });
  t.sh.appendRow(row);
  return { ok: true, note: { pid: String(pid), by: String(by || ''), at: now, text: text } };
}

// ---------- helpers ----------
function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing tab: ' + name);
  return sh;
}
function table_(name) {
  var sh = sheet_(name), vals = sh.getDataRange().getValues();
  var headers = (vals[0] || []).map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var o = { _row: i + 1 };
    headers.forEach(function (h, j) { o[h] = vals[i][j]; });
    rows.push(o);
  }
  return { sh: sh, headers: headers, rows: rows };
}
function toRecord_(r) {
  var rec = { score: {}, scoreNotes: {}, contact: {}, calc: {}, tenants: [], sales: [] };
  PROP_MAP.forEach(function (m) {
    var v = r[m[0]];
    if (v === '' || v == null) return;
    if (m[2] === 'num') { v = Number(v); if (isNaN(v)) return; }
    else v = v instanceof Date ? fmtDate_(v) : String(v);
    setPath_(rec, m[1], v);
  });
  return rec;
}
function attachChildren_(props) {
  [[TAB.ten, 'tenants', TEN_MAP], [TAB.sales, 'sales', SALE_MAP]].forEach(function (cfg) {
    table_(cfg[0]).rows.forEach(function (r) {
      var p = props[String(r.ParcelID)];
      if (!p) return;
      var o = {};
      cfg[2].forEach(function (m) { var v = r[m[0]]; o[m[1]] = v instanceof Date ? fmtDate_(v) : (v == null ? '' : String(v)); });
      p[cfg[1]].push(o);
    });
  });
}
function replaceChildren_(name, pid, items, map) {
  deleteWhere_(name, pid);
  var t = table_(name);
  var rows = items.filter(function (it) { return map.some(function (m) { return String(it[m[1]] || '').trim(); }); }).map(function (it) {
    return t.headers.map(function (h) {
      if (h === 'ParcelID') return pid;
      var m = map.filter(function (x) { return x[0] === h; })[0];
      return m ? String(it[m[1]] == null ? '' : it[m[1]]).slice(0, 2000) : '';
    });
  });
  if (rows.length) t.sh.getRange(t.sh.getLastRow() + 1, 1, rows.length, t.headers.length).setValues(rows);
}
function deleteWhere_(name, pid) {
  var t = table_(name);
  for (var i = t.rows.length - 1; i >= 0; i--) if (String(t.rows[i].ParcelID) === pid) t.sh.deleteRow(t.rows[i]._row);
}
function locked_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}
function getPath_(o, p) { return p.split('.').reduce(function (a, k) { return a == null ? null : a[k]; }, o); }
function setPath_(o, p, v) { var ks = p.split('.'), last = ks.pop(); ks.forEach(function (k) { o = o[k] = o[k] || {}; }); o[last] = v; }
function col_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function fmtDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function toMs_(v) { return v instanceof Date ? v.getTime() : (Number(v) || Date.parse(v) || 0); }
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
