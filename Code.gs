/**
 * English Department Portal · Google Sheet back end
 * Salalah Vocational College · English Department © 2026
 *
 * How to install (once):
 *   1. Create a new Google Sheet, e.g. "English Department Portal Data".
 *   2. Extensions > Apps Script. Delete what is there and paste this whole file. Save.
 *   3. Choose the function "setup" at the top and press Run. Allow the permissions it asks for.
 *   4. Deploy > New deployment > type "Web app".
 *        Execute as: Me      Who has access: Anyone
 *      Press Deploy and copy the Web app URL.
 *   5. Paste that URL into SCRIPT_URL at the top of the portal's index.html.
 *   6. Sign in as "Head of English Department" with the starting password HOD@2026,
 *      then change it (click your name at the top right). This is the only supervisor
 *      account; add every lecturer in Admin > Staff accounts.
 *
 * Uploaded files, pictures and videos are saved in the Drive folder
 * "English Portal Uploads", shared so that anyone with the link can view them.
 *
 * To run the whole portal from this script (no GitHub needed):
 *   - In the Apps Script editor, add an HTML file named "index" (Files > + > HTML),
 *     delete its contents and paste the whole portal index.html. Leave SCRIPT_URL empty.
 *   - Deploy as above. The web app address now opens the portal itself, and it can be
 *     placed on the college website as a link or inside an iframe.
 *
 * After changing this code, use Deploy > Manage deployments > Edit > New version,
 * so the same URL keeps working.
 */

const RULES = {
  "notices": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "calendar": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "courses": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "lessonplans": {
    "add": "all",
    "manage": "coord",
    "viewAll": "coord",
    "ownEdit": true
  },
  "resources": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "library": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "exams": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "assessment": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "scoresheets": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true,
    "owner": "staff",
    "ownFields": [
      "status",
      "link",
      "fileId",
      "fileUrl",
      "fileName",
      "fileType",
      "fileSize"
    ],
    "defaults": {
      "status": "Missing"
    }
  },
  "marks": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "portfolio": {
    "add": "all",
    "manage": "sup",
    "viewAll": "all",
    "ownEdit": true,
    "owner": "teacher",
    "selfOwned": true
  },
  "timetable": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "meetings": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "students": {
    "add": "all",
    "manage": "all",
    "viewAll": "all",
    "defaults": {
      "warning": "None"
    }
  },
  "passive": {
    "add": "all",
    "manage": "all",
    "viewAll": "all",
    "defaults": {
      "status": "Open"
    }
  },
  "studentcorner": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "support": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "faq": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "pd": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "observations": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "plc": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "events": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "achievements": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "gallery": {
    "add": "all",
    "manage": "coord",
    "viewAll": "all",
    "ownEdit": true
  },
  "documents": {
    "add": "coord",
    "manage": "coord",
    "viewAll": "all"
  },
  "feedback": {
    "add": "all",
    "manage": "sup",
    "viewAll": "sup",
    "ownEdit": true,
    "restricted": {
      "status": "sup",
      "reply": "sup"
    },
    "defaults": {
      "status": "New"
    }
  }
};

const SALT = 'SVC-ENGLISH';
const RANK = { lecturer: 1, coordinator: 2, supervisor: 3 };
const NEED = { all: 1, coord: 2, sup: 3 };
const SYSTEM = ['id', 'createdBy', 'createdAt', 'updatedAt'];
const STAFF_HEADERS = ['name', 'title', 'access', 'office', 'ext', 'email', 'duties', 'hash'];
const MAX_TRIES = 5;
const LOCK_SECONDS = 300;
const SESSION_SECONDS = 21600; // 6 hours (the page also signs out after 20 idle minutes)
const MAX_UPLOAD = 20 * 1024 * 1024; // 20 MB per file
const UPLOAD_FOLDER = 'English Portal Uploads';

/* ---------------- menu & setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Portal')
    .addItem('Set up sheets', 'setup')
    .addItem('Back up now', 'backupNow')
    .addItem('Turn on weekly backup', 'installWeeklyBackup')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staff = sheet_('Staff', STAFF_HEADERS);
  sheet_('Log', ['time', 'name', 'result']);
  sheet_('Settings', ['key', 'value']);
  Object.keys(RULES).forEach(function (k) { sheet_(k, SYSTEM); });
  if (staff.getLastRow() < 2) {
    appendText_(staff, ['Head of English Department', 'Head of English Department', 'supervisor', '', '', '', 'Leads the department', hash_('Head of English Department', 'HOD@2026')]);
  }
  folder_('');
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return 'Ready';
}

function backupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'English Portal Backups';
  const folders = DriveApp.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  DriveApp.getFileById(ss.getId()).makeCopy('English Portal backup ' + stamp, folder);
}

function installWeeklyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupNow').timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(15).create();
}

/* ---------------- web app ---------------- */

/* Opening the web app address shows the portal itself when this project has an HTML file
   named "index" (the portal's index.html). Without it, the address only answers the portal. */
function doGet(e) {
  if (!(e && e.parameter && e.parameter.check)) {
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('English Department Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) { /* no index file: fall through */ }
  }
  return json_({ ok: true, message: 'English Department Portal is connected to this Google Sheet.' });
}

/* Called by the portal page when it runs inside Apps Script (google.script.run). */
function api(body) {
  try { return JSON.stringify(handle_(JSON.parse(body))); }
  catch (err) { return JSON.stringify({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  try { return json_(handle_(req)); }
  catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function handle_(req) {
  if (req.action === 'names') { const st = readRows_('Staff'); return { ok: true, names: st.map(function (s) { return s.name; }), staff: st.map(function (s) { return { name: s.name, access: s.access || 'lecturer' }; }) }; }
  if (req.action === 'login') return login_(req);
  const user = auth_(req.token);
  if (!user) return { ok: false, error: 'expired' };
  switch (req.action) {
    case 'load':         return Object.assign({ ok: true }, loadAll_(user));
    case 'save':         return saveRow_(user, req.collection, req.row || {});
    case 'upload':       return upload_(user, req);
    case 'import':       return importRows_(user, req.collection, req.rows || []);
    case 'remove':       return removeRow_(user, req.collection, req.id);
    case 'saveStaff':    return saveStaff_(user, req);
    case 'removeStaff':  return removeStaff_(user, req.name);
    case 'profile':      return profile_(user, req);
    case 'saveSettings': return saveSettings_(user, req.settings || {});
    case 'logout':       CacheService.getScriptCache().remove('t_' + req.token); return { ok: true };
  }
  return { ok: false, error: 'Unknown action' };
}

/* ---------------- sign-in ---------------- */

function login_(req) {
  const name = String(req.name || '');
  const key = norm_(name);
  const cache = CacheService.getScriptCache();
  if (cache.get('lock_' + key)) { log_(name, 'Blocked: too many wrong passwords'); return { ok: false, error: 'locked' }; }
  const staff = readRows_('Staff').filter(function (s) { return s.name === name; })[0];
  if (!staff || staff.hash !== hash_(name, String(req.password || ''))) {
    const n = Number(cache.get('f_' + key) || 0) + 1;
    if (n >= MAX_TRIES) { cache.put('lock_' + key, '1', LOCK_SECONDS); cache.remove('f_' + key); }
    else cache.put('f_' + key, String(n), 900);
    log_(name, 'Wrong password');
    return { ok: false, error: n >= MAX_TRIES ? 'locked' : 'wrong', left: Math.max(0, MAX_TRIES - n) };
  }
  cache.remove('f_' + key);
  const token = Utilities.getUuid();
  cache.put('t_' + token, staff.name, SESSION_SECONDS);
  log_(staff.name, 'Signed in');
  return Object.assign({ ok: true, token: token }, loadAll_(public_(staff)));
}

function auth_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const name = cache.get('t_' + token);
  if (!name) return null;
  const staff = readRows_('Staff').filter(function (s) { return s.name === name; })[0];
  if (!staff) return null;
  cache.put('t_' + token, name, SESSION_SECONDS);
  return public_(staff);
}

function public_(s) {
  return { name: s.name, title: s.title, access: s.access || 'lecturer', office: s.office || '', ext: s.ext || '', email: s.email || '', duties: s.duties || '' };
}

/* ---------------- data ---------------- */

function rank_(user) { return RANK[user.access] || 1; }
function ownerOf_(rule, row) { return row[rule.owner || 'createdBy']; }
function canSee_(user, rule, row) {
  return rank_(user) >= NEED[rule.viewAll] || ownerOf_(rule, row) === user.name || row.createdBy === user.name;
}

function loadAll_(user) {
  const data = {};
  Object.keys(RULES).forEach(function (col) {
    const rule = RULES[col];
    data[col] = readRows_(col).filter(function (r) { return canSee_(user, rule, r); });
  });
  const out = { user: user, data: data, staff: readRows_('Staff').map(public_), settings: readSettings_() };
  if (rank_(user) >= NEED.sup) out.log = readRows_('Log').slice(-300).reverse();
  return out;
}

function saveRow_(user, col, row) {
  const rule = RULES[col];
  if (!rule) throw new Error('Unknown list: ' + col);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(col, SYSTEM);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    let index = -1;
    if (row.id) {
      for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(row.id)) { index = i; break; }
      if (index < 0) throw new Error('This item no longer exists. Reload the page.');
    }
    const existing = index > 0 ? toObj_(headers, values[index]) : null;
    const now = new Date().toISOString();
    const manage = rank_(user) >= NEED[rule.manage];
    let out;
    if (!existing) {
      if (rank_(user) < NEED[rule.add]) throw new Error('Your role cannot add items here.');
      out = {};
      Object.keys(row).forEach(function (k) {
        if (SYSTEM.indexOf(k) >= 0) return;
        if (rule.restricted && rule.restricted[k] && rank_(user) < NEED[rule.restricted[k]]) return;
        out[k] = row[k];
      });
      Object.keys(rule.defaults || {}).forEach(function (k) { if (out[k] === undefined || out[k] === '') out[k] = rule.defaults[k]; });
      if (rule.selfOwned && (!manage || !out[rule.owner])) out[rule.owner] = manage && out[rule.owner] ? out[rule.owner] : user.name;
      out.id = Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
      out.createdBy = (col === 'feedback' && truthy_(row.anonymous)) ? '' : user.name;
      out.createdAt = now;
    } else {
      const isOwner = rule.ownEdit && ownerOf_(rule, existing) === user.name;
      if (!manage && !isOwner) throw new Error('Your role cannot change this item.');
      out = Object.assign({}, existing);
      const allowed = manage ? null : (rule.ownFields || null);
      Object.keys(row).forEach(function (k) {
        if (SYSTEM.indexOf(k) >= 0) return;
        if (allowed && allowed.indexOf(k) < 0) return;
        if (!manage && rule.owner && k === rule.owner) return;
        if (rule.restricted && rule.restricted[k] && rank_(user) < NEED[rule.restricted[k]]) return;
        out[k] = row[k];
      });
    }
    out.updatedAt = now;
    Object.keys(out).forEach(function (k) {
      if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    const vals = headers.map(function (h) { return out[h] === undefined || out[h] === null ? '' : String(out[h]); });
    const rowNo = existing ? index + 1 : sh.getLastRow() + 1;
    sh.getRange(rowNo, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
    return { ok: true, row: out };
  } finally { lock.releaseLock(); }
}

function removeRow_(user, col, id) {
  const rule = RULES[col];
  if (!rule) throw new Error('Unknown list: ' + col);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(col, SYSTEM);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== String(id)) continue;
      const row = toObj_(headers, values[i]);
      const manage = rank_(user) >= NEED[rule.manage];
      const owner = rule.ownEdit && !rule.ownFields && ownerOf_(rule, row) === user.name;
      if (!manage && !owner) throw new Error('Your role cannot delete this item.');
      sh.deleteRow(i + 1);
      return { ok: true };
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------------- import many rows (paste from Excel) ---------------- */

function importRows_(user, col, rows) {
  const rule = RULES[col];
  if (!rule) throw new Error('Unknown list: ' + col);
  if (rank_(user) < NEED[rule.add]) throw new Error('Your role cannot add items here.');
  if (!rows.length) throw new Error('Nothing to import.');
  if (rows.length > 2000) throw new Error('Import up to 2000 rows at a time.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_(col, SYSTEM);
    const headers = sh.getDataRange().getValues()[0].map(String);
    const now = new Date().toISOString(), manage = rank_(user) >= NEED[rule.manage];
    const outs = rows.map(function (row, i) {
      const out = {};
      Object.keys(row).forEach(function (k) {
        if (SYSTEM.indexOf(k) >= 0) return;
        if (rule.restricted && rule.restricted[k] && rank_(user) < NEED[rule.restricted[k]]) return;
        out[k] = row[k];
      });
      Object.keys(rule.defaults || {}).forEach(function (k) { if (out[k] === undefined || out[k] === '') out[k] = rule.defaults[k]; });
      if (rule.selfOwned && (!manage || !out[rule.owner])) out[rule.owner] = user.name;
      out.id = Utilities.getUuid().slice(0, 8) + Date.now().toString(36) + i;
      out.createdBy = user.name; out.createdAt = now; out.updatedAt = now;
      return out;
    });
    outs.forEach(function (o) { Object.keys(o).forEach(function (k) { if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); } }); });
    const vals = outs.map(function (o) { return headers.map(function (h) { return o[h] === undefined || o[h] === null ? '' : String(o[h]); }); });
    sh.getRange(sh.getLastRow() + 1, 1, vals.length, headers.length).setNumberFormat('@').setValues(vals);
    return { ok: true, rows: outs };
  } finally { lock.releaseLock(); }
}

/* ---------------- uploads ---------------- */

function upload_(user, req) {
  if (!RULES[req.collection]) throw new Error('Unknown list: ' + req.collection);
  const bytes = Utilities.base64Decode(String(req.data || ''));
  if (!bytes.length) throw new Error('The file is empty.');
  if (bytes.length > MAX_UPLOAD) throw new Error('Files can be up to 20 MB. For a longer video, upload it to Drive or YouTube and add the link.');
  const name = String(req.name || 'upload').replace(/[\\/:*?"<>|]/g, '_').slice(0, 180);
  const blob = Utilities.newBlob(bytes, String(req.mimeType || 'application/octet-stream'), name);
  const file = folder_(req.collection).createFile(blob);
  file.setDescription('Uploaded by ' + user.name + ' to the English Department Portal');
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* sharing may be limited by the domain */ }
  return { ok: true, fileId: file.getId(), fileUrl: file.getUrl() };
}

function folder_(sub) {
  const roots = DriveApp.getFoldersByName(UPLOAD_FOLDER);
  const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(UPLOAD_FOLDER);
  if (!sub) return root;
  const subs = root.getFoldersByName(sub);
  return subs.hasNext() ? subs.next() : root.createFolder(sub);
}

/* ---------------- staff & settings ---------------- */

function saveStaff_(user, req) {
  if (rank_(user) < NEED.sup) throw new Error('Only the supervisor can manage staff accounts.');
  const s = req.staff || {};
  const name = String(s.name || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('Enter the staff member\'s name.');
  const original = req.original || '';
  const pw = String(req.password || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_('Staff', STAFF_HEADERS);
    STAFF_HEADERS.forEach(function (h) { const hs = sh.getDataRange().getValues()[0].map(String); if (hs.indexOf(h) < 0) sh.getRange(1, hs.length + 1).setValue(h); });
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    let index = -1;
    for (let i = 1; i < values.length; i++) if (String(values[i][0]) === (original || name)) { index = i; break; }
    for (let i = 1; i < values.length; i++) if (i !== index && String(values[i][0]) === name) throw new Error(name + ' already has an account.');
    const existing = index > 0 ? toObj_(headers, values[index]) : null;
    if (!existing && pw.length < 6) throw new Error('Set a password of at least 6 characters for the new account.');
    if (existing && existing.name !== name && pw.length < 6) throw new Error('Set a new password when you change a name.');
    if (pw && pw.length < 6) throw new Error('Use a password of at least 6 characters.');
    const out = {
      name: name, title: s.title || '', access: RANK[s.access] ? s.access : 'lecturer',
      office: s.office || '', ext: s.ext || '', email: s.email || '', duties: s.duties || '',
      hash: pw ? hash_(name, pw) : existing.hash
    };
    const vals = headers.map(function (h) { return out[h] === undefined ? '' : String(out[h]); });
    const rowNo = existing ? index + 1 : sh.getLastRow() + 1;
    sh.getRange(rowNo, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
    return { ok: true, staff: readRows_('Staff').map(public_) };
  } finally { lock.releaseLock(); }
}

function removeStaff_(user, name) {
  if (rank_(user) < NEED.sup) throw new Error('Only the supervisor can manage staff accounts.');
  if (name === user.name) throw new Error('You cannot remove your own account.');
  const sh = sheet_('Staff', STAFF_HEADERS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === name) { sh.deleteRow(i + 1); break; }
  return { ok: true, staff: readRows_('Staff').map(public_) };
}

function profile_(user, req) {
  const sh = sheet_('Staff', STAFF_HEADERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) !== user.name) continue;
    const row = toObj_(headers, values[i]);
    if (req.newPassword) {
      if (row.hash !== hash_(user.name, String(req.current || ''))) throw new Error('Your current password is not correct.');
      if (String(req.newPassword).length < 6) throw new Error('Use a new password of at least 6 characters.');
      row.hash = hash_(user.name, String(req.newPassword));
    }
    ['office', 'ext', 'email'].forEach(function (k) { if (req[k] !== undefined) row[k] = req[k]; });
    const vals = headers.map(function (h) { return row[h] === undefined ? '' : String(row[h]); });
    sh.getRange(i + 1, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
    return { ok: true, user: public_(row), staff: readRows_('Staff').map(public_) };
  }
  throw new Error('Account not found.');
}

function readSettings_() {
  const out = {};
  readRows_('Settings').forEach(function (r) {
    try { out[r.key] = JSON.parse(r.value); } catch (e) { out[r.key] = r.value; }
  });
  return out;
}

function saveSettings_(user, settings) {
  if (rank_(user) < NEED.sup) throw new Error('Only the supervisor can change portal options.');
  const sh = sheet_('Settings', ['key', 'value']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(settings).map(function (k) { return [k, JSON.stringify(settings[k])]; });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  return { ok: true, settings: readSettings_() };
}

/* ---------------- helpers ---------------- */

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(function (r) { return String(r[0]) !== ''; }).map(function (r) { return toObj_(headers, r); });
}

function toObj_(headers, row) {
  const o = {};
  headers.forEach(function (h, i) {
    if (!h) return;
    let v = row[i];
    if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
    o[h] = v === null || v === undefined ? '' : String(v);
  });
  return o;
}

function appendText_(sh, vals) {
  sh.getRange(sh.getLastRow() + 1, 1, 1, vals.length).setNumberFormat('@').setValues([vals.map(String)]);
}

function log_(name, result) {
  appendText_(sheet_('Log', ['time', 'name', 'result']), [new Date().toISOString(), name, result]);
}

function norm_(n) { return String(n).trim().replace(/\s+/g, ' ').toLowerCase(); }
function truthy_(v) { return v === true || v === 'true' || v === 'TRUE' || v === 'Yes'; }

function hash_(name, pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + '|' + norm_(name) + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
