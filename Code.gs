/**
 * Installation Tracker - Backend (Google Apps Script)
 *
 * SETUP
 * 1. Run setup() once (creates sheets + default admin: admin / 1234  -> change it!)
 * 2. Project Settings > Script properties: add BOT_TOKEN and CHAT_ID (Telegram)
 * 3. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
 *    Copy the Web app URL into API_URL in index.html (hosted on GitHub Pages).
 */

const TZ = 'Asia/Phnom_Penh';
const SHEET_INST = 'Installations';
const SHEET_USERS = 'Users';
const SESSION_TTL = 6 * 60 * 60; // seconds (CacheService max is 6h)
const MAX_PHOTOS = 10;

const INST_HEADERS = ['Timestamp', 'User', 'Business Name', 'Merchant Name', 'MID', 'Remarks', 'Photos'];
const USER_HEADERS = ['Username', 'PIN Hash', 'Role', 'Topic ID'];

/* ------------------------------ Web app ------------------------------ */

// Public functions the website may call (everything else is blocked).
const API_METHODS = {
  login: login, checkSession: checkSession, logout: logout,
  submitInstallations: submitInstallations, getSummary: getSummary, getRecords: getRecords,
  listUsers: listUsers, saveUser: saveUser, deleteUser: deleteUser
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, service: 'Installation Tracker API' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API_METHODS[req.fn];
    if (!fn) throw new Error('Unknown action.');
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/* ------------------------------ Setup ------------------------------ */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inst = ss.getSheetByName(SHEET_INST);
  if (!inst) inst = ss.insertSheet(SHEET_INST);
  if (inst.getLastRow() === 0) {
    inst.getRange(1, 1, 1, INST_HEADERS.length).setValues([INST_HEADERS]).setFontWeight('bold');
    inst.setFrozenRows(1);
  }
  inst.getRange('A:A').setNumberFormat('@'); // keep timestamps as text
  inst.getRange('E:E').setNumberFormat('@'); // keep long MIDs intact

  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) users = ss.insertSheet(SHEET_USERS);
  if (users.getLastRow() === 0) {
    users.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight('bold');
    users.setFrozenRows(1);
    users.getRange('B:B').setNumberFormat('@');
    users.appendRow(['admin', hashPin_('admin', '1234'), 'Admin', '']);
  }
  Logger.log('Setup complete. Default login: admin / 1234 (change it right away).');
}

/**
 * RUN ONCE: saves the Telegram settings into Script properties.
 * After it succeeds, DELETE the two values below (set them to '') so the token isn't left in your code.
 */
function configureTelegram() {
  const BOT_TOKEN = '8998303518:AAHU5bIcNj3ApJ2OWjxDrWs-HtfjwZyDJyw';
  const CHAT_ID = '-1004442156087';
  PropertiesService.getScriptProperties().setProperties({ BOT_TOKEN: BOT_TOKEN, CHAT_ID: CHAT_ID });
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/getChat?chat_id=' + encodeURIComponent(CHAT_ID), { muteHttpExceptions: true });
  Logger.log('Saved. Telegram check: ' + res.getContentText());
}

/* ------------------------------ Helpers ------------------------------ */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setup() first.');
  return sh;
}

function hashPin_(username, pin) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    'it-salt|' + String(username).toLowerCase() + '|' + String(pin)
  );
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 200);
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function readUsers_() {
  const sh = sheet_(SHEET_USERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(function (r, i) {
    return { row: i + 2, username: String(r[0]), hash: String(r[1]), role: String(r[2] || 'User'), topicId: String(r[3] || '') };
  }).filter(function (u) { return u.username; });
}

function findUser_(username) {
  const key = String(username).toLowerCase();
  const list = readUsers_();
  for (var i = 0; i < list.length; i++) if (list[i].username.toLowerCase() === key) return list[i];
  return null;
}

function publicUser_(u) {
  return { username: u.username, role: u.role, topicId: u.topicId };
}

/** Validates a session token and returns the live user record. */
function auth_(token, adminOnly) {
  const raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (!raw) throw new Error('SESSION_EXPIRED');
  const u = findUser_(JSON.parse(raw).u);
  if (!u) throw new Error('SESSION_EXPIRED');
  if (adminOnly && u.role !== 'Admin') throw new Error('Admins only.');
  return u;
}

/* ------------------------------ Auth ------------------------------ */

function login(username, pin) {
  username = clean_(username, 40);
  pin = clean_(pin, 20);
  if (!username || !pin) throw new Error('Enter username and PIN.');

  const cache = CacheService.getScriptCache();
  const failKey = 'f_' + username.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Try again in 5 minutes.');

  const u = findUser_(username);
  if (!u || u.hash !== hashPin_(u.username, pin)) {
    cache.put(failKey, String(fails + 1), 300);
    throw new Error('Wrong username or PIN.');
  }
  cache.remove(failKey);

  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify({ u: u.username }), SESSION_TTL);
  return { token: token, user: publicUser_(u) };
}

function checkSession(token) {
  return { user: publicUser_(auth_(token)) };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return true;
}

/* ------------------------------ Installations ------------------------------ */

function submitInstallations(token, items) {
  const user = auth_(token);
  if (!items || !items.length) throw new Error('Nothing to submit.');
  if (items.length > 20) throw new Error('Too many merchants in one batch (max 20).');

  const sh = sheet_(SHEET_INST);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];
  const clean = items.map(function (it) {
    const c = {
      business: clean_(it.business, 120) || 'N/A',
      merchant: clean_(it.merchant, 120),
      mid: clean_(it.mid, 60),
      remarks: clean_(it.remarks, 500),
      photos: (it.photos || []).slice(0, MAX_PHOTOS)
    };
    if (!c.merchant && !c.mid) throw new Error('Each merchant needs a Merchant Name or MID.');
    rows.push([now, user.username, c.business, c.merchant, c.mid, c.remarks, c.photos.length]);
    return c;
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, INST_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  var tgOk = 0, tgFail = 0, tgSkipped = false;
  clean.forEach(function (c) {
    const r = sendTelegram_(user, c, now);
    if (r === 'skipped') tgSkipped = true;
    else if (r === 'ok') tgOk++;
    else tgFail++;
  });

  return { saved: rows.length, telegram: tgSkipped ? 'not configured' : (tgFail ? tgOk + ' sent, ' + tgFail + ' failed' : 'sent') };
}

function sendTelegram_(user, c, ts) {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('BOT_TOKEN');
  const chatId = props.getProperty('CHAT_ID');
  if (!botToken || !chatId) return 'skipped';

  const caption =
    '🛠 <b>New Installation</b>\n' +
    '👤 Staff: ' + esc_(user.username) + '\n' +
    '🏪 Business: ' + esc_(c.business) + '\n' +
    '🧾 Merchant: ' + esc_(c.merchant || '-') + '\n' +
    '🆔 MID: ' + esc_(c.mid || '-') + '\n' +
    (c.remarks ? '📝 Remarks: ' + esc_(c.remarks) + '\n' : '') +
    '🕒 ' + ts;

  const base = { chat_id: chatId, parse_mode: 'HTML' };
  if (/^\d+$/.test(user.topicId)) base.message_thread_id = user.topicId;

  const api = 'https://api.telegram.org/bot' + botToken + '/';
  try {
    var url, payload;
    if (!c.photos.length) {
      url = api + 'sendMessage';
      payload = Object.assign({}, base, { text: caption });
    } else if (c.photos.length === 1) {
      url = api + 'sendPhoto';
      payload = Object.assign({}, base, { caption: caption, photo: toBlob_(c.photos[0], 0) });
    } else {
      url = api + 'sendMediaGroup';
      const media = [];
      payload = Object.assign({}, base);
      delete payload.parse_mode;
      c.photos.forEach(function (p, i) {
        const m = { type: 'photo', media: 'attach://file' + i };
        if (i === 0) { m.caption = caption; m.parse_mode = 'HTML'; }
        media.push(m);
        payload['file' + i] = toBlob_(p, i);
      });
      payload.media = JSON.stringify(media);
    }
    const res = UrlFetchApp.fetch(url, { method: 'post', payload: payload, muteHttpExceptions: true });
    return JSON.parse(res.getContentText()).ok ? 'ok' : 'failed';
  } catch (e) {
    Logger.log('Telegram error: ' + e);
    return 'failed';
  }
}

function toBlob_(dataUrl, i) {
  const m = /^data:(image\/[a-z+]+);base64,(.+)$/i.exec(dataUrl || '');
  if (!m) throw new Error('Invalid photo data.');
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'photo' + i + '.jpg');
}

function readInstallations_() {
  const sh = sheet_(SHEET_INST);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, INST_HEADERS.length).getValues().map(function (r) {
    return {
      ts: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
      user: String(r[1]),
      business: String(r[2]),
      merchant: String(r[3]),
      mid: String(r[4]),
      remarks: String(r[5]),
      photos: Number(r[6]) || 0
    };
  });
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

/** Banner numbers: own count today for users, everyone's for admins. */
function getSummary(token) {
  const user = auth_(token);
  const day = today_();
  const all = readInstallations_();
  const isAdmin = user.role === 'Admin';
  const scoped = isAdmin ? all : all.filter(function (r) { return r.user.toLowerCase() === user.username.toLowerCase(); });
  return {
    today: scoped.filter(function (r) { return r.ts.slice(0, 10) === day; }).length,
    total: scoped.length,
    isAdmin: isAdmin,
    date: day
  };
}

/** filter = { user, from, to } with dates as yyyy-MM-dd (all optional). */
function getRecords(token, filter) {
  const me = auth_(token);
  const isAdmin = me.role === 'Admin';
  filter = filter || {};
  var rows = readInstallations_();

  if (!isAdmin) rows = rows.filter(function (r) { return r.user.toLowerCase() === me.username.toLowerCase(); });
  else if (filter.user) rows = rows.filter(function (r) { return r.user.toLowerCase() === String(filter.user).toLowerCase(); });

  if (filter.from) rows = rows.filter(function (r) { return r.ts.slice(0, 10) >= filter.from; });
  if (filter.to) rows = rows.filter(function (r) { return r.ts.slice(0, 10) <= filter.to; });

  rows.sort(function (a, b) { return a.ts < b.ts ? 1 : -1; });

  const byUser = {};
  rows.forEach(function (r) { byUser[r.user] = (byUser[r.user] || 0) + 1; });

  return {
    rows: rows.slice(0, 2000),
    count: rows.length,
    byUser: byUser,
    users: isAdmin ? readUsers_().map(function (u) { return u.username; }) : []
  };
}

/* ------------------------------ User admin ------------------------------ */

function listUsers(token) {
  auth_(token, true);
  return readUsers_().map(publicUser_);
}

/** u = { original, username, pin, role, topicId }. `original` set when editing. */
function saveUser(token, u) {
  const me = auth_(token, true);
  const username = clean_(u.username, 40);
  const pin = clean_(u.pin, 20);
  const role = u.role === 'Admin' ? 'Admin' : 'User';
  const topicId = clean_(u.topicId, 20);

  if (!/^[A-Za-z0-9_.-]{2,40}$/.test(username)) throw new Error('Username: 2-40 letters, numbers, . _ -');
  if (topicId && !/^\d+$/.test(topicId)) throw new Error('Topic ID must be a number.');

  const sh = sheet_(SHEET_USERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existing = u.original ? findUser_(u.original) : null;
    const clash = findUser_(username);
    if (clash && (!existing || clash.row !== existing.row)) throw new Error('That username already exists.');

    if (existing) {
      if (existing.username.toLowerCase() === me.username.toLowerCase() && role !== 'Admin') {
        throw new Error("You can't remove your own admin role.");
      }
      if (pin && !/^\d{4,12}$/.test(pin)) throw new Error('PIN must be 4-12 digits.');
      const hash = pin ? hashPin_(username, pin) : (username.toLowerCase() === existing.username.toLowerCase() ? existing.hash : null);
      if (!hash) throw new Error('Enter a new PIN when renaming a user.');
      sh.getRange(existing.row, 1, 1, 4).setValues([[username, hash, role, topicId]]);
    } else {
      if (!/^\d{4,12}$/.test(pin)) throw new Error('PIN must be 4-12 digits.');
      sh.appendRow([username, hashPin_(username, pin), role, topicId]);
    }
  } finally {
    lock.releaseLock();
  }
  return listUsers(token);
}

function deleteUser(token, username) {
  const me = auth_(token, true);
  if (String(username).toLowerCase() === me.username.toLowerCase()) throw new Error("You can't delete yourself.");
  const target = findUser_(username);
  if (!target) throw new Error('User not found.');
  sheet_(SHEET_USERS).deleteRow(target.row);
  return listUsers(token);
}
