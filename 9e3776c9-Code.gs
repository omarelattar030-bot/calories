/**
 * متابعة السعرات - السيرفر (Google Apps Script)
 * يتحط داخل جوجل شيت: Extensions > Apps Script
 *
 * الأول: غيّر CLIENT_ID تحت.
 * بعدين شغّل دالة setup() مرة واحدة (علشان يطلب الصلاحيات ويعمل التابات).
 * بعدين Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

const CLIENT_ID = 'PUT_YOUR_CLIENT_ID.apps.googleusercontent.com';
const SESSION_DAYS = 60;

const AC = ['خامل', 'خفيف', 'متوسط', 'عالي', 'رياضي جدًا'];
const GL = { '-1': 'تنزيل وزن', '0': 'تثبيت وزن', '1': 'زيادة وزن', '2': 'بناء عضلات' };

const USERS = ['email', 'الاسم', 'السن', 'النوع', 'الوزن', 'الطول', 'النشاط', 'الهدف', 'السعرات', 'بروتين', 'كارب', 'دهون', 'فيسبوك', 'واتساب', 'بداية', 'تسجيل', 'آخر تحديث', 'آخر دخول', 'profile_json'];
const DAYS = ['email', 'date', 'json', 'updated'];
const FOODS = ['email', 'json', 'created'];

/* شغّلها مرة واحدة من المحرر */
function setup() {
  sh('Users', USERS);
  sh('Days', DAYS);
  sh('Foods', FOODS);
  UrlFetchApp.fetch('https://www.google.com'); // علشان يطلب صلاحية الاتصال الخارجي
}

function doGet() { return out({ ok: true, msg: 'running' }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    return out(route(b));
  } catch (err) {
    return out({ ok: false, err: String((err && err.message) || err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route(b) {
  if (b.a === 'login') return login(b.cred);
  const email = checkSess(b.s);
  if (!email) return { ok: false, err: 'auth' };
  switch (b.a) {
    case 'load': return withLock(() => { touchUser(email); return Object.assign({ ok: true }, state(email)); });
    case 'profile': return withLock(() => saveProfile(email, b.u, b.start));
    case 'day': return withLock(() => saveDay(email, b.date, b.entries));
    case 'food': return withLock(() => saveFood(email, b.d));
  }
  return { ok: false, err: 'unknown action' };
}

/* ---------- الدخول والجلسات ---------- */

function login(cred) {
  if (!cred) return { ok: false, err: 'مفيش توكن' };
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(cred), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { ok: false, err: 'توكن جوجل غير صالح' };
  const t = JSON.parse(r.getContentText());
  if (t.aud !== CLIENT_ID) return { ok: false, err: 'الـ Client ID مش مطابق' };
  if (t.iss !== 'accounts.google.com' && t.iss !== 'https://accounts.google.com') return { ok: false, err: 'مصدر التوكن غلط' };
  if (String(t.email_verified) !== 'true') return { ok: false, err: 'الإيميل مش متأكد منه' };
  const email = String(t.email).toLowerCase();
  withLock(() => { ensureUser(email, t.name || ''); touchUser(email); });
  return Object.assign({ ok: true, sess: makeSess(email), email: email, name: t.name || '' }, state(email));
}

function secret() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}

function sign(p) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(p, secret()));
}

function makeSess(email) {
  const exp = Date.now() + SESSION_DAYS * 864e5;
  const p = Utilities.base64EncodeWebSafe(email + '|' + exp);
  return p + '.' + sign(p);
}

function checkSess(s) {
  if (!s || typeof s !== 'string') return null;
  const parts = s.split('.');
  if (parts.length !== 2 || sign(parts[0]) !== parts[1]) return null;
  const dec = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString().split('|');
  if (dec.length !== 2 || Date.now() > +dec[1]) return null;
  return dec[0];
}

/* ---------- الشيت ---------- */

function withLock(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}

function sh(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).setNumberFormat('@'); // كله نص علشان التواريخ ماتتحولش
    s.appendRow(headers);
    s.setFrozenRows(1);
  }
  return s;
}

function findRow(s, email) {
  const n = s.getLastRow() - 1;
  if (n < 1) return 0;
  const v = s.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(v[i][0]).toLowerCase() === email) return i + 2;
  return 0;
}

function ensureUser(email, name) {
  const s = sh('Users', USERS);
  let row = findRow(s, email);
  if (!row) {
    const now = new Date().toISOString();
    s.appendRow([email, name, '', '', '', '', '', '', '', '', '', '', '', '', '', now, now, now, '']);
    row = s.getLastRow();
  }
  return row;
}

function touchUser(email) {
  const s = sh('Users', USERS);
  const row = findRow(s, email);
  if (row) s.getRange(row, 18).setValue(new Date().toISOString());
}

function state(email) {
  const us = sh('Users', USERS);
  const row = findRow(us, email);
  let profile = null, start = '';
  if (row) {
    const v = us.getRange(row, 1, 1, 19).getValues()[0];
    start = String(v[14] || '');
    try { profile = v[18] ? JSON.parse(v[18]) : null; } catch (e) {}
  }
  const ds = sh('Days', DAYS);
  const days = {};
  let n = ds.getLastRow() - 1;
  if (n > 0) {
    ds.getRange(2, 1, n, 3).getValues().forEach(r => {
      if (String(r[0]).toLowerCase() === email) { try { days[String(r[1])] = JSON.parse(r[2]); } catch (e) {} }
    });
  }
  const fs = sh('Foods', FOODS);
  const foods = [];
  n = fs.getLastRow() - 1;
  if (n > 0) {
    fs.getRange(2, 1, n, 2).getValues().forEach(r => {
      if (String(r[0]).toLowerCase() === email) { try { foods.push(JSON.parse(r[1])); } catch (e) {} }
    });
  }
  return { profile: profile, start: start, days: days, foods: foods };
}

const str = (x, n) => String(x == null ? '' : x).slice(0, n);
const num = x => (isFinite(+x) ? +x : 0);

function saveProfile(email, u, start) {
  if (!u || typeof u !== 'object') return { ok: false, err: 'بيانات ناقصة' };
  const clean = {
    n: str(u.n, 60), a: num(u.a), s: u.s === 'f' ? 'f' : 'm', w: num(u.w), h: num(u.h),
    act: num(u.act), g: num(u.g), fb: str(u.fb, 200), wa: str(u.wa, 30),
    bmr: num(u.bmr), k: num(u.k), p: num(u.p), c: num(u.c), f: num(u.f)
  };
  const s = sh('Users', USERS);
  const row = ensureUser(email, clean.n);
  s.getRange(row, 2, 1, 13).setValues([[
    clean.n, clean.a, clean.s === 'm' ? 'ذكر' : 'أنثى', clean.w, clean.h,
    AC[clean.act] || '', GL[String(clean.g)] || '', clean.k, clean.p, clean.c, clean.f, clean.fb, clean.wa
  ]]);
  let st = String(s.getRange(row, 15).getValue() || '');
  if (!st && /^\d{4}-\d{2}-\d{2}$/.test(String(start))) { st = String(start); s.getRange(row, 15).setValue(st); }
  s.getRange(row, 17).setValue(new Date().toISOString());
  s.getRange(row, 19).setValue(JSON.stringify(clean));
  return { ok: true, start: st };
}

function saveDay(email, date, entries) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date)) || !Array.isArray(entries)) return { ok: false, err: 'بيانات غلط' };
  const list = entries.slice(0, 80).map(x => ({
    n: str(x.n, 80), q: str(x.q, 30), m: str(x.m, 10),
    k: num(x.k), p: num(x.p), c: num(x.c), f: num(x.f)
  }));
  const json = JSON.stringify(list);
  if (json.length > 45000) return { ok: false, err: 'اليوم كبير زيادة' };
  const s = sh('Days', DAYS);
  const n = s.getLastRow() - 1;
  let row = 0;
  if (n > 0) {
    const v = s.getRange(2, 1, n, 2).getValues();
    for (let i = 0; i < n; i++) {
      if (String(v[i][0]).toLowerCase() === email && String(v[i][1]) === date) { row = i + 2; break; }
    }
  }
  if (!list.length) {
    if (row) s.deleteRow(row);
  } else if (row) {
    s.getRange(row, 3, 1, 2).setValues([[json, new Date().toISOString()]]);
  } else {
    s.appendRow([email, date, json, new Date().toISOString()]);
  }
  return { ok: true };
}

function saveFood(email, d) {
  if (!d || !d.n) return { ok: false, err: 'بيانات ناقصة' };
  const clean = { n: str(d.n, 80), s: str(d.s, 60), k: num(d.k), p: num(d.p), c: num(d.c), f: num(d.f), a: '' };
  sh('Foods', FOODS).appendRow([email, JSON.stringify(clean), new Date().toISOString()]);
  return { ok: true };
}
