/**
 * LabAlert - ติดตามผล Lab ผู้ป่วยอายุรกรรม
 * ฐานข้อมูล: Google Sheets (ชีต Patients, FollowUps, Staff, Log)
 * เข้าสู่ระบบด้วยรหัสพนักงาน + รหัสผ่านของตัวเอง (ชีต Staff) บันทึก Time stamp เป็นรหัสพนักงาน
 * ความลับทั้งหมดเก็บใน Script Properties (ไม่ฝังในโค้ด):
 *   LINE_TOKEN        Channel access token ของบอท (Messaging API)
 *   GROUP_ID          ID กลุ่ม LINE ที่รับแจ้งเตือน
 *   ADMIN_CODES       (ไม่บังคับ) รหัสพนักงานที่เป็นแอดมิน คั่นด้วยจุลภาค ค่าเริ่มต้น 54150
 *   APP_URL           (ไม่บังคับ) ลิงก์หน้าแอปที่ปุ่มใน LINE เปิด ค่าเริ่มต้น https://kung54150.github.io/LabAlert/
 *   LINE_SETUP        (ชั่วคราว) ตั้งเป็น 1 ตอนหา ID กลุ่ม แล้วลบทิ้ง
 *   LIFF_ID           (ไม่บังคับ) LIFF ID ถ้าอยากให้ปุ่มเปิดในแอป LINE
 *   LOGIN_CHANNEL_ID  Channel ID ของ LINE Login channel ที่ LIFF สังกัดอยู่
 *   SHEET_ID          (ไม่บังคับ) ถ้าไม่ได้ผูกสคริปต์กับชีต
 */
const REPEAT_MIN = 10;          // แจ้งซ้ำเมื่อไม่มีใครตอบ
const STALE_HOURS = 24;         // รายการที่ยังไม่พิมพ์ผลเกินกี่ชั่วโมงให้ย้ายไปแท็บเสร็จสิ้นของวันที่ลงทะเบียน และหยุดแจ้งเตือน
const SCHEMA = {
  Patients: ['id', 'queue', 'hn', 'registered_at', 'wait_min', 'next_alert_at', 'alert_count',
             'last_alert_at', 'registered_by', 'printed_at', 'printed_by'],
  FollowUps: ['id', 'patient_id', 'n', 'at', 'by', 'via', 'ext_min'],
  Staff: ['code', 'name', 'role', 'pin_hash', 'salt', 'temp_pin', 'updated_at'],
  Log: ['at', 'who', 'action', 'detail'],
};

/* ---------- helpers ---------- */
const prop_ = (k) => PropertiesService.getScriptProperties().getProperty(k) || '';
const ms_ = (v) => (v instanceof Date ? v.getTime() : v ? new Date(v).getTime() : 0);

function ss_() {
  const id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  const ss = ss_();
  if (!ss) throw new Error('ไม่พบชีต: ให้สร้างสคริปต์จากเมนู Extensions > Apps Script ในไฟล์ Google Sheet หรือตั้ง SHEET_ID');
  let sh = ss.getSheetByName(name);
  if (!sh) { // สร้างแท็บที่ขาดให้เองถ้ายังไม่ได้รัน setup
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === 'Patients') sh.getRange('C:C').setNumberFormat('@');
    if (name === 'Staff') sh.getRange('A:F').setNumberFormat('@');
  }
  return sh;
}

function readAll_(name) {
  const sh = sheet_(name), vals = sh.getDataRange().getValues();
  const head = vals[0];
  return vals.slice(1).map((row, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, c) => (o[h] = row[c]));
    return o;
  });
}
function append_(name, obj) {
  const row = SCHEMA[name].map((k) => (obj[k] === undefined ? '' : obj[k]));
  sheet_(name).appendRow(row);
}
function update_(name, rowNo, patch) {
  const sh = sheet_(name), head = SCHEMA[name];
  Object.keys(patch).forEach((k) => sh.getRange(rowNo, head.indexOf(k) + 1).setValue(patch[k]));
}
function log_(who, action, detail) {
  append_('Log', { at: new Date(), who: who, action: action, detail: detail || '' });
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function assertOwner_() {
  const a = Session.getActiveUser().getEmail(), e = Session.getEffectiveUser().getEmail();
  if (!a || a !== e) throw new Error('เฉพาะเจ้าของสคริปต์เท่านั้น');
}

/* ---------- ติดตั้งครั้งแรก: รันจาก Apps Script editor ---------- */
function setup() {
  assertOwner_();
  const ss = ss_();
  Object.keys(SCHEMA).forEach((name) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  sheet_('Patients').getRange('C:C').setNumberFormat('@'); // HN เป็นข้อความ
  sheet_('Staff').getRange('A:F').setNumberFormat('@');    // รหัสพนักงาน/รหัสผ่านเป็นข้อความ (ไม่ตัดเลข 0 นำหน้า)
  ensureAdmins_();
  ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'checkAlerts')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkAlerts').timeBased().everyMinutes(1).create();
  return 'ติดตั้งเสร็จ: สร้างชีตและตัวตั้งเวลาทุก 1 นาทีแล้ว ต่อไปกรอกรายชื่อพนักงานในชีต Staff แล้วรัน issueTempPins';
}

/* ---------- บัญชีพนักงาน (ชีต Staff: code, name, role, pin_hash, salt, temp_pin, updated_at) ---------- */
const adminCodes_ = () => (prop_('ADMIN_CODES') || '54150').split(',').map((x) => x.trim()).filter(Boolean);
function ensureAdmins_() {
  const have = readAll_('Staff').map((r) => String(r.code).trim());
  adminCodes_().forEach((c) => { if (have.indexOf(c) < 0) append_('Staff', { code: c, name: '', role: 'Admin' }); });
}
function staffByCode_(code) { return readAll_('Staff').filter((r) => String(r.code).trim() === code)[0] || null; }
function roleOf_(st, code) {
  return adminCodes_().indexOf(code) >= 0 || (st && String(st.role).trim().toLowerCase() === 'admin') ? 'Admin' : 'Staff';
}
function hashPin_(pin, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin));
}
/** สร้างรหัสผ่านชั่วคราว 6 หลักให้พนักงานที่ยังไม่มีรหัส (รันจาก editor) ดูรหัสได้ที่ชีต Staff คอลัมน์ temp_pin พนักงานจะถูกบังคับตั้งรหัสใหม่เมื่อเข้าครั้งแรก */
function issueTempPins() {
  assertOwner_();
  ensureAdmins_();
  let n = 0;
  readAll_('Staff').forEach((r) => {
    if (String(r.code).trim() && !String(r.pin_hash || '') && !String(r.temp_pin || '')) {
      const pin = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 900000 + 100000);
      update_('Staff', r._row, { temp_pin: pin, updated_at: new Date() });
      n++;
    }
  });
  return 'สร้างรหัสผ่านชั่วคราวให้ ' + n + ' คน ดูที่ชีต Staff คอลัมน์ temp_pin';
}

/* ---------- Web API (หน้าเว็บอยู่บน GitHub Pages เรียกผ่าน fetch แบบ POST) ---------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const CODE_VERSION = 'line-setup-2';
function doGet() { return json_({ ok: true, data: 'LabAlert API', v: CODE_VERSION }); }
function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (Array.isArray(req.events)) return json_(lineSetupEvents_(req.events));   // เหตุการณ์จาก LINE Webhook (ใช้ตอนหา ID กลุ่ม)
    const table = { apiLogin: apiLogin, apiLogout: apiLogout, apiChangePin: apiChangePin, apiConfig: apiConfig, apiList: apiList, apiAdd: apiAdd, apiFollow: apiFollow, apiPrint: apiPrint };
    const fn = Object.prototype.hasOwnProperty.call(table, req.fn) ? table[req.fn] : null;
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง');
    return json_({ ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}
function apiConfig() { return { liffId: prop_('LIFF_ID') }; }

/**
 * ใช้ตอนตั้งค่า LINE ครั้งแรกเท่านั้น: ตั้ง Script Property LINE_SETUP=1 แล้วชี้ Webhook ของบอทมาที่ลิงก์ Web app นี้ชั่วคราว
 * พิมพ์ข้อความในกลุ่ม ระบบจะจด ID กลุ่มลงชีต Log (action = line_group_seen) ให้ผู้ดูแลคัดลอกไปใส่ GROUP_ID เอง
 * ไม่ตั้งค่า GROUP_ID อัตโนมัติ เพราะ Apps Script อ่านลายเซ็นของ LINE ไม่ได้ จึงยืนยันไม่ได้ว่าคำขอมาจาก LINE จริง
 */
function lineSetupEvents_(events) {
  if (prop_('LINE_SETUP') !== '1') return { ok: true };
  if (!events.length) log_('line', 'line_ping', 'LINE ติดต่อเข้ามาได้ (เช่น กด Verify)');   // ยืนยันว่า Webhook ชี้มาถึงระบบนี้
  events.slice(0, 20).forEach((ev) => {
    const src = (ev && ev.source) || {};
    const id = src.groupId || src.roomId;
    if (id) {
      const text = ev.message && ev.message.type === 'text' ? String(ev.message.text).slice(0, 40) : '(' + (ev.type || '?') + ')';
      log_('line', 'line_group_seen', 'groupId=' + id + ' | ข้อความ: ' + text);
    } else {
      log_('line', 'line_event', 'type=' + (ev && ev.type) + ' source=' + (src.type || '?'));   // ไม่ได้มาจากกลุ่ม (เช่น แชตส่วนตัวกับบอท)
    }
  });
  return { ok: true };
}

/* ---------- ยืนยันตัวตน ---------- */
function apiLogin(code, pin) {
  code = String(code || '').trim();
  pin = String(pin || '').trim();
  if (!/^\d{3,10}$/.test(code) || !/^\d{4,8}$/.test(pin)) throw new Error('กรอกรหัสพนักงานและรหัสผ่านให้ครบ');
  const cache = CacheService.getScriptCache(), key = 'f_' + code;
  const fails = +cache.get(key) || 0;
  if (fails >= 10) throw new Error('ใส่รหัสผิดหลายครั้ง รอ 10 นาทีแล้วลองใหม่');
  const st = staffByCode_(code);
  let ok = false, must = false;
  if (st) {
    if (String(st.pin_hash || '')) ok = hashPin_(pin, String(st.salt)) === String(st.pin_hash);
    else if (String(st.temp_pin || '') !== '' && String(st.temp_pin) === pin) { ok = true; must = true; }
  }
  if (!ok) {
    cache.put(key, String(fails + 1), 600);
    log_(code, 'login_fail', '');
    throw new Error('รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(key);
  const role = roleOf_(st, code), token = Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify({ code: code, role: role, must: must }), 21600);
  log_(code, 'login', must ? 'temp pin' : '');
  return { token: token, must: must, me: { code: code, role: role } };
}

/** พนักงานตั้งรหัสผ่านใหม่ (6-8 หลัก) ใช้ครั้งแรกหลังเข้าด้วยรหัสชั่วคราว */
function apiChangePin(auth, newPin) {
  const ses = session_(auth);
  newPin = String(newPin || '').trim();
  if (!/^\d{6,8}$/.test(newPin)) throw new Error('รหัสผ่านใหม่ต้องเป็นตัวเลข 6-8 หลัก');
  return withLock_(() => {
    const st = staffByCode_(ses.code);
    if (!st) throw new Error('ไม่พบบัญชีพนักงาน');
    const salt = Utilities.getUuid();
    update_('Staff', st._row, { pin_hash: hashPin_(newPin, salt), salt: salt, temp_pin: '', updated_at: new Date() });
    ses.must = false;
    CacheService.getScriptCache().put('s_' + auth.token, JSON.stringify(ses), 21600);
    log_(ses.code, 'change_pin', '');
    return { ok: true };
  });
}

function session_(auth) {
  if (!auth || auth.type !== 'app' || !auth.token) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
  const v = CacheService.getScriptCache().get('s_' + auth.token);
  let ses = null;
  try { ses = v ? JSON.parse(v) : null; } catch (e) { ses = null; }
  // session ต้องมีรหัสพนักงาน (ตัดทิ้ง session แบบเก่าจากรหัสแผนกรวมที่ยังค้างในแคช)
  if (!ses || typeof ses !== 'object' || !ses.code) throw new Error('หมดเวลา กรุณาเข้าสู่ระบบใหม่');
  return ses;
}

/** ออกจากระบบ: ลบ session ฝั่งเซิร์ฟเวอร์ */
function apiLogout(auth) {
  if (auth && auth.type === 'app' && auth.token) CacheService.getScriptCache().remove('s_' + auth.token);
  return { ok: true };
}

function who_(auth) {
  if (!auth) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
  if (auth.type === 'app') {
    const ses = session_(auth);
    if (ses.must) throw new Error('ต้องตั้งรหัสผ่านใหม่ก่อนใช้งาน');
    return { name: ses.code, code: ses.code, role: ses.role, via: 'แอป' };
  }
  if (auth.type === 'line') return lineWho_(String(auth.accessToken || ''));
  throw new Error('ไม่รู้จักวิธีเข้าสู่ระบบ');
}

function lineWho_(token) {
  if (!token) throw new Error('ไม่มี LINE access token');
  const cache = CacheService.getScriptCache();
  const key = 'l_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const v = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  const vj = JSON.parse(v.getContentText() || '{}');
  if (v.getResponseCode() !== 200 || String(vj.client_id) !== prop_('LOGIN_CHANNEL_ID') || !(vj.expires_in > 0)) {
    throw new Error('ยืนยันตัวตนผ่าน LINE ไม่สำเร็จ');
  }
  const r = UrlFetchApp.fetch('https://api.line.me/v2/profile', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('อ่านโปรไฟล์ LINE ไม่สำเร็จ');
  const pf = JSON.parse(r.getContentText());
  const who = { name: pf.displayName, code: pf.displayName, role: 'Staff', via: 'LINE', userId: pf.userId };
  cache.put(key, JSON.stringify(who), 300);
  return who;
}

/* ---------- API สำหรับหน้าเว็บ ---------- */
const dayStr_ = (ms) => Utilities.formatDate(new Date(ms), 'Asia/Bangkok', 'yyyyMMdd');

/**
 * รายการที่ส่งให้หน้าแอป:
 * - กำลังติดตาม = ยังไม่พิมพ์ผล และลงทะเบียนไม่เกิน STALE_HOURS ชั่วโมง (ครอบคลุมเวรดึกข้ามเที่ยงคืน)
 * - เสร็จสิ้นของวัน `day` (ค่าเริ่มต้นวันนี้) = พิมพ์ผลในวันนั้น หรือค้างเกิน STALE_HOURS ที่ลงทะเบียนในวันนั้น
 */
function snapshot_(day) {
  const now = Date.now(), today = dayStr_(now);
  day = /^\d{8}$/.test(String(day || '')) ? String(day) : today;
  const staleMs = STALE_HOURS * 3600e3;
  const byP = {};
  readAll_('FollowUps').forEach((f) => {
    (byP[f.patient_id] = byP[f.patient_id] || []).push({ n: +f.n, at: ms_(f.at), by: f.by, via: f.via, ext: +f.ext_min });
  });
  const keep = (r) => {
    const reg = ms_(r.registered_at), pr = ms_(r.printed_at);
    if (!pr && now - reg < staleMs) return true;
    return dayStr_(pr || reg) === day;
  };
  const list = readAll_('Patients')
    .filter((r) => r.id && keep(r))
    .map((r) => ({
      id: r.id, queue: +r.queue, hn: String(r.hn), reg: ms_(r.registered_at), wait: +r.wait_min,
      next: ms_(r.next_alert_at), alerts: +r.alert_count || 0, lastAlert: ms_(r.last_alert_at),
      regBy: r.registered_by, pr: ms_(r.printed_at), prBy: r.printed_by,
      stale: !ms_(r.printed_at) && now - ms_(r.registered_at) >= staleMs,
      fols: (byP[r.id] || []).sort((a, b) => a.n - b.n),
    }))
    .sort((a, b) => (a.reg + a.wait * 60000) - (b.reg + b.wait * 60000));
  return { now: now, today: today, day: day, patients: list };
}

function apiList(auth, day) {
  const w = who_(auth);
  const s = snapshot_(day);
  s.me = { code: w.code || w.name, role: w.role || 'Staff' };
  return s;
}

/** HN รูปแบบ 69-4948: ถ้ามีแต่ตัวเลข 5-10 หลัก ใส่ขีดหลังเลข 2 หลักแรก */
function normHn_(v) {
  const t = String(v == null ? '' : v).trim();
  if (/^\d{2}-\d+$/.test(t)) return t;
  const d = t.replace(/[\s-]/g, '');
  return /^\d{5,10}$/.test(d) ? d.slice(0, 2) + '-' + d.slice(2) : t;
}

function apiAdd(auth, hn, waitMin) {
  const w = who_(auth);
  hn = normHn_(hn);
  waitMin = Math.round(+waitMin);
  if (!/^[0-9A-Za-z\-\/]{3,20}$/.test(hn)) throw new Error('รูปแบบ HN ไม่ถูกต้อง');
  if (!(waitMin >= 5 && waitMin <= 720)) throw new Error('เวลาแจ้งเตือนต้องอยู่ระหว่าง 5-720 นาที');
  return withLock_(() => {
    const rows = readAll_('Patients');
    if (rows.some((r) => normHn_(r.hn) === hn && !r.printed_at)) throw new Error('HN ' + hn + ' อยู่ในคิวแล้ว');
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
    const queue = rows.filter((r) => Utilities.formatDate(new Date(ms_(r.registered_at)), 'Asia/Bangkok', 'yyyyMMdd') === today)
      .reduce((m, r) => Math.max(m, +r.queue || 0), 0) + 1;
    const now = new Date();
    append_('Patients', {
      id: Utilities.getUuid(), queue: queue, hn: hn, registered_at: now, wait_min: waitMin,
      next_alert_at: new Date(now.getTime() + waitMin * 60000), alert_count: 0, registered_by: w.name,
    });
    log_(w.name, 'register', 'HN ' + hn + ' wait ' + waitMin + ' via ' + w.via);
    return snapshot_();
  });
}

function findActive_(id) {
  const p = readAll_('Patients').filter((r) => r.id === id)[0];
  if (!p) throw new Error('ไม่พบรายการ');
  if (p.printed_at) throw new Error('รายการนี้พิมพ์ผลแล้ว');
  return p;
}

function apiFollow(auth, id) {
  const w = who_(auth);
  return withLock_(() => {
    const p = findActive_(id);
    const n = readAll_('FollowUps').filter((f) => f.patient_id === id).length + 1;
    const now = new Date();
    append_('FollowUps', { id: Utilities.getUuid(), patient_id: id, n: n, at: now, by: w.name, via: w.via, ext_min: REPEAT_MIN });
    update_('Patients', p._row, { next_alert_at: new Date(now.getTime() + REPEAT_MIN * 60000) });
    log_(w.name, 'follow', 'HN ' + p.hn + ' ครั้งที่ ' + n + ' via ' + w.via);
    return snapshot_();
  });
}

function apiPrint(auth, id) {
  const w = who_(auth);
  return withLock_(() => {
    const p = findActive_(id);
    // TODO: เมื่อมีรายการ Lab ให้ตรวจว่าผลออกครบก่อนจึงอนุญาตให้พิมพ์
    update_('Patients', p._row, { printed_at: new Date(), printed_by: w.name });
    log_(w.name, 'print', 'HN ' + p.hn);
    return snapshot_();
  });
}

/* ---------- แจ้งเตือนทาง LINE (ตัวตั้งเวลาเรียกทุก 1 นาที) ---------- */
function checkAlerts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const now = Date.now();
    const due = readAll_('Patients').filter((r) => r.id && !r.printed_at && now - ms_(r.registered_at) < STALE_HOURS * 3600e3 && ms_(r.next_alert_at) && ms_(r.next_alert_at) <= now);
    if (!due.length) return;
    const items = due.map((r) => ({
      hn: String(r.hn), count: +r.alert_count || 0,
      over: Math.max(0, Math.round((now - ms_(r.registered_at) - (+r.wait_min) * 60000) / 60000)),
    }));
    // 1 เคสต่อ 1 กล่องข้อความ (LINE ส่งได้สูงสุด 5 กล่องต่อคำขอ) ส่งสำเร็จเฉพาะชุดไหนจึงอัปเดตเฉพาะชุดนั้น
    const sent = [];
    const lineOn = !!(prop_('LINE_TOKEN') && prop_('GROUP_ID'));
    // ถ้ายังไม่ได้ตั้งค่า LINE ให้นับรอบเตือนต่อไปตามปกติ (แท็บและเวลาในแอปยังทำงาน) โดยไม่ส่งข้อความ
    if (!lineOn) items.forEach((_, j) => sent.push(j));
    for (let i = 0; lineOn && i < items.length; i += 5) {
      try {
        push_(items.slice(i, i + 5).map(alertFlex_));
        for (let j = i; j < Math.min(i + 5, items.length); j++) sent.push(j);
      } catch (err) {
        log_('system', 'push_error', String(err));
      }
    }
    if (!sent.length) return; // ลองใหม่ในนาทีถัดไป
    const at = new Date(now), next = new Date(now + REPEAT_MIN * 60000);
    sent.map((j) => due[j]).forEach((r) => {
      update_('Patients', r._row, { alert_count: (+r.alert_count || 0) + 1, last_alert_at: at, next_alert_at: next });
      log_('system', 'alert', 'HN ' + r.hn + ' ครั้งที่ ' + ((+r.alert_count || 0) + 1));
    });
  } finally { lock.releaseLock(); }
}

function push_(messages) {
  const token = prop_('LINE_TOKEN'), to = prop_('GROUP_ID');
  if (!token || !to) throw new Error('ยังไม่ได้ตั้ง LINE_TOKEN / GROUP_ID');
  const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: messages }), muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('LINE ' + r.getResponseCode() + ' ' + r.getContentText());
}

/**
 * ลิงก์ที่ปุ่มในข้อความ LINE เปิด: ถ้าตั้ง LIFF_ID ใช้ลิงก์ LIFF (เปิดในแอป LINE) ถ้าไม่ตั้ง เปิดหน้าแอปบน GitHub Pages ตรง ๆ
 * (APP_URL ค่าเริ่มต้นคือหน้าแอปของ repo นี้) พนักงานเข้าสู่ระบบด้วยรหัสพนักงานอยู่แล้ว จึงไม่ต้องใช้ LIFF
 */
function liffUrl_(query) {
  const liff = prop_('LIFF_ID');
  if (liff) return 'https://liff.line.me/' + liff + '?' + (query ? query + '&' : '') + 'api=' + encodeURIComponent(ScriptApp.getService().getUrl());
  const base = (prop_('APP_URL') || 'https://kung54150.github.io/LabAlert/').trim();
  return base + (query ? (base.indexOf('?') < 0 ? '?' : '&') + query : '');
}

function alertFlex_(i) {
  const text = i.count === 0 ? 'ครบเวลาตามผล' : 'แจ้งซ้ำครั้งที่ ' + (i.count + 1) + ' · เลยกำหนด ' + i.over + ' นาที';
  return {
    type: 'flex', altText: 'HN ' + i.hn + ' ' + text,
    contents: {
      type: 'bubble',
      header: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: 'LabAlert · ' + text, weight: 'bold', size: 'sm', color: i.count ? '#B3261E' : '#946000' }] },
      body: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: 'HN ' + i.hn, weight: 'bold', size: 'xl' }] },
      footer: { type: 'box', layout: 'vertical', contents: [
        { type: 'button', style: 'primary', height: 'sm', action: { type: 'uri', label: 'ตามผล', uri: liffUrl_('mode=follow&hn=' + encodeURIComponent(i.hn)) } },
        { type: 'button', style: 'secondary', height: 'sm', margin: 'sm', action: { type: 'uri', label: 'เปิดรายการทั้งหมด', uri: liffUrl_('') } },
      ] },
    },
  };
}

/** ส่งข้อความเมนู (สแกน / ดูคิว) เข้ากลุ่ม แล้วให้ปักหมุดเป็นประกาศ: รันจาก editor ครั้งเดียว */
function postMenu() {
  assertOwner_();
  push_([{
    type: 'flex', altText: 'เมนู LabAlert',
    contents: {
      type: 'bubble',
      body: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: 'LabAlert อายุรกรรม', weight: 'bold', size: 'lg' }] },
      footer: { type: 'box', layout: 'vertical', contents: [
        { type: 'button', style: 'primary', margin: 'sm', action: { type: 'uri', label: 'สแกน Barcode / QR', uri: liffUrl_('mode=scan') } },
        { type: 'button', style: 'secondary', margin: 'sm', action: { type: 'uri', label: 'ดูคิวผู้ป่วย', uri: liffUrl_('') } },
      ] },
    },
  }]);
}

/** ส่งตัวอย่างข้อความแจ้งเตือน 2 กล่อง (ครั้งแรก และแจ้งซ้ำ) เข้ากลุ่ม ด้วย HN ทดสอบ 69-123456 ไม่สร้างรายการผู้ป่วยจริง ใช้ดูหน้าตาข้อความและปุ่มได้ทันทีโดยไม่ต้องรอเวลา */
function testAlert() {
  assertOwner_();
  push_([alertFlex_({ hn: '69-123456', count: 0, over: 0 }), alertFlex_({ hn: '69-123456', count: 1, over: 10 })]);
  return 'ส่งตัวอย่างข้อความแจ้งเตือน 2 กล่องเข้ากลุ่มแล้ว';
}

/**
 * ดูโควตาข้อความ LINE ของเดือนนี้: โควตาทั้งหมด จำนวนที่ใช้ไปแล้ว และที่เหลือ (ถามจาก LINE ตรง ๆ ผ่าน token)
 * ผลลัพธ์อยู่ในบันทึกการดำเนินการ (Execution log) และชีต Log แถว line_quota ไม่แสดงหรือเก็บ token
 */
function lineQuota() {
  assertOwner_();
  const token = prop_('LINE_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ตั้ง LINE_TOKEN');
  const get = (path) => {
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('LINE ' + r.getResponseCode() + ' ' + r.getContentText());
    return JSON.parse(r.getContentText());
  };
  const q = get('quota'), c = get('quota/consumption');
  const used = +c.totalUsage || 0;
  const msg = q.type === 'limited'
    ? 'โควตาเดือนนี้ ' + q.value + ' ข้อความ ใช้ไปแล้ว ' + used + ' เหลือ ' + Math.max(0, q.value - used)
    : 'ไม่จำกัดโควตา ใช้ไปแล้ว ' + used + ' ข้อความ';
  Logger.log(msg);
  log_('system', 'line_quota', msg);
  return msg;
}

/** ทดสอบส่งข้อความเข้ากลุ่ม */
function testPush() {
  assertOwner_();
  push_([{ type: 'text', text: 'ทดสอบ LabAlert: ส่งข้อความเข้ากลุ่มสำเร็จ' }]);
}
