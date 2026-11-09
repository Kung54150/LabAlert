/**
 * LabAlert - ติดตามผล Lab ผู้ป่วยอายุรกรรม
 * ฐานข้อมูล: Google Sheets (ชีต Patients, FollowUps, Log)
 * ความลับทั้งหมดเก็บใน Script Properties (ไม่ฝังในโค้ด):
 *   LINE_TOKEN        Channel access token ของบอท (Messaging API)
 *   GROUP_ID          ID กลุ่ม LINE ที่รับแจ้งเตือน
 *   DEPT_PIN          รหัสแผนก (ตัวเลข) สำหรับเข้าใช้ผ่านแอป/เว็บ
 *   LIFF_ID           LIFF ID ของแอปนี้
 *   LOGIN_CHANNEL_ID  Channel ID ของ LINE Login channel ที่ LIFF สังกัดอยู่
 *   SHEET_ID          (ไม่บังคับ) ถ้าไม่ได้ผูกสคริปต์กับชีต
 */
const REPEAT_MIN = 10;          // แจ้งซ้ำเมื่อไม่มีใครตอบ
const KEEP_DONE_HOURS = 24;     // แสดงรายการในแท็บเสร็จสิ้นย้อนหลังกี่ชั่วโมง
const SCHEMA = {
  Patients: ['id', 'queue', 'hn', 'registered_at', 'wait_min', 'next_alert_at', 'alert_count',
             'last_alert_at', 'registered_by', 'printed_at', 'printed_by'],
  FollowUps: ['id', 'patient_id', 'n', 'at', 'by', 'via', 'ext_min'],
  Log: ['at', 'who', 'action', 'detail'],
};

/* ---------- helpers ---------- */
const prop_ = (k) => PropertiesService.getScriptProperties().getProperty(k) || '';
const ms_ = (v) => (v instanceof Date ? v.getTime() : v ? new Date(v).getTime() : 0);

function ss_() {
  const id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) { return ss_().getSheetByName(name); }

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
  ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'checkAlerts')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkAlerts').timeBased().everyMinutes(1).create();
  return 'ติดตั้งเสร็จ: สร้างชีตและตัวตั้งเวลาทุก 1 นาทีแล้ว';
}

/* ---------- Web App ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cfg = JSON.stringify({ liffId: prop_('LIFF_ID'), mode: String(p.mode || ''), hn: String(p.hn || '') })
    .replace(/</g, '\\u003c');
  const t = HtmlService.createTemplateFromFile('Index');
  t.cfg = cfg;
  return t.evaluate().setTitle('LabAlert อายุรกรรม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- ยืนยันตัวตน ---------- */
function apiLogin(pin) {
  const cache = CacheService.getScriptCache();
  const fails = +cache.get('fails') || 0;
  if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง รอ 10 นาทีแล้วลองใหม่');
  const real = prop_('DEPT_PIN');
  if (!real) throw new Error('ยังไม่ได้ตั้งรหัสแผนก (DEPT_PIN)');
  if (String(pin) !== real) {
    cache.put('fails', String(fails + 1), 600);
    throw new Error('รหัสไม่ถูกต้อง');
  }
  const token = Utilities.getUuid();
  cache.put('s_' + token, '1', 21600);
  return { token: token };
}

function who_(auth) {
  if (!auth) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
  if (auth.type === 'app') {
    if (!auth.token || !CacheService.getScriptCache().get('s_' + auth.token)) throw new Error('หมดเวลา กรุณาเข้าสู่ระบบใหม่');
    const nick = String(auth.nick || '').trim().slice(0, 30);
    return { name: nick ? nick + ' (แผนก)' : 'แผนก (แอป)', via: 'แอป' };
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
  const who = { name: pf.displayName, via: 'LINE', userId: pf.userId };
  cache.put(key, JSON.stringify(who), 300);
  return who;
}

/* ---------- API สำหรับหน้าเว็บ ---------- */
function snapshot_() {
  const now = Date.now();
  const byP = {};
  readAll_('FollowUps').forEach((f) => {
    (byP[f.patient_id] = byP[f.patient_id] || []).push({ n: +f.n, at: ms_(f.at), by: f.by, via: f.via, ext: +f.ext_min });
  });
  const list = readAll_('Patients')
    .filter((r) => r.id && (!r.printed_at || now - ms_(r.printed_at) < KEEP_DONE_HOURS * 3600e3))
    .map((r) => ({
      id: r.id, queue: +r.queue, hn: String(r.hn), reg: ms_(r.registered_at), wait: +r.wait_min,
      next: ms_(r.next_alert_at), alerts: +r.alert_count || 0, lastAlert: ms_(r.last_alert_at),
      regBy: r.registered_by, pr: ms_(r.printed_at), prBy: r.printed_by,
      fols: (byP[r.id] || []).sort((a, b) => a.n - b.n),
    }))
    // เรียงตามเวลาที่ผลควรออก (ลงทะเบียน + เวลารอ) คนที่ครบกำหนดก่อนอยู่บนสุด หน้าเว็บแยกแท็บเสร็จสิ้นเอง
    .sort((a, b) => (a.reg + a.wait * 60000) - (b.reg + b.wait * 60000));
  return { now: now, patients: list };
}

function apiList(auth) { who_(auth); return snapshot_(); }

function apiAdd(auth, hn, waitMin) {
  const w = who_(auth);
  hn = String(hn || '').trim();
  waitMin = Math.round(+waitMin);
  if (!/^[0-9A-Za-z\-\/]{3,20}$/.test(hn)) throw new Error('รูปแบบ HN ไม่ถูกต้อง');
  if (!(waitMin >= 5 && waitMin <= 720)) throw new Error('เวลาแจ้งเตือนต้องอยู่ระหว่าง 5-720 นาที');
  return withLock_(() => {
    const rows = readAll_('Patients');
    if (rows.some((r) => String(r.hn) === hn && !r.printed_at)) throw new Error('HN ' + hn + ' อยู่ในคิวแล้ว');
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
    const queue = rows.filter((r) => Utilities.formatDate(new Date(ms_(r.registered_at)), 'Asia/Bangkok', 'yyyyMMdd') === today)
      .reduce((m, r) => Math.max(m, +r.queue || 0), 0) + 1;
    const now = new Date();
    append_('Patients', {
      id: Utilities.getUuid(), queue: queue, hn: hn, registered_at: now, wait_min: waitMin,
      next_alert_at: new Date(now.getTime() + waitMin * 60000), alert_count: 0, registered_by: w.name + ' (' + w.via + ')',
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
    update_('Patients', p._row, { printed_at: new Date(), printed_by: w.name + ' (' + w.via + ')' });
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
    const due = readAll_('Patients').filter((r) => r.id && !r.printed_at && ms_(r.next_alert_at) && ms_(r.next_alert_at) <= now);
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

function liffUrl_(query) { return 'https://liff.line.me/' + prop_('LIFF_ID') + (query ? '?' + query : ''); }

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

/** ทดสอบส่งข้อความเข้ากลุ่ม */
function testPush() {
  assertOwner_();
  push_([{ type: 'text', text: 'ทดสอบ LabAlert: ส่งข้อความเข้ากลุ่มสำเร็จ' }]);
}
