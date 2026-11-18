/**
 * ปฏิทินงานทีม — Google Apps Script API (หน้าเว็บอยู่บน GitHub Pages)
 * ใช้ Google Sheet ที่ผูกกับสคริปต์นี้เป็นฐานข้อมูล
 *   GET  ?action=getAll                      -> {ok:true, data:{staff, events, holidays, url, at}}
 *   POST {"action":"save","col","id","data"} -> {ok:true, data:...}
 *   POST {"action":"delete","col","id"}       -> {ok:true, data:...}
 *   แท็บ "รายการ"  : ลา / ไปราชการ / ประชุม / งานที่ต้องส่ง / นัดหมาย
 *   แท็บ "บุคลากร" : รายชื่อ
 *   แท็บ "วันหยุด" : วันหยุดราชการ
 */

const SH_EVENTS = 'รายการ';
const SH_STAFF = 'บุคลากร';
const SH_HOL = 'วันหยุด';

const EV_HEAD = ['รหัสรายการ', 'ประเภท', 'รหัสในบัญชีลงเวลา', 'วันที่เริ่ม', 'วันที่สิ้นสุด', 'เวลาเริ่ม', 'เวลาสิ้นสุด',
  'เรื่อง', 'สถานที่', 'ผู้เกี่ยวข้อง', 'สถานะงานส่ง', 'หมายเหตุ', 'บันทึกล่าสุด', 'รหัสผู้เกี่ยวข้อง', 'ชนิด', 'ผู้บันทึก'];
const ST_HEAD = ['รหัส', 'ชื่อ-สกุล', 'ชื่อเล่น', 'ตำแหน่ง', 'กลุ่ม', 'ลำดับในกลุ่ม', 'วันลาพักผ่อนยกมา', 'สิทธิ์ลาพักผ่อนปีนี้', 'รวมวันลาพักผ่อน (ไม่เกินเพดาน)', 'เพดานวันลาสะสม', 'วันที่เข้ากลุ่ม', 'วันที่ออกจากกลุ่ม', 'ชื่อ LINE (สำหรับแท็ก)'];
const HO_HEAD = ['วันที่', 'ชื่อวันหยุด'];

const TYPES = {
  sick: ['ลาป่วย', 'ป'], biz: ['ลากิจ', 'ก'], vac: ['ลาพักผ่อน', 'พ'], study: ['ลาไปศึกษา', 'ศ'],
  trip: ['ไปราชการ', 'รช'], train: ['อบรม', 'อ'], wfh: ['WFH', 'WFH'], late: ['มาสาย', 'ส'],
  meet: ['ประชุม', 'ปช'], task: ['งานที่ต้องส่ง', ''], other: ['นัดหมาย/อื่นๆ', '']
};
const GROUPS = ['ข้าราชการ', 'จ้างเหมาบริการ', 'พนักงานราชการ'];
const TZ = 'Asia/Bangkok';

/* ---------- web API ---------- */
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action !== 'getAll') throw new Error('ไม่รู้จักคำสั่ง');
    return json_({ ok: true, data: getAll() });
  } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (Array.isArray(body.events)) { captureGroupId_(body); return ContentService.createTextOutput('OK'); }  // จาก LINE (ใช้ตอนหารหัสกลุ่มเท่านั้น)
    if (body.action === 'save') return json_({ ok: true, data: saveItem(body.col, body.id, body.data) });
    if (body.action === 'delete') return json_({ ok: true, data: deleteItem(body.col, body.id) });
    if (body.action === 'getAll') return json_({ ok: true, data: getAll() });
    throw new Error('ไม่รู้จักคำสั่ง');
  } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

/** เรียกครั้งแรกครั้งเดียว (Run > setup) เพื่อสร้างแท็บและหัวตาราง */
function setup() {
  sheet_(SH_EVENTS, EV_HEAD);
  sheet_(SH_STAFF, ST_HEAD);
  sheet_(SH_HOL, HO_HEAD);
  return 'ok';
}

/* ---------- helpers ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, head) {
  let sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);
  const cur = sh.getRange(1, 1, 1, head.length).getDisplayValues()[0];
  if (cur.join('|') !== head.join('|')) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#ffe0e9').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(sh, n) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getDisplayValues();
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
// กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้กลายเป็นสูตร
function txt_(v, max) {
  v = v == null ? '' : String(v).trim();
  if (max) v = v.slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function isoDate_(v) { v = String(v || '').trim(); if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('รูปแบบวันที่ไม่ถูกต้อง'); return v; }
function hhmm_(v) { v = String(v || '').trim(); return /^\d{1,2}:\d{2}$/.test(v) ? v : ''; }
function idOk_(v) { v = String(v || '').trim(); if (!/^[A-Za-z0-9_-]{1,40}$/.test(v)) throw new Error('รหัสไม่ถูกต้อง'); return v; }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function num_(v) { const n = Number(String(v).replace(/,/g, '')); return isFinite(n) ? n : 0; }

/* ---------- read ---------- */
function getAll() {
  const staff = {}, events = {}, holidays = {};
  rows_(sheet_(SH_STAFF, ST_HEAD), ST_HEAD.length).forEach(r => {
    if (!r[0]) return;
    staff[r[0]] = { name: r[1], nick: r[2], position: r[3], group: r[4] || GROUPS[0], order: num_(r[5]) || 999, carry: num_(r[6]), quota: num_(r[7]), cap: num_(r[9]) || 20, joined: r[10] || '', left: r[11] || '', lineName: r[12] || '' };
  });
  rows_(sheet_(SH_EVENTS, EV_HEAD), EV_HEAD.length).forEach(r => {
    if (!r[0]) return;
    const type = TYPES[r[14]] ? r[14] : (Object.keys(TYPES).find(k => TYPES[k][0] === r[1]) || 'other');
    events[r[0]] = {
      type, start: r[3], end: r[4] || r[3], time: r[5], endTime: r[6], title: r[7], place: r[8],
      people: String(r[13] || '').split(',').map(s => s.trim()).filter(Boolean),
      done: r[10] === 'ส่งแล้ว', note: r[11], updatedAt: r[12], by: r[15]
    };
  });
  rows_(sheet_(SH_HOL, HO_HEAD), HO_HEAD.length).forEach(r => { if (r[0]) holidays[r[0]] = { name: r[1] }; });
  return { staff, events, holidays, url: ss_().getUrl(), at: now_() };
}

/* ---------- write ---------- */
function saveItem(col, id, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    data = data || {};
    if (col === 'events') saveEvent_(idOk_(id), data);
    else if (col === 'staff') saveStaff_(idOk_(id), data);
    else if (col === 'holidays') saveHoliday_(isoDate_(id), data);
    else throw new Error('ไม่รู้จักชนิดข้อมูล');
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return getAll();
}

function deleteItem(col, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const name = col === 'events' ? SH_EVENTS : col === 'staff' ? SH_STAFF : col === 'holidays' ? SH_HOL : null;
    if (!name) throw new Error('ไม่รู้จักชนิดข้อมูล');
    const sh = ss_().getSheetByName(name);
    const row = sh ? findRow_(sh, id) : -1;
    if (row > 0) sh.deleteRow(row);
    if (col === 'staff') refreshNames_();
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return getAll();
}

function saveEvent_(id, d) {
  const sh = sheet_(SH_EVENTS, EV_HEAD);
  const type = TYPES[d.type] ? d.type : 'other';
  const start = isoDate_(d.start);
  const end = d.end ? isoDate_(d.end) : start;
  if (end < start) throw new Error('วันที่สิ้นสุดต้องไม่ก่อนวันเริ่ม');
  const people = (Array.isArray(d.people) ? d.people : []).map(String).filter(p => /^[A-Za-z0-9_-]{1,40}$/.test(p)).slice(0, 100);
  if (!people.length) throw new Error('ต้องเลือกอย่างน้อย 1 คน');
  const names = staffNames_();
  const row = [id, TYPES[type][0], TYPES[type][1], start, end, hhmm_(d.time), hhmm_(d.endTime),
    txt_(d.title, 300), txt_(d.place, 200), txt_(people.map(p => names[p] || '(ลบแล้ว)').join(', ')),
    type === 'task' ? (d.done ? 'ส่งแล้ว' : 'ยังไม่ส่ง') : '', txt_(d.note, 1000), now_(), people.join(','), type, txt_(d.by, 80)];
  let r = findRow_(sh, id);
  if (r < 0) r = sh.getLastRow() + 1;
  const range = sh.getRange(r, 1, 1, row.length);
  range.setNumberFormat('@');
  range.setValues([row]);
}

function saveStaff_(id, d) {
  const sh = sheet_(SH_STAFF, ST_HEAD);
  const name = txt_(d.name, 120);
  if (!name) throw new Error('ต้องใส่ชื่อ');
  let r = findRow_(sh, id);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 5).setNumberFormat('@')
    .setValues([[id, name, txt_(d.nick, 60), txt_(d.position, 160), GROUPS.indexOf(d.group) >= 0 ? d.group : GROUPS[0]]]);
  const cap = [10, 15, 20, 30].indexOf(num_(d.cap)) >= 0 ? num_(d.cap) : 20;
  sh.getRange(r, 6, 1, 5).setNumberFormat('0.##')
    .setValues([[num_(d.order) || 999, num_(d.carry), num_(d.quota), '=MIN(G' + r + '+H' + r + ',J' + r + ')', cap]]);
  const optDate = v => (v && /^\d{4}-\d{2}-\d{2}$/.test(String(v))) ? String(v) : '';
  sh.getRange(r, 11, 1, 2).setNumberFormat('@').setValues([[optDate(d.joined), optDate(d.left)]]);
  refreshNames_();
}

function saveHoliday_(date, d) {
  const sh = sheet_(SH_HOL, HO_HEAD);
  let r = findRow_(sh, date);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 2).setNumberFormat('@').setValues([[date, txt_(d.name, 120) || 'วันหยุดราชการ']]);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).sort(1);
}

function staffNames_() {
  const m = {};
  rows_(sheet_(SH_STAFF, ST_HEAD), 2).forEach(r => { if (r[0]) m[r[0]] = r[1]; });
  return m;
}

// อัปเดตคอลัมน์ "ผู้เกี่ยวข้อง" ให้ตรงกับรายชื่อล่าสุด
function refreshNames_() {
  const sh = ss_().getSheetByName(SH_EVENTS);
  if (!sh || sh.getLastRow() < 2) return;
  const names = staffNames_();
  const n = sh.getLastRow() - 1;
  const ids = sh.getRange(2, 14, n, 1).getDisplayValues();
  const out = ids.map(r => [txt_(String(r[0] || '').split(',').filter(Boolean).map(p => names[p.trim()] || '(ลบแล้ว)').join(', '))]);
  sh.getRange(2, 10, n, 1).setNumberFormat('@').setValues(out);
}


/* =====================================================================
 *  แจ้งเตือนนัดหมายเข้ากลุ่ม LINE ทุกเช้า (ส่งอย่างเดียว ไม่ตอบแชท)
 *  ตั้งค่า
 *   1) ใส่รหัสกลุ่มที่ GROUP_ID ด้านล่าง (ขึ้นต้นด้วย C ตามด้วยตัวอักษร 32 ตัว)
 *   2) การตั้งค่าโปรเจกต์ (⚙️) > Script Properties > เพิ่ม  LINE_TOKEN = Channel access token (long-lived)
 *   3) รัน setupLineTrigger() หนึ่งครั้ง (อนุญาตสิทธิ์)
 *   ทดสอบ: รัน previewAgenda() (ไม่ส่ง) หรือ testSendNow() (ส่งจริง)
 * ===================================================================== */
const GROUP_ID = 'C0a8ade06f93c42ee40b1f870d2593f44';                 // ← ใส่รหัสกลุ่ม LINE ตรงนี้ เช่น 'C1234abcd...'
const LINE_SEND_HOUR = 7;            // ส่งตอนกี่โมง (ชั่วโมง)
const LINE_SEND_MINUTE = 30;         // นาที (Apps Script อาจคลาดเคลื่อน ±15 นาที)
const LINE_SKIP_OFFDAYS = true;      // ไม่ส่งเสาร์-อาทิตย์ และวันหยุดในแท็บ "วันหยุด"

const TH_DAY = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
/* ---------- LINE API ---------- */
function props_() { return PropertiesService.getScriptProperties(); }
function lineToken_() {
  const t = props_().getProperty('LINE_TOKEN');
  if (!t) throw new Error('ยังไม่ได้ใส่ LINE_TOKEN ใน Script Properties');
  return t;
}
function lineApi_(path, payload, method) {
  const res = UrlFetchApp.fetch('https://api.line.me' + path, {
    method: method || 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + lineToken_() },
    payload: payload ? JSON.stringify(payload) : undefined
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('LINE API ' + code + ': ' + res.getContentText());
  return res.getContentText();
}
function linePush_(to, msg) {
  const m = typeof msg === 'string' ? { type: 'text', text: msg.slice(0, 4900) } : msg;
  lineApi_('/v2/bot/message/push', { to, messages: [m] });
}
// ตรวจว่าบอทอยู่ในกลุ่มนี้จริง (กันข้อมูลปลอมที่ยิงเข้ามาที่ /exec)
function groupIsReal_(groupId) {
  if (!/^C[0-9a-f]{32}$/.test(String(groupId || ''))) return false;
  try { lineApi_('/v2/bot/group/' + groupId + '/summary', null, 'get'); return true; } catch (e) { return false; }
}

/* ---------- สร้างข้อความสรุปนัดของวัน ---------- */
function isoToday_(offsetDays) {
  const d = new Date(); d.setDate(d.getDate() + (offsetDays || 0));
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function weekday_(iso) { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); }
function thDateLong_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return 'วัน' + TH_DAY[weekday_(iso)] + 'ที่ ' + d + ' ' + TH_MON[m - 1] + ' ' + (y + 543);
}
function thTime_(t) { return t ? String(t).replace(':', '.') : ''; }
function shortName_(n) { return String(n || '').replace(/^(นาย|นางสาว|นาง|น\.ส\.|ว่าที่\s*ร\.ต\.)\s*/, '').split(/\s+/)[0]; }

// คืนค่าเป็นข้อความ LINE (textV2 + แท็กชื่อ) หรือ null ถ้าวันนั้นไม่มีนัด
function buildAgenda_(iso, data) {
  data = data || getAll();
  const users = lineUsers_();                       // ชื่อที่แสดงใน LINE -> userId
  const sub = {}; let n = 0;
  const esc = t => String(t || '').replace(/[{}]/g, c => c + c);   // textV2: { } ต้องเขียนซ้ำ
  const who1 = id => {
    const p = data.staff[id]; if (!p) return '?';
    const ln = String(p.lineName || '').trim();
    const uid = /^U[0-9a-f]{32}$/.test(ln) ? ln : users[normName_(ln)];
    if (uid && n < 20) { const k = 'm' + (n++); sub[k] = { type: 'mention', mentionee: { type: 'user', userId: uid } }; return '{' + k + '}'; }
    return esc(p.nick || shortName_(p.name));
  };
  const evs = Object.keys(data.events).map(id => Object.assign({ id }, data.events[id]))
    .filter(e => ['meet', 'task', 'other'].indexOf(e.type) >= 0 && e.start <= iso && (e.end || e.start) >= iso)
    .sort((a, b) => (a.time || '99').localeCompare(b.time || '99'));
  if (!evs.length) return null;
  const line = e => {
    const time = e.time ? thTime_(e.time) + (e.endTime ? '–' + thTime_(e.endTime) : '') + ' น. ' : '';
    const title = esc(e.title || (e.type === 'meet' ? 'ประชุม' : e.type === 'task' ? 'งานที่ต้องส่ง' : 'นัดหมาย'));
    const place = e.place ? ' @' + esc(e.place) : '';
    const who = (e.people || []).map(who1).join(' ');
    const done = e.type === 'task' && e.done ? ' ✅ ส่งแล้ว' : '';
    return '• ' + time + title + place + done + (who ? '\n   👤 ' + who : '');
  };
  const sec = (label, type) => { const l = evs.filter(e => e.type === type); return l.length ? '\n' + label + '\n' + l.map(line).join('\n') : ''; };
  const text = '📅 นัดหมาย' + thDateLong_(iso) + '\n' + sec('🗓️ ประชุม', 'meet') + sec('📌 งานที่ต้องส่ง', 'task') + sec('⭐ นัดหมาย', 'other');
  return n ? { type: 'textV2', text: text.slice(0, 4900), substitution: sub } : { type: 'text', text: text.replace(/\{\{/g, '{').replace(/\}\}/g, '}').slice(0, 4900) };
}

/* ---------- ส่งอัตโนมัติทุกเช้า (ทำงานจาก Trigger) ---------- */
function sendDailyAgenda() {
  const iso = isoToday_(0);
  const data = getAll();
  if (LINE_SKIP_OFFDAYS) {
    const wd = weekday_(iso);
    if (wd === 0 || wd === 6 || data.holidays[iso]) { console.log('วันหยุด ไม่ส่ง'); return 'skip-holiday'; }
  }
  const text = buildAgenda_(iso, data);
  if (!text) { console.log('วันนี้ไม่มีนัด ไม่ส่ง (ไม่เสียโควตา)'); return 'skip-empty'; }
  const gid = GROUP_ID || props_().getProperty('LINE_GROUP_ID');
  if (!gid) throw new Error('ยังไม่ได้ใส่ GROUP_ID (รหัสกลุ่ม LINE) ด้านบนของไฟล์');
  linePush_(gid, text);
  return 'sent';
}

/** รันครั้งเดียวเพื่อตั้งเวลาส่งทุกวัน (ลบ trigger เก่าของฟังก์ชันนี้ก่อน) */
function setupLineTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendDailyAgenda').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyAgenda').timeBased().everyDays(1).atHour(LINE_SEND_HOUR).nearMinute(LINE_SEND_MINUTE).inTimezone(TZ).create();
  const msg = 'ตั้งเวลาส่งทุกวันประมาณ ' + LINE_SEND_HOUR + '.' + ('0' + LINE_SEND_MINUTE).slice(-2) + ' น. แล้ว';
  console.log(msg); return msg;
}

/** ดูตัวอย่างข้อความของวันนี้ใน Log โดยไม่ส่ง (ไม่เสียโควตา) */
function previewAgenda() {
  const m = buildAgenda_(isoToday_(0));
  const t = m ? m.text + (m.substitution ? '\n\n[แท็ก] ' + JSON.stringify(m.substitution) : '') : '(วันนี้ไม่มีนัด — ระบบจะไม่ส่ง)';
  console.log(t); return t;
}

/** ทดสอบส่งของวันนี้เข้ากลุ่มทันที (ใช้โควตา ถ้าวันนี้มีนัด) */
function testSendNow() {
  const t = buildAgenda_(isoToday_(0));
  const gid = GROUP_ID || props_().getProperty('LINE_GROUP_ID');
  if (!gid) throw new Error('ยังไม่ได้ใส่ GROUP_ID');
  linePush_(gid, t || '✅ ทดสอบบอทปฏิทินงานทีม: เชื่อมต่อกลุ่มสำเร็จ (วันนี้ไม่มีนัด)');
  return 'sent';
}

/* ---------- หารหัสกลุ่ม (ทำครั้งเดียว ถ้ายังไม่รู้รหัส) ----------
 *  1) ใส่ลิงก์ /exec ของโปรเจกต์นี้เป็น Webhook URL ใน LINE Developers ชั่วคราว และเปิด Use webhook
 *  2) พิมพ์อะไรก็ได้ในกลุ่มที่มีบอท (บอทจะไม่ตอบ)
 *  3) รัน showGroupId() คัดลอกรหัสไปใส่ GROUP_ID แล้วปิด Use webhook ใน LINE ได้เลย */
const SH_LINE = 'LINE';
function captureGroupId_(body) {
  (body.events || []).forEach(ev => {
    try {
      const src = ev.source || {};
      const gid = src.type === 'group' ? src.groupId : '';
      if (!gid || !groupIsReal_(gid)) return;
      if (props_().getProperty('LINE_GROUP_ID') !== gid) { props_().setProperty('LINE_GROUP_ID', gid); console.log('รหัสกลุ่ม: ' + gid); }
      const uids = [];
      if (src.userId) uids.push(src.userId);
      if (ev.type === 'memberJoined' && ev.joined) (ev.joined.members || []).forEach(m => m.userId && uids.push(m.userId));
      uids.forEach(uid => saveLineUser_(gid, uid));
    } catch (err) { console.error(err); }
  });
}
// เก็บ userId + ชื่อที่แสดงใน LINE ลงแท็บ "LINE" (ใช้จับคู่กับคอลัมน์ "ชื่อ LINE" ในแท็บบุคลากร)
function saveLineUser_(gid, uid) {
  if (!/^U[0-9a-f]{32}$/.test(uid)) return;
  const prof = JSON.parse(lineApi_('/v2/bot/group/' + gid + '/member/' + uid, null, 'get'));
  const sh = sheet_(SH_LINE, ['userId', 'ชื่อที่แสดงใน LINE', 'อัปเดตล่าสุด']);
  let r = findRow_(sh, uid);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[uid, txt_(prof.displayName || '', 100), now_()]]);
}
// เทียบชื่อแบบไม่สนอีโมจิ เว้นวรรค วงเล็บ และตัวพิมพ์เล็ก/ใหญ่
function normName_(n) { return String(n || '').toLowerCase().replace(/[\uFE0E\uFE0F\u200D]/g, '').replace(/[^\p{L}\p{M}\p{N}]/gu, ''); }
function lineUsers_() {
  const sh = ss_().getSheetByName(SH_LINE);
  const m = {};
  if (sh) rows_(sh, 2).forEach(r => { if (r[0] && r[1]) m[normName_(r[1])] = r[0]; });
  return m;
}
function showGroupId() {
  const g = props_().getProperty('LINE_GROUP_ID') || '(ยังไม่มี — พิมพ์ข้อความในกลุ่มหลังตั้ง Webhook ก่อน)';
  console.log(g); return g;
}
