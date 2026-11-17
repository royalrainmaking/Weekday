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
const ST_HEAD = ['รหัส', 'ชื่อ-สกุล', 'ชื่อเล่น', 'ตำแหน่ง', 'กลุ่ม', 'ลำดับในกลุ่ม', 'วันลาพักผ่อนยกมา', 'สิทธิ์ลาพักผ่อนปีนี้', 'รวมวันลาพักผ่อน (ไม่เกินเพดาน)', 'เพดานวันลาสะสม'];
const HO_HEAD = ['วันที่', 'ชื่อวันหยุด'];

const TYPES = {
  sick: ['ลาป่วย', 'ป'], biz: ['ลากิจ', 'ก'], vac: ['ลาพักผ่อน', 'พ'], study: ['ลาไปศึกษา', 'ศ'],
  trip: ['ไปราชการ', 'รช'], train: ['อบรม', 'อ'], wfh: ['WFH', 'w'], late: ['มาสาย', 'ส'],
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
    staff[r[0]] = { name: r[1], nick: r[2], position: r[3], group: r[4] || GROUPS[0], order: num_(r[5]) || 999, carry: num_(r[6]), quota: num_(r[7]), cap: num_(r[9]) || 20 };
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
