/**
 * ระบบบัญชีรายรับ-รายจ่าย โครงการธุรกิจน้ำมัน
 * Google Apps Script Web App  (Code.gs + index.html)
 *
 * ชีตที่ระบบสร้างให้อัตโนมัติ:
 *  - Transactions : รายการรับ/จ่ายทั้งหมด
 *  - Items        : รายการสินค้า/ค่าใช้จ่าย และราคาขายตั้งต้น
 */

const SHEET_TX = 'Transactions';
const SHEET_ITEMS = 'Items';
const TZ = 'Asia/Bangkok';
// ไอคอนแท็บเบราว์เซอร์: ใส่ลิงก์รูป .png/.ico (https) ที่อัปโหลดไว้ เช่น pump-icon-512.png
const FAVICON_URL = 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/26fd.png';
const FUEL_CAT = 'ซื้อน้ำมัน';
const TX_HEADERS = ['id', 'date', 'type', 'category', 'item', 'qty', 'amount',
  'costPerUnit', 'sellPrice', 'profitPerUnit', 'note', 'createdAt'];
const ITEM_HEADERS = ['name', 'category', 'sellPrice'];

/* ---------- Web app entry ---------- */
function doGet() {
  const out = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('บัญชีรายรับ-รายจ่าย | ธุรกิจน้ำมัน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

/* ---------- Spreadsheet helpers ---------- */
function getSS_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('ฐานข้อมูล - ธุรกิจน้ำมัน');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function sheets_() {
  const ss = getSS_();
  let tx = ss.getSheetByName(SHEET_TX);
  if (!tx) {
    tx = ss.insertSheet(SHEET_TX);
    tx.appendRow(TX_HEADERS);
    tx.getRange(1, 1, 1, TX_HEADERS.length).setFontWeight('bold').setBackground('#dbe7f3');
    tx.setFrozenRows(1);
    tx.getRange('B:B').setNumberFormat('@'); // เก็บวันที่เป็นข้อความ yyyy-MM-dd
  }
  let it = ss.getSheetByName(SHEET_ITEMS);
  if (!it) {
    it = ss.insertSheet(SHEET_ITEMS);
    it.appendRow(ITEM_HEADERS);
    it.getRange(1, 1, 1, ITEM_HEADERS.length).setFontWeight('bold').setBackground('#dbe7f3');
    it.setFrozenRows(1);
    const defaults = [
      ['แก๊สโซฮอล์ 95', 'ซื้อน้ำมัน', 43.9],
      ['ดีเซล', 'ซื้อน้ำมัน', 44.75],
      ['ตู้น้ำมัน', 'ลงทุนอุปกรณ์', 0],
      ['โครงน้ำมัน', 'ลงทุนอุปกรณ์', 0],
      ['เดินไฟ', 'ลงทุนอุปกรณ์', 0],
      ['ภาษี', 'ค่าใช้จ่ายอื่น', 0]
    ];
    it.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
  return { tx: tx, it: it };
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

const r2_ = n => Math.round(n * 100) / 100;

function readTx_() {
  const tx = sheets_().tx;
  const last = tx.getLastRow();
  if (last < 2) return [];
  return tx.getRange(2, 1, last - 1, TX_HEADERS.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      TX_HEADERS.forEach((h, i) => o[h] = r[i]);
      o.date = fmtDate_(o.date);
      o.createdAt = o.createdAt instanceof Date ? o.createdAt.toISOString() : String(o.createdAt || '');
      ['qty', 'amount', 'costPerUnit', 'sellPrice', 'profitPerUnit'].forEach(k => o[k] = Number(o[k]) || 0);
      o.id = String(o.id);
      return o;
    });
}

function readItems_() {
  const it = sheets_().it;
  const last = it.getLastRow();
  if (last < 2) return [];
  return it.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({ name: String(r[0]), category: String(r[1]), sellPrice: Number(r[2]) || 0 }));
}

/* ---------- API เรียกจากหน้าเว็บ ---------- */
function getInitData() {
  return { transactions: readTx_(), items: readItems_() };
}

function saveTransaction(t) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!t || !/^\d{4}-\d{2}-\d{2}$/.test(t.date)) throw new Error('กรุณาระบุวันที่');
    if (t.type !== 'income' && t.type !== 'expense') throw new Error('ประเภทรายการไม่ถูกต้อง');
    const amount = Number(t.amount) || 0;
    if (amount <= 0) throw new Error('กรุณาระบุจำนวนเงิน');

    const qty = Number(t.qty) || 0;
    const sell = Number(t.sellPrice) || 0;
    let cost = 0, profit = 0;
    if (t.type === 'expense' && t.category === FUEL_CAT && qty > 0) {
      cost = r2_(amount / qty);
      if (sell > 0) profit = r2_(sell - cost);
    }

    const tx = sheets_().tx;
    const last = tx.getLastRow();
    let row = -1;
    if (t.id && last > 1) {
      const idx = tx.getRange(2, 1, last - 1, 1).getValues().flat().map(String).indexOf(String(t.id));
      if (idx >= 0) row = idx + 2;
    }
    let id = t.id, createdAt;
    if (row < 0) {
      id = Utilities.getUuid();
      createdAt = new Date();
      row = last + 1;
    } else {
      createdAt = tx.getRange(row, 12).getValue();
    }
    tx.getRange(row, 2).setNumberFormat('@');
    tx.getRange(row, 1, 1, TX_HEADERS.length).setValues([[
      id, t.date, t.type, String(t.category || ''), String(t.item || '').trim(),
      qty, amount, cost, sell, profit, String(t.note || '').trim(), createdAt
    ]]);
    return readTx_();
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheets_().tx;
    const last = tx.getLastRow();
    if (last > 1) {
      const idx = tx.getRange(2, 1, last - 1, 1).getValues().flat().map(String).indexOf(String(id));
      if (idx >= 0) tx.deleteRow(idx + 2);
    }
    return readTx_();
  } finally {
    lock.releaseLock();
  }
}

function saveItems(list) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const it = sheets_().it;
    const clean = (list || [])
      .filter(x => x && String(x.name).trim() !== '')
      .map(x => [String(x.name).trim(), String(x.category || ''), Number(x.sellPrice) || 0]);
    const last = it.getLastRow();
    if (last > 1) it.getRange(2, 1, last - 1, 3).clearContent();
    if (clean.length) it.getRange(2, 1, clean.length, 3).setValues(clean);
    return readItems_();
  } finally {
    lock.releaseLock();
  }
}

/* ====================================================================
 *  ดึงข้อมูลจากระบบ QRcode (http://vps316.vpshispeed.net/qrnew/agent/)
 *  ตั้งค่าใน Script Properties (การตั้งค่าโปรเจ็กต์ → คุณสมบัติของสคริปต์):
 *    QR_USER = ชื่อผู้ใช้     QR_PASS = รหัสผ่าน
 *    QR_BASE = (ไม่บังคับ) http://vps316.vpshispeed.net/qrnew/agent/
 *  ข้อมูลเครื่องลูกสะสมไว้ในชีต QR_Daily (สรุปรายวัน) และ QR_Coin (รายการรับละเอียด)
 * ==================================================================== */
function qrCfg_() {
  const p = PropertiesService.getScriptProperties();
  const user = p.getProperty('QR_USER'), pass = p.getProperty('QR_PASS');
  if (!user || !pass) throw new Error('ยังไม่ได้ตั้งค่า QR_USER และ QR_PASS ใน Script Properties');
  let base = p.getProperty('QR_BASE') || 'http://vps316.vpshispeed.net/qrnew/agent/';
  if (!/\/$/.test(base)) base += '/';
  return { user: user, pass: pass, base: base };
}

const p2_ = n => ('0' + n).slice(-2);

function attr_(tag, name) {
  const m = tag.match(new RegExp('(?:^|\\s)' + name + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\'|([^\\s>]+))', 'i'));
  return m ? (m[1] || m[2] || m[3] || '') : '';
}

function resolveUrl_(from, loc) {
  if (/^https?:\/\//i.test(loc)) return loc;
  const m = from.match(/^(https?:\/\/[^\/]+)(\/[^?#]*)?/);
  const path = loc.charAt(0) === '/' ? loc : (m[2] || '/').replace(/[^\/]*$/, '') + loc;
  const q = path.search(/[?#]/);
  const pth = q < 0 ? path : path.slice(0, q), rest = q < 0 ? '' : path.slice(q);
  const out = [];
  pth.split('/').forEach(seg => {
    if (seg === '..') { if (out.length > 1) out.pop(); } else if (seg !== '.') out.push(seg);
  });
  return m[1] + out.join('/') + rest;
}

function mergeCookies_(jar, res) {
  const h = res.getAllHeaders();
  let sc = h['Set-Cookie'] || h['set-cookie'];
  if (!sc) return;
  if (!Array.isArray(sc)) sc = [sc];
  sc.forEach(c => {
    const kv = String(c).split(';')[0], i = kv.indexOf('=');
    if (i > 0) jar[kv.slice(0, i).trim()] = kv.slice(i + 1).trim();
  });
}

/** ส่งคำขอพร้อมเก็บ cookie และตามการ redirect เอง */
function qrReq_(url, jar, opt) {
  let o = opt || {};
  for (let i = 0; i < 6; i++) {
    const res = UrlFetchApp.fetch(url, {
      method: o.method || 'get', payload: o.payload, muteHttpExceptions: true, followRedirects: false,
      headers: { Cookie: Object.keys(jar).map(k => k + '=' + jar[k]).join('; '), 'User-Agent': 'Mozilla/5.0' }
    });
    mergeCookies_(jar, res);
    const code = res.getResponseCode();
    if (code >= 300 && code < 400) {
      const hd = res.getHeaders(), loc = hd['Location'] || hd['location'];
      if (!loc) return res;
      url = resolveUrl_(url, loc);
      o = {};
      continue;
    }
    return res;
  }
  throw new Error('redirect มากเกินไป');
}

/** อ่านฟอร์มในหน้า index.php แล้วล็อกอินให้อัตโนมัติ (หาชื่อช่องเอง) */
function qrLogin_(cfg, jar) {
  const loginUrl = cfg.base + 'index.php';
  const html = qrReq_(loginUrl, jar).getContentText();
  const form = html.match(/<form\b[\s\S]*?<\/form>/i);
  if (!form) throw new Error('ไม่พบฟอร์มล็อกอินใน index.php');
  const fields = {}, found = [];
  let userField = '', passField = '';
  const skip = ['hidden', 'submit', 'button', 'checkbox', 'radio', 'image', 'reset', 'file'];
  (form[0].match(/<input\b[^>]*>/gi) || []).forEach(tag => {
    const name = attr_(tag, 'name');
    const type = (attr_(tag, 'type') || 'text').toLowerCase();
    found.push((name || '(ไม่มีชื่อ)') + ':' + type);
    if (!name) return;
    if (type === 'password' || (!passField && /pass|pwd/i.test(name) && skip.indexOf(type) < 0)) passField = name;
    else if (skip.indexOf(type) < 0) { if (!userField) userField = name; }
    else if (type === 'hidden' || type === 'submit') fields[name] = attr_(tag, 'value');
  });
  (form[0].match(/<button\b[^>]*>/gi) || []).forEach(tag => {
    const name = attr_(tag, 'name');
    if (name) fields[name] = attr_(tag, 'value');
  });
  if (!userField || !passField) throw new Error('ไม่พบช่องชื่อผู้ใช้/รหัสผ่านในฟอร์ม (พบช่อง: ' + (found.join(', ') || 'ไม่มี') + ')');
  fields[userField] = cfg.user;
  fields[passField] = cfg.pass;
  const action = attr_(form[0].match(/<form\b[^>]*>/i)[0], 'action');
  qrReq_(action ? resolveUrl_(loginUrl, action) : loginUrl, jar, { method: 'post', payload: fields });
}

function qrSession_(cfg, force) {
  const cache = CacheService.getScriptCache();
  if (!force) {
    const c = cache.get('qr_jar');
    if (c) return JSON.parse(c);
  }
  const jar = {};
  qrLogin_(cfg, jar);
  cache.put('qr_jar', JSON.stringify(jar), 1200);
  return jar;
}

/** ดึงหน้า main.php (idcl = รหัสเครื่องลูก เว้นว่าง = ค่าเริ่มต้นของหน้า) */
function qrFetchMain_(cfg, idcl) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const jar = qrSession_(cfg, attempt > 0);
    const res = qrReq_(cfg.base + 'main.php', jar, idcl ? { method: 'post', payload: { idcl: idcl, submit: 'ตกลง' } } : {});
    const html = res.getContentText();
    if (res.getResponseCode() === 200 && html.indexOf('ผู้ส่ง') >= 0 && html.indexOf('type="password"') < 0) return html;
    CacheService.getScriptCache().remove('qr_jar');
  }
  throw new Error('ล็อกอินไม่สำเร็จ หรือรูปแบบหน้า main.php เปลี่ยนไป');
}

function qrParseMsg_(msg, recv) {
  const o = { txDate: '', direction: '', amount: 0, balance: 0 };
  const d = msg.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})/);
  if (d) {
    let y = +d[3];
    if (y < 100) y += 2500;
    if (y > 2400) y -= 543;
    o.txDate = y + '-' + p2_(d[2]) + '-' + p2_(d[1]) + ' ' + p2_(d[4]) + ':' + d[5];
  } else {
    const r = (recv || '').match(/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
    if (r) o.txDate = new Date().getFullYear() + '-' + p2_(r[2]) + '-' + p2_(r[1]) + ' ' + p2_(r[3]) + ':' + r[4];
  }
  const dir = msg.match(/เงิน(เข้า|ออก)\s*([\d,]+(?:\.\d+)?)/);
  if (dir) { o.direction = dir[1] === 'เข้า' ? 'in' : 'out'; o.amount = Number(dir[2].replace(/,/g, '')); }
  const b = msg.match(/คงเหลือ\s*([\d,]+(?:\.\d+)?)/);
  if (b) o.balance = Number(b[1].replace(/,/g, ''));
  return o;
}

function qrParse_(html) {
  const text = s => s.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
  const pick = re => { const m = html.match(re); return m ? text(m[1]) : ''; };
  const meta = { simId: pick(/รหัสซิม:\s*([^<]+)/), lastTime: pick(/เวลาล่าสุด:\s*([^<]+)/), machines: {} };
  const sel = html.match(/<select[\s\S]*?<\/select>/i);
  if (sel) (sel[0].match(/<option[^>]*>[\s\S]*?<\/option>/gi) || []).forEach(op => {
    const v = attr_(op.match(/<option[^>]*>/i)[0], 'value');
    if (v) meta.machines[v] = text(op);
  });
  const body = html.match(/<tbody[^>]*>([\s\S]*?)(?:<\/tbody>|$)/i);
  const rows = [];
  ((body ? body[1] : '').match(/<tr[\s\S]*?<\/tr>/gi) || []).forEach(tr => {
    const c = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(text);
    if (c.length < 6 || c.every(x => x === '-' || x === '')) return;
    const o = { sender: c[0], machine: c[1], price: Number(c[2].replace(/,/g, '')) || 0,
      recvTime: c[3], payTime: c[4], message: c[5] };
    Object.assign(o, qrParseMsg_(o.message, o.recvTime));
    if (!o.direction && o.price > 0) { o.direction = 'in'; o.amount = o.price; }
    o.key = [o.sender, o.machine, o.recvTime, o.payTime, o.message].join('|');
    rows.push(o);
  });
  return { meta: meta, rows: rows };
}

function qrPayload_(added, addedCoin) {
  const p = PropertiesService.getScriptProperties();
  const j = k => { try { return JSON.parse(p.getProperty(k) || '{}'); } catch (e) { return {}; } };
  return { meta: j('QR_META'), coins: readCoin_(), daily: readDaily_(), child: j('QR_CHILD_META'),
    added: added || 0, addedCoin: addedCoin || 0 };
}

/** อ่านข้อมูลที่สะสมไว้ (ไม่ติดต่อเว็บต้นทาง) */
function getQrData() {
  return qrPayload_(0);
}

/** ล็อกอิน ดึงข้อมูลล่าสุด รวมเข้าชีต QR_Log (ไม่ซ้ำ) แล้วส่งกลับ */
function syncQr() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cfg = qrCfg_();
    const first = qrParse_(qrFetchMain_(cfg, '')); // ใช้อ่านรายชื่อเครื่องลูกและรหัสซิมเท่านั้น
    const now = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('QR_META', JSON.stringify({
      simId: first.meta.simId, lastTime: first.meta.lastTime, machines: first.meta.machines, syncedAt: now }));
    return qrPayload_(0, syncChild_(cfg, first.meta.machines));
  } finally {
    lock.releaseLock();
  }
}

/** ตัวช่วยตั้งเวลา: รันฟังก์ชันนี้ครั้งเดียวในตัวแก้ไข เพื่อซิงก์อัตโนมัติทุก 15 นาที */
function setupAutoSync() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'autoSyncQr') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoSyncQr').timeBased().everyMinutes(15).create();
}

function autoSyncQr() {
  syncQr();
}

/* ====================================================================
 *  หน้า "เครื่องลูก" (ยอดหยอดเหรียญ/ธนบัตร ช่อง1 ช่อง2 และ QR)
 *  ระบบหาลิงก์เครื่องลูกจากหน้า sumall.php ให้เองอัตโนมัติ (ไม่ต้องตั้งค่า)
 *  ถ้าต้องการระบุเอง ตั้ง Script Property: QR_CHILD_URL โดยแทนรหัสด้วย {code} หรือ {id}
 *    {code} = รหัสเครื่องลูก เช่น 342002     {id} = ค่าในรายการเลือก เช่น 02
 *  ตัวอย่าง: http://vps316.vpshispeed.net/qrnew/agent/xxx/cl.php?id={code}
 *  หมายเหตุ: ดึงด้วย GET อย่างเดียว ไม่เรียกลิงก์ "เริ่มนับ" (clrcnt.php) เด็ดขาด
 * ==================================================================== */
const COIN_SHEET = 'QR_Coin';
const COIN_HEADERS = ['key', 'machine', 'channel', 'amount', 'txDate', 'fetchedAt'];

function htmlText_(s) {
  return s.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
}

function qrFetchPage_(cfg, url, marker) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const jar = qrSession_(cfg, attempt > 0);
    const res = qrReq_(url, jar);
    const html = res.getContentText();
    if (res.getResponseCode() === 200 && html.indexOf(marker) >= 0 && html.indexOf('type="password"') < 0) return html;
    CacheService.getScriptCache().remove('qr_jar');
  }
  throw new Error('เปิดหน้าไม่สำเร็จ: ' + url);
}

/**
 * เปิดหน้ารายการรับของเครื่องลูก เหมือนที่กดในเว็บ: เลือกเครื่องลูกในหน้า main.php แล้วกด "ตกลง"
 * (POST idcl) ระบบจะพาไปหน้า cli/main_cli.php ของเครื่องนั้น  idcl = ค่าในรายการเลือก เช่น 01, 02
 */
function qrFetchChild_(cfg, idcl) {
  const ok = html => html.indexOf('การรับ') >= 0 && html.indexOf('รหัสเครื่องลูก') >= 0;
  let note = '';
  for (let attempt = 0; attempt < 2; attempt++) {
    const jar = qrSession_(cfg, attempt > 0);
    let html = qrReq_(cfg.base + 'main.php', jar, { method: 'post', payload: { idcl: idcl, submit: 'ตกลง' } }).getContentText();
    if (!ok(html)) html = qrReq_(cfg.base + 'cli/main_cli.php', jar).getContentText(); // เผื่อพาไปหน้านี้ด้วย JavaScript
    if (ok(html)) return html;
    const links = (html.match(/<a\b[^>]*>/gi) || []).slice(0, 5).map(x => attr_(x, 'href')).join(' ');
    note = htmlText_(html.replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')).slice(0, 120) + ' | ลิงก์: ' + links;
    CacheService.getScriptCache().remove('qr_jar');
  }
  throw new Error('เปิดหน้ารายการรับไม่สำเร็จ (เครื่อง ' + idcl + ') → ' + note);
}

function qrParseChild_(html) {
  const pick = re => { const m = html.match(re); return m ? htmlText_(m[1]) : ''; };
  const num = x => Number(String(x).replace(/,/g, '')) || 0;
  const info = {
    code: pick(/รหัสเครื่องลูก:\s*([^<]+)/),
    expire: pick(/วันหมดอายุ:\s*([^<]+)/),
    cash: num(pick(/เงินสะสม=\s*([\d,.]+)/)),
    qr: num(pick(/QR=\s*([\d,.]+)/)),
    total: num(pick(/รวม=\s*([\d,.]+)/)),
    since: pick(/เริ่มนับตั้งแต่วันที่\s*([^<]+)/)
  };
  const body = html.match(/<tbody[^>]*>([\s\S]*?)(?:<\/tbody>|$)/i);
  const rows = [], seen = {};
  ((body ? body[1] : '').match(/<tr[\s\S]*?<\/tr>/gi) || []).forEach(tr => {
    const c = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(htmlText_);
    if (c.length < 3) return;
    const t = c[2].match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
    if (!t) return;
    let y = +t[3];
    if (y > 2400) y -= 543;
    const amount = Number(c[1].replace(/,/g, '')) || 0;
    const txDate = y + '-' + p2_(t[2]) + '-' + p2_(t[1]) + ' ' + p2_(t[4]) + ':' + t[5];
    const base = [info.code, c[0], amount, txDate].join('|');
    seen[base] = (seen[base] || 0) + 1; // รายการเหมือนกันทุกอย่างให้นับลำดับ ไม่ให้ซ้ำหาย
    rows.push({ key: base + '|' + seen[base], machine: info.code, channel: c[0], amount: amount, txDate: txDate });
  });
  return { info: info, rows: rows };
}

/** อ่านลิงก์เครื่องลูก (เช่น 342001, 342002) จากหน้า sumall.php ไม่ต้องตั้งลิงก์เอง */
function qrParseChildLinks_(html, pageUrl) {
  const links = {};
  (html.match(/<a\b[^>]*>[\s\S]*?<\/a>/gi) || []).forEach(a => {
    const code = htmlText_(a);
    const href = attr_(a.match(/<a\b[^>]*>/i)[0], 'href').replace(/&amp;/g, '&');
    if (/^\d{4,}$/.test(code) && href && !/clrcnt|clear|delete|delsms|refill/i.test(href)) links[code] = resolveUrl_(pageUrl, href);
  });
  return links;
}

function coinSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(COIN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COIN_SHEET);
    sh.appendRow(COIN_HEADERS);
    sh.getRange(1, 1, 1, COIN_HEADERS.length).setFontWeight('bold').setBackground('#dbe7f3');
    sh.setFrozenRows(1);
    ['B:B', 'E:E'].forEach(a => sh.getRange(a).setNumberFormat('@'));
  }
  return sh;
}

function readCoin_() {
  const sh = getSS_().getSheetByName(COIN_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COIN_HEADERS.length).getValues().filter(r => r[0] !== '').map(r => ({
    machine: String(r[1]), channel: String(r[2]), amount: Number(r[3]) || 0,
    txDate: r[4] instanceof Date ? Utilities.formatDate(r[4], TZ, 'yyyy-MM-dd HH:mm') : String(r[4])
  }));
}

/** อ่านหน้า sum_cli.php (สรุปรายวันของเครื่องลูก): วันที่ | โอน | ส่งเงิน | โอนส่ง | ช่อง1 | ช่อง2 | เงินสด */
function qrParseDaily_(html) {
  const num = x => Number(String(x).replace(/,/g, '')) || 0;
  const code = (html.match(/รหัสเครื่อง:\s*(\d+)/) || [])[1] || '';
  const rows = [];
  (html.match(/<tr[\s\S]*?<\/tr>/gi) || []).forEach(tr => {
    const c = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(htmlText_);
    if (c.length < 7 || !/^\d{4}-\d{2}-\d{2}$/.test(c[0])) return; // ข้ามแถว "รวม" และ "ทั้งหมด"
    rows.push({ date: c[0], transfer: num(c[1]), send: num(c[2]), transferSend: num(c[3]), ch1: num(c[4]), ch2: num(c[5]), cash: num(c[6]) });
  });
  return { code: code, rows: rows };
}

const DAILY_SHEET = 'QR_Daily';
const DAILY_HEADERS = ['key', 'machine', 'date', 'transfer', 'send', 'transferSend', 'ch1', 'ch2', 'cash', 'fetchedAt'];

function dailySheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(DAILY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DAILY_SHEET);
    sh.appendRow(DAILY_HEADERS);
    sh.getRange(1, 1, 1, DAILY_HEADERS.length).setFontWeight('bold').setBackground('#dbe7f3');
    sh.setFrozenRows(1);
    sh.getRange('B:C').setNumberFormat('@');
  }
  return sh;
}

function readDaily_() {
  const sh = getSS_().getSheetByName(DAILY_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, DAILY_HEADERS.length).getValues().filter(r => r[0] !== '').map(r => ({
    machine: String(r[1]),
    date: r[2] instanceof Date ? Utilities.formatDate(r[2], TZ, 'yyyy-MM-dd') : String(r[2]),
    transfer: Number(r[3]) || 0, send: Number(r[4]) || 0, transferSend: Number(r[5]) || 0,
    ch1: Number(r[6]) || 0, ch2: Number(r[7]) || 0, cash: Number(r[8]) || 0
  }));
}

/** เพิ่มวันใหม่ และอัปเดตวันที่ตัวเลขเปลี่ยน (เช่น วันนี้ที่ยอดยังเพิ่ม) คืนจำนวนแถวที่เพิ่ม/แก้ */
function upsertDaily_(rows) {
  if (!rows.length) return 0;
  const sh = dailySheet_(), last = sh.getLastRow(), have = {};
  if (last > 1) sh.getRange(2, 1, last - 1, DAILY_HEADERS.length).getValues().forEach((r, i) => have[String(r[0])] = { row: i + 2, v: r });
  const now = new Date().toISOString(), append = [];
  let changed = 0;
  rows.forEach(r => {
    const vals = [r.key, r.machine, r.date, r.transfer, r.send, r.transferSend, r.ch1, r.ch2, r.cash, now];
    const old = have[r.key];
    if (!old) { append.push(vals); return; }
    const same = [3, 4, 5, 6, 7, 8].every(k => Number(old.v[k]) === vals[k]);
    if (!same) { sh.getRange(old.row, 1, 1, DAILY_HEADERS.length).setValues([vals]); changed++; }
  });
  if (append.length) sh.getRange(last + 1, 1, append.length, DAILY_HEADERS.length).setValues(append);
  return append.length + changed;
}

function appendCoins_(rows) {
  if (!rows.length) return 0;
  const sh = coinSheet_(), last = sh.getLastRow(), have = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => have[String(r[0])] = 1);
  const now = new Date().toISOString();
  const add = rows.filter(x => !have[x.key]).map(x => [x.key, x.machine, x.channel, x.amount, x.txDate, now]);
  if (add.length) sh.getRange(last + 1, 1, add.length, COIN_HEADERS.length).setValues(add);
  return add.length;
}

/** ดึงข้อมูลเครื่องลูกทุกเครื่อง: สรุปรายวันจาก sum_cli.php + รายการรับละเอียด (ถ้าเปิดได้) */
function syncChild_(cfg, machines) {
  const p = PropertiesService.getScriptProperties();
  const meta = { machines: {}, errors: [], notes: [], syncedAt: new Date().toISOString() };
  let added = 0;
  const ids = Object.keys(machines || {});
  if (!ids.length) meta.errors.push('ไม่พบรายชื่อเครื่องลูกจากหน้า main.php');
  else {
    let links = null;
    try {
      const su = cfg.base + 'sumall.php?back=0&end=0';
      links = qrParseChildLinks_(qrFetchPage_(cfg, su, 'สรุปเครื่องลูก'), su);
    } catch (e) { meta.errors.push('เปิดหน้า sumall.php ไม่ได้: ' + (e.message || e)); }
    if (links) ids.forEach(id => {
      const code = machines[id], link = links[code];
      if (!link) { meta.errors.push('ไม่พบลิงก์ของเครื่อง ' + code + ' ในหน้า sumall.php'); return; }
      try { // สรุปรายวันย้อนหลังสูงสุด 60 วัน
        const dUrl = /[?&]back=\d+/.test(link) ? link.replace(/([?&])back=\d+/, '$1back=59')
          : link + (link.indexOf('?') < 0 ? '?' : '&') + 'back=59&end=0';
        const d = qrParseDaily_(qrFetchPage_(cfg, dUrl, 'ช่อง1'));
        added += upsertDaily_(d.rows.map(r => Object.assign({ machine: code, key: code + '|' + r.date }, r)));
      } catch (e) { meta.errors.push(code + ' (สรุปรายวัน): ' + (e.message || e)); }
      try { // รายการรับละเอียดทีละครั้ง (ใช้ได้เมื่อเปิดหน้า cli/main_cli.php ได้)
        const r = qrParseChild_(qrFetchChild_(cfg, id));
        if (r.info.code && String(r.info.code) !== String(code)) throw new Error('ได้หน้าของเครื่อง ' + r.info.code);
        meta.machines[r.info.code || code] = r.info;
        added += appendCoins_(r.rows);
      } catch (e) { meta.notes.push(code + ' (รายการรับละเอียด): ' + String(e.message || e).slice(0, 300)); }
    });
  }
  p.setProperty('QR_CHILD_META', JSON.stringify(meta));
  return added;
}
