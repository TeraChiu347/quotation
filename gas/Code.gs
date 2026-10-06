// ── Sheet Names ──
const SH_CONFIG    = '設定';
const SH_CLIENTS   = '甲方名單';
const SH_TEMPLATES = '常用項目';
const SH_QUOTES    = '報價紀錄';
const SH_PAYMENTS  = '收款記錄';

const DRIVE_ROOT_FOLDER_NAME = '報價系統-案件資料夾';

function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  if (token) {
    return HtmlService.createTemplateFromFile('client').evaluate()
      .setTitle('報價單')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('報價單系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Spreadsheet Helper ──
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheetRows(name) {
  const s = ss().getSheetByName(name);
  return s && s.getLastRow() > 0 ? s.getDataRange().getValues() : [];
}

// ── Config ──
function configMap() {
  const map = {};
  sheetRows(SH_CONFIG).forEach(r => { if (r[0]) map[String(r[0])] = r[1]; });
  return map;
}

function setConfig(key, value) {
  const sheet = ss().getSheetByName(SH_CONFIG);
  const data  = sheet.getDataRange().getValues();
  const idx   = data.findIndex(r => r[0] === key);
  if (idx >= 0) sheet.getRange(idx + 1, 2).setValue(value);
  else          sheet.appendRow([key, value]);
}

function saveCompanyConfig(fields) {
  Object.entries(fields).forEach(([k, v]) => setConfig(k, v));
  return true;
}

// ── Init Sheets ──
function initSheets() {
  const s = ss();

  if (!s.getSheetByName(SH_CONFIG)) {
    const sh = s.insertSheet(SH_CONFIG);
    const year = new Date().getFullYear();
    [
      ['COMPANY_NAME',    ''],
      ['COMPANY_ADDRESS', ''],
      ['COMPANY_PHONE',   ''],
      ['COMPANY_TAX_ID',  ''],
      ['LOGO_1_URL',      ''],
      ['LOGO_1_NAME',     'LOGO 1'],
      ['LOGO_2_URL',      ''],
      ['LOGO_2_NAME',     'LOGO 2'],
      ['TAX_RATE',        0.05],
      ['QUOTE_YEAR',      year],
      ['QUOTE_COUNTER',   0],
      ['PAYMENT_DEFAULT', '1. 簽約後預付總金額 30% 作為訂金\n2. 專案完成並驗收後 7 日內支付尾款'],
      ['NOTES_DEFAULT',   '1. 本報價未列項目需另行報價\n2. 客戶延遲提供資料，工期將等比例順延'],
    ].forEach(r => sh.appendRow(r));
  }

  if (!s.getSheetByName(SH_CLIENTS)) {
    s.insertSheet(SH_CLIENTS)
      .appendRow(['ID', '公司名稱', '電話', '聯絡人', 'Email']);
  }

  if (!s.getSheetByName(SH_TEMPLATES)) {
    s.insertSheet(SH_TEMPLATES)
      .appendRow(['類別', '項目名稱', '規格說明', '單位', '單價']);
  }

  if (!s.getSheetByName(SH_QUOTES)) {
    s.insertSheet(SH_QUOTES)
      .appendRow(['報價編號','建立日期','有效天數','甲方公司','聯絡人','電話','Email',
                  'LOGO','品項JSON','付款條件','備註條款','未稅小計','稅額','含稅總計','狀態',
                  'ShareToken','DriveFolderID','DriveFolderURL','開案時間']);
  }

  ensurePaymentSheet_();

  return true;
}

// ── Migration: add case-management columns to an existing 報價紀錄 sheet ──
function migrateQuoteSheetV2() {
  const sheet = ss().getSheetByName(SH_QUOTES);
  if (!sheet) return false;
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const newCols = ['ShareToken', 'DriveFolderID', 'DriveFolderURL', '開案時間'];
  const missing = newCols.filter(c => header.indexOf(c) === -1);
  if (missing.length) {
    sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
  }
  ensurePaymentSheet_();
  return true;
}

function ensurePaymentSheet_() {
  if (!ss().getSheetByName(SH_PAYMENTS)) {
    ss().insertSheet(SH_PAYMENTS)
      .appendRow(['記錄ID', '報價編號', '回報時間', '金額', '匯款後五碼', '付款方式', '回報來源', '確認狀態', '確認時間', '備註']);
  }
}

// ── Load All (admin only — never expose this to the public client page) ──
function loadAll() {
  const tz = Session.getScriptTimeZone();

  const config = configMap();

  const clients = sheetRows(SH_CLIENTS).slice(1)
    .map(r => ({ id: r[0], company: r[1], phone: r[2], contact: r[3], email: r[4] }))
    .filter(c => c.company);

  const templates = {};
  sheetRows(SH_TEMPLATES).slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const cat = String(r[0]);
    if (!templates[cat]) templates[cat] = [];
    templates[cat].push({ name: r[1], spec: r[2] || '', unit: r[3] || '式', price: Number(r[4]) || 0 });
  });

  const quotes = sheetRows(SH_QUOTES).slice(1).reverse()
    .map(r => quoteRowToObj_(r, tz))
    .filter(q => q.no);

  ensurePaymentSheet_();
  const payments = sheetRows(SH_PAYMENTS).slice(1).map(r => ({
    id:       r[0],
    no:       r[1],
    time:     r[2] ? Utilities.formatDate(new Date(r[2]), tz, 'yyyy-MM-dd HH:mm') : '',
    amount:   Number(r[3]) || 0,
    last5:    r[4] || '',
    method:   r[5] || '轉帳',
    source:   r[6] || 'admin',
    status:   r[7] || '待確認',
    confirmedAt: r[8] ? Utilities.formatDate(new Date(r[8]), tz, 'yyyy-MM-dd HH:mm') : '',
    note:     r[9] || '',
  })).filter(p => p.id);

  return { config, clients, templates, quotes, payments };
}

function quoteRowToObj_(r, tz) {
  const rawDate = r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') : '';
  return {
    no:       r[0],
    date:     rawDate ? rawDate.replace(/-/g, '/') : '',
    rawDate,
    validDays: Number(r[2]) || 7,
    client:   r[3], contact: r[4], phone: r[5], email: r[6],
    logo:     Number(r[7]) || 1,
    items:    r[8] ? JSON.parse(r[8]) : [],
    payment:  r[9]  || '',
    notes:    r[10] || '',
    subtotal: Number(r[11]) || 0,
    tax:      Number(r[12]) || 0,
    total:    Number(r[13]) || 0,
    status:   r[14] || '草稿',
    shareToken:     r[15] || '',
    driveFolderId:  r[16] || '',
    driveFolderUrl: r[17] || '',
    openedAt:       r[18] ? Utilities.formatDate(new Date(r[18]), tz, 'yyyy-MM-dd HH:mm') : '',
  };
}

// ── Quote Number ──
function nextQuoteNo() {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const cfg  = configMap();
    const year = new Date().getFullYear();
    let counter = 1;
    if (String(cfg.QUOTE_YEAR) === String(year)) {
      counter = (Number(cfg.QUOTE_COUNTER) || 0) + 1;
    }
    setConfig('QUOTE_YEAR',    year);
    setConfig('QUOTE_COUNTER', counter);
    return `${year}-${String(counter).padStart(3, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

// ── Save Quote ──
function saveQuote(q) {
  const sheet = ss().getSheetByName(SH_QUOTES);
  const data  = sheet.getDataRange().getValues();
  const isNew = !q.no;
  if (isNew) q.no = nextQuoteNo();

  let createdDate = new Date();
  let idx = -1;
  if (!isNew) {
    idx = data.findIndex(r => r[0] === q.no);
    if (idx > 0 && data[idx][1]) createdDate = data[idx][1]; // preserve original creation date on edit
  }

  const row = [
    q.no, createdDate, q.validDays || 7,
    q.client || '', q.contact || '', q.phone || '', q.email || '',
    q.logo || 1, JSON.stringify(q.items || []),
    q.payment || '', q.notes || '',
    q.subtotal || 0, q.tax || 0, q.total || 0,
    q.status || '草稿',
  ];

  if (!isNew && idx > 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
    return q.no;
  }
  sheet.appendRow(row);
  return q.no;
}

// ── Delete Quote ──
function deleteQuote(no) {
  const sheet = ss().getSheetByName(SH_QUOTES);
  const data  = sheet.getDataRange().getValues();
  const idx   = data.findIndex(r => r[0] === no);
  if (idx > 0) sheet.deleteRow(idx + 1);
  return true;
}

// ── Save Client ──
function saveClient(c) {
  const sheet = ss().getSheetByName(SH_CLIENTS);
  const data  = sheet.getDataRange().getValues();
  const id    = c.id || Utilities.getUuid();
  const row   = [id, c.company || '', c.phone || '', c.contact || '', c.email || ''];

  if (c.id) {
    const idx = data.findIndex(r => r[0] === c.id);
    if (idx > 0) { sheet.getRange(idx + 1, 1, 1, 5).setValues([row]); return id; }
  }
  sheet.appendRow(row);
  return id;
}

// ── Delete Client ──
function deleteClient(id) {
  const sheet = ss().getSheetByName(SH_CLIENTS);
  const data  = sheet.getDataRange().getValues();
  const idx   = data.findIndex(r => r[0] === id);
  if (idx > 0) sheet.deleteRow(idx + 1);
  return true;
}

// ── Save Template Item ──
function saveTemplateItem(item) {
  ss().getSheetByName(SH_TEMPLATES)
    .appendRow([item.category, item.name, item.spec || '', item.unit || '式', Number(item.price) || 0]);
  return true;
}

// ── Delete Template Item ──
function deleteTemplateItem(cat, name) {
  const sheet = ss().getSheetByName(SH_TEMPLATES);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === cat && data[i][1] === name) sheet.deleteRow(i + 1);
  }
  return true;
}

// ── Case Opening: generate share link + Drive folder ──
function openCase(no) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = ss().getSheetByName(SH_QUOTES);
    const data  = sheet.getDataRange().getValues();
    const idx   = data.findIndex(r => r[0] === no);
    if (idx < 0) throw new Error('報價單不存在');
    const row = data[idx];

    let token = row[15];
    if (!token) {
      token = Utilities.getUuid();
      sheet.getRange(idx + 1, 16).setValue(token);
    }

    const email  = row[6] || '';
    const client = row[3] || '';
    const folder = getOrCreateCaseFolder_(sheet, idx, no, email, client);

    const status = row[14] || '草稿';
    if (status === '草稿' || status === '已送出') {
      sheet.getRange(idx + 1, 15).setValue('進行中');
    }
    if (!row[18]) sheet.getRange(idx + 1, 19).setValue(new Date());

    const shareUrl = ScriptApp.getService().getUrl() + '?token=' + token;
    return { shareUrl, folderUrl: folder.url, warning: folder.warning || '' };
  } finally {
    lock.releaseLock();
  }
}

// ── The single parent folder all case folders live under, ID cached in 設定 (avoids
//    a name-based Drive search, which the narrow drive.file scope may not support) ──
function getOrCreateRootFolder_() {
  const cfg = configMap();
  if (cfg.DRIVE_ROOT_FOLDER_ID) {
    try { return DriveApp.getFolderById(cfg.DRIVE_ROOT_FOLDER_ID); }
    catch (err) { /* was deleted externally — fall through and recreate */ }
  }
  const folder = DriveApp.createFolder(DRIVE_ROOT_FOLDER_NAME);
  setConfig('DRIVE_ROOT_FOLDER_ID', folder.getId());
  return folder;
}

// ── Create or reuse the per-case Drive folder, shared with the client's email ──
function getOrCreateCaseFolder_(sheet, rowIdx, no, email, clientName) {
  let folder = null;
  const existingId = sheet.getRange(rowIdx + 1, 17).getValue();

  if (existingId) {
    try { folder = DriveApp.getFolderById(existingId); }
    catch (err) { folder = null; } // folder was deleted/trashed externally — fall through and recreate
  }

  if (!folder) {
    folder = getOrCreateRootFolder_().createFolder(`${no}_${clientName || '未命名客戶'}`);
    sheet.getRange(rowIdx + 1, 17).setValue(folder.getId());
    sheet.getRange(rowIdx + 1, 18).setValue(folder.getUrl());
  }

  let warning = '';
  if (email) {
    try { folder.addEditor(email); }
    catch (err) { warning = `資料夾已建立，但無法把權限分享給 ${email}，請手動到 Drive 確認信箱是否正確`; }
  } else {
    warning = '此報價單未填 Email，資料夾已建立但尚未分享給任何人';
  }

  return { url: folder.getUrl(), warning };
}

// ── Client-facing read (never reuse loadAll() here — must not leak all quotes/clients) ──
function getQuoteByToken(token) {
  const tz = Session.getScriptTimeZone();
  const rows = sheetRows(SH_QUOTES).slice(1);
  const row  = rows.find(r => r[15] && r[15] === token);
  if (!row) return null;

  const quote = quoteRowToObj_(row, tz);
  const cfg = configMap();
  const config = {
    COMPANY_NAME: cfg.COMPANY_NAME, COMPANY_ADDRESS: cfg.COMPANY_ADDRESS,
    COMPANY_PHONE: cfg.COMPANY_PHONE, COMPANY_TAX_ID: cfg.COMPANY_TAX_ID,
    LOGO_1_URL: cfg.LOGO_1_URL, LOGO_2_URL: cfg.LOGO_2_URL,
    BANK_NAME: cfg.BANK_NAME, BANK_CODE: cfg.BANK_CODE,
    BANK_HOLDER: cfg.BANK_HOLDER, BANK_ACCOUNT: cfg.BANK_ACCOUNT,
  };

  ensurePaymentSheet_();
  const ledger = sheetRows(SH_PAYMENTS).slice(1)
    .filter(r => r[1] === quote.no)
    .map(r => ({
      time:   r[2] ? Utilities.formatDate(new Date(r[2]), tz, 'yyyy-MM-dd HH:mm') : '',
      amount: Number(r[3]) || 0,
      last5:  r[4] || '',
      status: r[7] || '待確認',
    }));
  const paidConfirmed = ledger.filter(p => p.status === '已確認').reduce((s, p) => s + p.amount, 0);

  return { quote, config, folderUrl: quote.driveFolderUrl, ledger, paidConfirmed };
}

// ── Client-facing payment report (always re-derive `no` from the token, never trust client input) ──
function reportPayment(token, amount, last5, note) {
  const rows = sheetRows(SH_QUOTES).slice(1);
  const row  = rows.find(r => r[15] && r[15] === token);
  if (!row) throw new Error('找不到對應的報價單');
  const no = row[0];

  const amt = Number(amount);
  if (!amt || amt <= 0) throw new Error('請輸入正確的金額');
  if (!/^\d{5}$/.test(String(last5 || '').trim())) throw new Error('請輸入匯款帳號末五碼（5 位數字）');

  ensurePaymentSheet_();
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    ss().getSheetByName(SH_PAYMENTS).appendRow([
      Utilities.getUuid(), no, new Date(), amt, String(last5).trim(),
      '轉帳', 'client', '待確認', '', note || '',
    ]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

// ── Admin: list payments awaiting confirmation ──
function listPendingPayments() {
  ensurePaymentSheet_();
  const tz = Session.getScriptTimeZone();
  return sheetRows(SH_PAYMENTS).slice(1)
    .filter(r => (r[7] || '待確認') === '待確認')
    .map(r => ({
      id: r[0], no: r[1],
      time: r[2] ? Utilities.formatDate(new Date(r[2]), tz, 'yyyy-MM-dd HH:mm') : '',
      amount: Number(r[3]) || 0, last5: r[4] || '', method: r[5] || '轉帳', source: r[6] || 'admin',
    }));
}

// ── Admin: manually log a payment (e.g. cash received in person) ──
function addManualPayment(no, amount, method, note) {
  const amt = Number(amount);
  if (!no || !amt || amt <= 0) throw new Error('請輸入正確的報價編號與金額');
  ensurePaymentSheet_();
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    ss().getSheetByName(SH_PAYMENTS).appendRow([
      Utilities.getUuid(), no, new Date(), amt, '',
      method || '現金', 'admin', '已確認', new Date(), note || '',
    ]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

// ── Admin: confirm or reject a pending payment report ──
function confirmPayment(recordId, approve, note) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = ss().getSheetByName(SH_PAYMENTS);
    const data  = sheet.getDataRange().getValues();
    const idx   = data.findIndex(r => r[0] === recordId);
    if (idx < 0) throw new Error('找不到這筆收款記錄');
    sheet.getRange(idx + 1, 8).setValue(approve ? '已確認' : '已駁回');
    sheet.getRange(idx + 1, 9).setValue(new Date());
    if (note) sheet.getRange(idx + 1, 10).setValue(note);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ── Admin: update case status (進行中/已完成/已取消) ──
function setCaseStatus(no, status) {
  const sheet = ss().getSheetByName(SH_QUOTES);
  const data  = sheet.getDataRange().getValues();
  const idx   = data.findIndex(r => r[0] === no);
  if (idx < 0) throw new Error('報價單不存在');
  sheet.getRange(idx + 1, 15).setValue(status);
  return true;
}
