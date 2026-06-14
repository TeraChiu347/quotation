// ── Sheet Names ──
const SH_CONFIG    = '設定';
const SH_CLIENTS   = '甲方名單';
const SH_TEMPLATES = '常用項目';
const SH_QUOTES    = '報價紀錄';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('報價單系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
                  'LOGO','品項JSON','付款條件','備註條款','未稅小計','稅額','含稅總計','狀態']);
  }

  return true;
}

// ── Load All ──
function loadAll() {
  const tz = Session.getScriptTimeZone();

  // Config
  const config = configMap();

  // Clients
  const clients = sheetRows(SH_CLIENTS).slice(1)
    .map(r => ({ id: r[0], company: r[1], phone: r[2], contact: r[3], email: r[4] }))
    .filter(c => c.company);

  // Templates
  const templates = {};
  sheetRows(SH_TEMPLATES).slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const cat = String(r[0]);
    if (!templates[cat]) templates[cat] = [];
    templates[cat].push({ name: r[1], spec: r[2] || '', unit: r[3] || '式', price: Number(r[4]) || 0 });
  });

  // Quotes
  const quotes = sheetRows(SH_QUOTES).slice(1).reverse()
    .map(r => {
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
      };
    })
    .filter(q => q.no);

  return { config, clients, templates, quotes };
}

// ── Quote Number ──
function nextQuoteNo() {
  const cfg  = configMap();
  const year = new Date().getFullYear();
  let counter = 1;
  if (String(cfg.QUOTE_YEAR) === String(year)) {
    counter = (Number(cfg.QUOTE_COUNTER) || 0) + 1;
  }
  setConfig('QUOTE_YEAR',    year);
  setConfig('QUOTE_COUNTER', counter);
  return `${year}-${String(counter).padStart(3, '0')}`;
}

// ── Save Quote ──
function saveQuote(q) {
  const sheet = ss().getSheetByName(SH_QUOTES);
  const data  = sheet.getDataRange().getValues();
  const isNew = !q.no;
  if (isNew) q.no = nextQuoteNo();

  const row = [
    q.no, new Date(), q.validDays || 7,
    q.client || '', q.contact || '', q.phone || '', q.email || '',
    q.logo || 1, JSON.stringify(q.items || []),
    q.payment || '', q.notes || '',
    q.subtotal || 0, q.tax || 0, q.total || 0,
    q.status || '草稿',
  ];

  if (!isNew) {
    const idx = data.findIndex(r => r[0] === q.no);
    if (idx > 0) {
      sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
      return q.no;
    }
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
