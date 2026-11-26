/**
 * 個人記帳 — Google Apps Script 後端
 * ------------------------------------------------------------
 * 用法（詳見 README.md）：
 *   1. 在 Google 雲端硬碟建立一個試算表 → 擴充功能 → Apps Script
 *   2. 把本檔內容整個貼上、儲存
 *   3. 上方函式選單選 setup → 執行 → 授權 → 從「執行紀錄」複製 Token
 *   4. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我／存取權：所有人）
 *   5. 把 /exec 網址與 Token 填到網頁的「⚙ 設定」
 *
 * 試算表結構：
 *   支出 / 收入：ID | 日期 | 時間 | 類別 | 項目 | 金額 | 幣別 | 備註 | 建立時間
 *   類別：類型（支出/收入）| 名稱
 *   設定：鍵 | 值 | 說明
 */

const CFG = {
  SHEETS: { expense: '支出', income: '收入' },
  CAT_SHEET: '類別',
  SET_SHEET: '設定',
  HEADERS: ['ID', '日期', '時間', '類別', '項目', '金額', '幣別', '備註', '建立時間'],
  ROW_FORMATS: ['@', '@', '@', '@', '@', '#,##0.00', '@', '@', '@'],
  TYPE_LABEL: { expense: '支出', income: '收入' },
  DEFAULT_CATEGORIES: {
    expense: ['早餐', '午餐', '晚餐', '飲料', '交通', '購物', '日用品', '娛樂', '醫療', '住房', '水電瓦斯', '通訊', '教育', '其他'],
    income: ['薪資', '獎金', '投資', '副業', '利息', '退款', '其他'],
  },
  DEFAULT_SETTINGS: {
    baseCurrency: 'TWD',
    rates: { TWD: 1, USD: 31, JPY: 0.2, EUR: 34, CNY: 4.3, HKD: 4 },
  },
};

/* ============================================================
 * 一次性設定：建立工作表、預設類別、產生 API Token
 * ============================================================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從試算表的「擴充功能 → Apps Script」開啟本專案後再執行 setup');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(CFG.SHEETS).forEach(function (type) { ensureRecordSheet_(ss, CFG.SHEETS[type]); });
  ensureCategorySheet_(ss);
  ensureSettingsSheet_(ss);

  // 移除空白的預設工作表
  ['工作表1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  let token = props.getProperty('API_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('API_TOKEN', token);
  }
  Logger.log('✅ 設定完成。你的 API Token（請貼到網頁設定，不要公開）：\n%s', token);
}

/** 忘記 Token 時執行，從執行紀錄查看 */
function showToken() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('API_TOKEN') || '尚未產生，請先執行 setup');
}

/** 懷疑 Token 外流時執行：產生新 Token，舊的立即失效 */
function resetToken() {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', token);
  Logger.log('🔑 新 Token：\n%s', token);
}

/* ============================================================
 * Web App 進入點
 * ============================================================ */
function doGet() {
  return json_({ ok: true, data: { message: '個人記帳 API 運作中，請使用 POST。' } });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkToken_(req.token);
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('未知的動作：' + req.action);
    const mutating = req.action !== 'load' && req.action !== 'ping';
    if (!mutating) return json_({ ok: true, data: fn(req) });
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      return json_({ ok: true, data: fn(req) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

const ACTIONS = {
  ping: function () {
    const ss = getSs_();
    return { name: ss.getName(), sheetUrl: ss.getUrl(), timeZone: ss.getSpreadsheetTimeZone() };
  },

  load: function () {
    const ss = getSs_();
    return {
      expense: readRecords_(ss, 'expense'),
      income: readRecords_(ss, 'income'),
      categories: readCategories_(ss),
      settings: readSettings_(ss),
      sheetUrl: ss.getUrl(),
    };
  },

  add: function (req) {
    const type = checkType_(req.type);
    const rec = cleanRecord_(req.record);
    const sh = recordSheet_(type);
    if (findRow_(sh, rec.id)) return rec; // 重送（離線佇列）時避免重複
    writeRows_(sh, sh.getLastRow() + 1, [recordToRow_(rec)]);
    ensureCategoryExists_(type, rec.category);
    return rec;
  },

  update: function (req) {
    const type = checkType_(req.type);
    const rec = cleanRecord_(req.record);
    const sh = recordSheet_(type);
    const row = findRow_(sh, rec.id);
    if (!row) throw new Error('找不到要更新的記錄（可能已被刪除）');
    writeRows_(sh, row, [recordToRow_(rec)]);
    ensureCategoryExists_(type, rec.category);
    return rec;
  },

  delete: function (req) {
    const type = checkType_(req.type);
    const sh = recordSheet_(type);
    const row = findRow_(sh, String(req.id || ''));
    if (row) sh.deleteRow(row);
    return { deleted: !!row };
  },

  addCategory: function (req) {
    const type = checkType_(req.type);
    const name = str_(req.name, 30);
    if (!name) throw new Error('類別名稱不可空白');
    ensureCategoryExists_(type, name);
    return readCategories_(getSs_());
  },

  deleteCategory: function (req) {
    const type = checkType_(req.type);
    const name = str_(req.name, 30);
    const cats = readCategories_(getSs_());
    cats[type] = cats[type].filter(function (c) { return c !== name; });
    writeCategories_(cats);
    return cats;
  },

  saveSettings: function (req) {
    const s = req.settings || {};
    const base = String(s.baseCurrency || 'TWD').toUpperCase();
    if (!/^[A-Z]{3}$/.test(base)) throw new Error('基準幣別格式錯誤');
    const rates = {};
    Object.keys(s.rates || {}).forEach(function (k) {
      const code = String(k).toUpperCase();
      const v = Number(s.rates[k]);
      if (/^[A-Z]{3}$/.test(code) && isFinite(v) && v > 0) rates[code] = v;
    });
    rates[base] = 1;
    writeSettings_({ baseCurrency: base, rates: rates });
    return readSettings_(getSs_());
  },

  importRecords: function (req) {
    const list = Array.isArray(req.records) ? req.records : [];
    const result = { added: 0, skipped: 0, errors: [] };
    ['expense', 'income'].forEach(function (type) {
      const sh = recordSheet_(type);
      const existing = existingIds_(sh);
      const rows = [];
      const newCats = {};
      list.forEach(function (r, i) {
        if (!r || r.type !== type) return;
        let rec;
        try { rec = cleanRecord_(r); } catch (err) { result.errors.push('第 ' + (i + 1) + ' 筆：' + err.message); return; }
        if (existing[rec.id]) { result.skipped++; return; }
        existing[rec.id] = true;
        rows.push(recordToRow_(rec));
        newCats[rec.category] = true;
      });
      if (rows.length) writeRows_(sh, sh.getLastRow() + 1, rows);
      result.added += rows.length;
      Object.keys(newCats).forEach(function (c) { ensureCategoryExists_(type, c); });
    });
    if (result.errors.length > 20) result.errors = result.errors.slice(0, 20).concat(['…']);
    return result;
  },
};

/* ============================================================
 * 工作表存取
 * ============================================================ */
function getSs_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表，請先在 Apps Script 執行 setup');
  return ss;
}

function recordSheet_(type) {
  const ss = getSs_();
  return ss.getSheetByName(CFG.SHEETS[type]) || ensureRecordSheet_(ss, CFG.SHEETS[type]);
}

function ensureRecordSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CFG.HEADERS.length).setValues([CFG.HEADERS])
      .setFontWeight('bold').setBackground('#d9f2ec');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 90);
    sh.setColumnWidth(5, 220);
  }
  return sh;
}

function ensureCategorySheet_(ss) {
  let sh = ss.getSheetByName(CFG.CAT_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.CAT_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 2).setValues([['類型', '名稱']]).setFontWeight('bold').setBackground('#d9f2ec');
    sh.setFrozenRows(1);
    const rows = [];
    ['expense', 'income'].forEach(function (t) {
      CFG.DEFAULT_CATEGORIES[t].forEach(function (c) { rows.push([CFG.TYPE_LABEL[t], c]); });
    });
    sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  }
  return sh;
}

function ensureSettingsSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SET_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.SET_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['鍵', '值', '說明']]).setFontWeight('bold').setBackground('#d9f2ec');
    sh.setFrozenRows(1);
    writeSettings_(CFG.DEFAULT_SETTINGS, sh);
  }
  return sh;
}

/** 依序寫入多列，並先把該範圍設為文字格式（避免「1/2」被轉成日期、ID 被轉成數字） */
function writeRows_(sh, startRow, rows) {
  const range = sh.getRange(startRow, 1, rows.length, CFG.HEADERS.length);
  range.setNumberFormats(rows.map(function () { return CFG.ROW_FORMATS; }));
  range.setValues(rows);
}

function readRecords_(ss, type) {
  const sh = ss.getSheetByName(CFG.SHEETS[type]);
  if (!sh) return [];
  const n = sh.getLastRow();
  if (n < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const values = sh.getRange(2, 1, n - 1, CFG.HEADERS.length).getValues();
  const out = [];
  const missingId = [];
  values.forEach(function (r, i) {
    const empty = r.every(function (v) { return v === '' || v === null; });
    if (empty) return;
    if (r[0] === '' && r[1] !== '' && r[5] !== '') { // 使用者在試算表手動新增、沒填 ID
      r[0] = Utilities.getUuid().slice(0, 13);
      missingId.push({ row: i + 2, id: r[0] });
    }
    if (r[0] === '') return;
    out.push(rowToRecord_(r, tz));
  });
  missingId.forEach(function (m) { sh.getRange(m.row, 1).setNumberFormat('@').setValue(m.id); });
  return out;
}

function rowToRecord_(r, tz) {
  const date = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : normDate_(String(r[1]));
  const time = r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'HH:mm') : String(r[2] || '00:00').slice(0, 5);
  return {
    id: String(r[0]),
    date: date,
    time: /^\d:\d\d$/.test(time) ? '0' + time : time,
    category: unsafe_(String(r[3] || '其他')),
    item: unsafe_(String(r[4] || '')),
    amount: Number(r[5]) || 0,
    currency: String(r[6] || 'TWD').toUpperCase(),
    note: unsafe_(String(r[7] || '')),
    createdAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8] || ''),
  };
}

function recordToRow_(r) {
  return [r.id, r.date, r.time, safe_(r.category), safe_(r.item), r.amount, r.currency, safe_(r.note), r.createdAt];
}

function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const cell = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function existingIds_(sh) {
  const map = {};
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0] !== '') map[String(r[0])] = true; });
  return map;
}

/* ---------- 類別 ---------- */
function readCategories_(ss) {
  const sh = ss.getSheetByName(CFG.CAT_SHEET) || ensureCategorySheet_(ss);
  const cats = { expense: [], income: [] };
  if (sh.getLastRow() < 2) return cats;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    const label = String(r[0]).trim();
    const name = unsafe_(String(r[1]).trim());
    const type = label === '收入' || label === 'income' ? 'income' : (label === '支出' || label === 'expense' ? 'expense' : null);
    if (type && name && cats[type].indexOf(name) < 0) cats[type].push(name);
  });
  return cats;
}

function writeCategories_(cats) {
  const ss = getSs_();
  const sh = ss.getSheetByName(CFG.CAT_SHEET) || ensureCategorySheet_(ss);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = [];
  ['expense', 'income'].forEach(function (t) {
    cats[t].forEach(function (c) { rows.push([CFG.TYPE_LABEL[t], safe_(c)]); });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}

function ensureCategoryExists_(type, name) {
  if (!name) return;
  const ss = getSs_();
  const cats = readCategories_(ss);
  if (cats[type].indexOf(name) >= 0) return;
  const sh = ss.getSheetByName(CFG.CAT_SHEET);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[CFG.TYPE_LABEL[type], safe_(name)]]);
}

/* ---------- 設定（基準幣別、匯率） ---------- */
function readSettings_(ss) {
  const sh = ss.getSheetByName(CFG.SET_SHEET) || ensureSettingsSheet_(ss);
  const s = { baseCurrency: 'TWD', rates: {} };
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      const k = String(r[0]).trim();
      if (k === 'baseCurrency') s.baseCurrency = String(r[1]).trim().toUpperCase() || 'TWD';
      else if (k.indexOf('rate.') === 0) {
        const v = Number(r[1]);
        if (isFinite(v) && v > 0) s.rates[k.slice(5).toUpperCase()] = v;
      }
    });
  }
  s.rates[s.baseCurrency] = 1;
  return s;
}

function writeSettings_(s, shOpt) {
  const sh = shOpt || getSs_().getSheetByName(CFG.SET_SHEET) || ensureSettingsSheet_(getSs_());
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  const rows = [['baseCurrency', s.baseCurrency, '統計時換算的基準幣別']];
  Object.keys(s.rates).sort().forEach(function (c) {
    rows.push(['rate.' + c, s.rates[c], '1 ' + c + ' = ? ' + s.baseCurrency]);
  });
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

/* ============================================================
 * 驗證與工具
 * ============================================================ */
function checkToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!expected) throw new Error('伺服器尚未設定 Token，請在 Apps Script 執行 setup');
  if (!token || String(token) !== expected) throw new Error('Token 錯誤，拒絕存取');
}

function checkType_(t) {
  if (t !== 'expense' && t !== 'income') throw new Error('類型錯誤：' + t);
  return t;
}

function cleanRecord_(r) {
  if (!r) throw new Error('缺少記錄內容');
  const date = normDate_(String(r.date || ''));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日期格式錯誤：' + r.date);
  let time = String(r.time || '00:00').trim().slice(0, 5);
  if (/^\d:\d\d$/.test(time)) time = '0' + time;
  if (!/^\d{2}:\d{2}$/.test(time)) throw new Error('時間格式錯誤：' + r.time);
  const amount = Number(String(r.amount).replace(/,/g, ''));
  if (!isFinite(amount) || amount < 0) throw new Error('金額錯誤：' + r.amount);
  const currency = String(r.currency || 'TWD').trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(currency)) throw new Error('幣別格式錯誤：' + r.currency);
  return {
    id: str_(r.id, 64) || Utilities.getUuid(),
    date: date,
    time: time,
    category: str_(r.category, 30) || '其他',
    item: str_(r.item, 200),
    amount: Math.round(amount * 100) / 100,
    currency: currency,
    note: str_(r.note, 500),
    createdAt: str_(r.createdAt, 40) || new Date().toISOString(),
  };
}

/** 2026/1/5、2026.01.05 → 2026-01-05 */
function normDate_(s) {
  const m = String(s).trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!m) return String(s).trim();
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }

/** 防止以 = + - @ 開頭的文字被當成公式執行 */
function safe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function unsafe_(s) { return /^'[=+\-@]/.test(s) ? s.slice(1) : s; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
