// ============================================================
// helpers.gs — Init, Cache, Lock, Validation, Date Utils, Logging
// ============================================================

// ── Sheet Init ────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_STOCK,       ["id","name","cat","unit","qty","dispatched","low","expiry","showLow","serialType","serialValue","refNumber"]);
  ensureSheet(ss, SHEET_INSTITUTES,  ["id","name","category","contact","phone","address"]);
  ensureSheet(ss, SHEET_RECORDS,     ["id","voucherNo","date","instId","instName","itemId","itemName","qty","unit","remarks","receivedBy","receivedPhone","issuedBy","status","serialType","serialNotes"]);
  ensureSheet(ss, SHEET_SETTINGS,    ["key","value"]);
  ensureSheet(ss, SHEET_ADJUSTMENTS, ["id","date","itemId","itemName","unit","qty","reason","remarks","stockAfter"]);
  ensureSheet(ss, SHEET_TEMPLATES,   ["templateId","templateName","purpose","targetDate","status","createdAt"]);
  ensureSheet(ss, SHEET_ALLOCATIONS, ["allocationId","templateId","instituteId","instituteName","itemId","itemName","unit","plannedQty","dispatchedQty"]);
  ensureSheet(ss, SHEET_LOGS,        ["timestamp","action","status","detail"]);

  // Force date columns to plain text — prevents Sheets auto-converting YYYY-MM-DD to Date serial (causes day-shift in IST)
  [ [SHEET_RECORDS, 3],
    [SHEET_ADJUSTMENTS, 2],
    [SHEET_TEMPLATES, 6],
    [SHEET_STOCK, 8],
  ].forEach(([sheetName, col]) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, col, sh.getLastRow() - 1, 1).setNumberFormat('@');
  });

  // Seed default settings
  const sSheet  = ss.getSheetByName(SHEET_SETTINGS);
  const sData   = sSheet.getDataRange().getValues();
  const existing = {};
  sData.slice(1).forEach(r => { existing[r[0]] = true; });
  const defaults = { systemName:"Inventory Management System", issuedBy:"", nearExpiry:"90", highExpiry:"30" };
  const newRows  = [];
  for (const [k, v] of Object.entries(defaults)) {
    if (!existing[k]) newRows.push([k, v]);
  }
  if (newRows.length > 0) {
    sSheet.getRange(sSheet.getLastRow()+1, 1, newRows.length, 2).setValues(newRows);
  }
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#2d5016")
      .setFontColor("white");
  }
  return sheet;
}

// ── Cache ─────────────────────────────────────────────────────
function getCached() {
  try {
    const cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}
  return null;
}

function setCache(data) {
  try {
    const str = JSON.stringify(data);
    if (str.length < 90000) {
      CacheService.getScriptCache().put(CACHE_KEY, str, CACHE_DURATION);
    }
  } catch(e) {}
}

function clearCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

// ── Lock ──────────────────────────────────────────────────────
function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return fn();
  } catch(e) {
    if (e.message && e.message.includes("lock")) {
      throw new Error("System is busy. Please try again in a moment.");
    }
    throw e;
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ── Validation ────────────────────────────────────────────────
function validateRequired(obj, fields) {
  const missing = fields.filter(f => !obj[f] || obj[f].toString().trim() === "");
  if (missing.length > 0) throw new Error("Missing required fields: " + missing.join(", "));
}

function validatePositiveInt(val, name) {
  const n = parseInt(val);
  if (isNaN(n) || n < 0) throw new Error(name + " must be a positive number.");
  return n;
}

// ── Date Utilities ────────────────────────────────────────────
// Safe parse: splits "YYYY-MM-DD" to avoid UTC timezone shift in IST and other UTC+ zones
function safeParseDate(str) {
  if (!str) return new Date();
  const parts = str.split('-');
  if (parts.length !== 3) return new Date(str);
  const d = new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]));
  d.setHours(0,0,0,0);
  return d;
}

// Format Date object to "YYYY-MM-DD" using local date methods (not UTC)
function fmtDateGS(d) {
  const dt = d instanceof Date ? d : new Date(d);
  return dt.getFullYear() + '-' +
    ('0'+(dt.getMonth()+1)).slice(-2) + '-' +
    ('0'+dt.getDate()).slice(-2);
}

// Returns YYYY-MM-DD string safe for sheet storage (won't be auto-converted)
function safeDateStr(d) {
  if (!d) return '';
  if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  if (d instanceof Date) return fmtDateGS(d);
  return String(d);
}

// ── Logging ───────────────────────────────────────────────────
function writeLog(action, status, detail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGS);
    if (!sheet) return;
    if (status === "success" && action === "getAll") return;
    sheet.appendRow([new Date().toISOString(), action, status, detail.toString().slice(0, 500)]);
    const rows = sheet.getLastRow();
    if (rows > 501) sheet.deleteRows(2, rows - 501);
  } catch(e) {}
}

// ── Data Access ───────────────────────────────────────────────
function getAllData() {
  const cached = getCached();
  if (cached) return cached;
  const data = {
    stock:         getSheetData(SHEET_STOCK),
    institutes:    getSheetData(SHEET_INSTITUTES),
    records:       getSheetData(SHEET_RECORDS),
    adjustments:   getSheetData(SHEET_ADJUSTMENTS),
    distTemplates: getSheetData(SHEET_TEMPLATES),
    distAllocations: getSheetData(SHEET_ALLOCATIONS),
    settings:      getSettings()
  };
  setCache(data);
  return data;
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      // Convert Date objects to YYYY-MM-DD string using local methods - never JSON.stringify a Date
      if (v instanceof Date && !isNaN(v.getTime())) {
        obj[h] = fmtDateGS(v);
      } else {
        obj[h] = v;
      }
    });
    return obj;
  });
}

function deleteRow(sheetName, id) {
  return withLock(() => {
    if (!id) throw new Error("ID is required for delete.");
    clearCache();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { error: "Row not found" };
  });
}
