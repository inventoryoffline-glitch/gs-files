// ============================================================
// stock.gs — Stock CRUD, Settings, Institutes, Remove Stock
// ============================================================

// ── Settings ──────────────────────────────────────────────────
function getSettings() {
  const data = getSheetData(SHEET_SETTINGS);
  const obj  = {};
  data.forEach(r => { obj[r.key] = r.value; });
  return obj;
}

function saveSetting(key, value) {
  return withLock(() => {
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i+1, 2).setValue(value);
        return { success: true };
      }
    }
    sheet.appendRow([key, value]);
    return { success: true };
  });
}

// ── Stock ─────────────────────────────────────────────────────
function addStock(p) {
  return withLock(() => {
    validateRequired(p, ["name"]);
    const qty = validatePositiveInt(p.qty || 0, "Quantity");
    const low = validatePositiveInt(p.low || 10, "Low stock threshold");
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOCK);
    const id    = Date.now().toString();
    sheet.appendRow([id, p.name.trim(), p.cat||"", p.unit||"", qty, 0, low, p.expiry||"", p.showLow||"true", p.serialType||"", p.serialValue||"", p.refNumber||""]);
    return { success: true, id };
  });
}

function updateStock(p) {
  return withLock(() => {
    validateRequired(p, ["id"]);
    const addQty = validatePositiveInt(p.addQty || 0, "Quantity to add");
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOCK);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === p.id.toString()) {
        const cur = parseInt(data[i][4]) || 0;
        sheet.getRange(i+1, 5).setValue(cur + addQty);
        return { success: true };
      }
    }
    return { error: "Item not found" };
  });
}

function updateShowLow(p) {
  return withLock(() => {
    validateRequired(p, ["id"]);
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOCK);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === p.id.toString()) {
        sheet.getRange(i+1, 9).setValue(p.showLow);
        return { success: true };
      }
    }
    return { error: "Item not found" };
  });
}

function removeStock(p) {
  return withLock(() => {
    validateRequired(p, ["itemId","date","reason"]);
    const qty = validatePositiveInt(p.qty, "Quantity");
    clearCache();
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet = ss.getSheetByName(SHEET_STOCK);
    const adjSheet   = ss.getSheetByName(SHEET_ADJUSTMENTS);
    const stockData  = stockSheet.getDataRange().getValues();
    for (let i = 1; i < stockData.length; i++) {
      if (stockData[i][0].toString() === p.itemId.toString()) {
        const curQty   = parseInt(stockData[i][4]) || 0;
        if (qty > curQty) throw new Error("Removal quantity exceeds available stock.");
        const newQty   = curQty - qty;
        const itemName = stockData[i][1];
        const unit     = stockData[i][3];
        stockSheet.getRange(i+1, 5).setValue(newQty);
        adjSheet.appendRow([Date.now().toString(), safeDateStr(p.date), p.itemId, itemName, unit, qty, p.reason||"Other", p.remarks||"", newQty]);
        return { success: true, stockAfter: newQty };
      }
    }
    return { error: "Item not found" };
  });
}

// ── Institutes ────────────────────────────────────────────────
function addInstitute(p) {
  return withLock(() => {
    validateRequired(p, ["name"]);
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSTITUTES);
    const id    = Date.now().toString();
    sheet.appendRow([id, p.name.trim(), p.category||"", p.contact||"", p.phone||"", p.address||""]);
    return { success: true, id };
  });
}
