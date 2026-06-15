// ============================================================
// stock.gs — UPDATED: Added Batch Processing
// ============================================================

// ... (existing getSettings and saveSetting functions)

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

/**
 * BATCH ADD STOCK: Processes multiple items in a single lock session
 */
function batchAddStock(p) {
  return withLock(() => {
    if (!p.itemsJson) throw new Error("No data provided");
    const items = JSON.parse(p.itemsJson);
    if (!Array.isArray(items) || items.length === 0) throw new Error("Invalid data format");

    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOCK);
    const timestamp = Date.now();
    
    const rows = items.map((item, index) => {
      const id = (timestamp + index).toString();
      const qty = parseInt(item.qty) || 0;
      const low = parseInt(item.low) || 10;
      return [
        id, 
        (item.name || "Unknown").trim(), 
        item.cat || "", 
        item.unit || "", 
        qty, 
        0, 
        low, 
        item.expiry || "", 
        item.showLow || "true",
        item.serialType || "",
        item.serialValue || "",
        item.refNumber || ""
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: rows.length };
  });
}

// ... (existing updateStock, updateShowLow, removeStock functions)

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

/**
 * BATCH ADD INSTITUTES
 */
function batchAddInstitutes(p) {
  return withLock(() => {
    if (!p.itemsJson) throw new Error("No data provided");
    const items = JSON.parse(p.itemsJson);
    if (!Array.isArray(items) || items.length === 0) throw new Error("Invalid data format");

    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSTITUTES);
    const timestamp = Date.now();
    
    const rows = items.map((inst, index) => {
      const id = (timestamp + index).toString();
      return [
        id, 
        (inst.name || "Unknown").trim(), 
        inst.category || "", 
        inst.contact || "", 
        inst.phone || "", 
        inst.address || ""
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: rows.length };
  });
}
