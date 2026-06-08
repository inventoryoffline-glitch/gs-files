// ============================================================
// dispatch.gs — Dispatch & Cancel Dispatch
// ============================================================

function dispatch(p) {
  return withLock(() => {
    validateRequired(p, ["voucherNo","instId","instName","date","receivedBy","items"]);
    clearCache();

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet  = ss.getSheetByName(SHEET_STOCK);
    const recordSheet = ss.getSheetByName(SHEET_RECORDS);
    const allocSheet  = ss.getSheetByName(SHEET_ALLOCATIONS);
    const items       = JSON.parse(p.items);

    if (!items || items.length === 0) throw new Error("No items provided for dispatch.");

    // Check duplicate active voucher
    const recData    = recordSheet.getDataRange().getValues();
    const recHeaders = recData[0];
    const vIdx  = recHeaders.indexOf("voucherNo");
    const stIdx = recHeaders.indexOf("status");
    for (let i = 1; i < recData.length; i++) {
      if (recData[i][vIdx].toString() === p.voucherNo &&
          (recData[i][stIdx]||"active").toString() === "active") {
        throw new Error("Voucher number already exists as active: " + p.voucherNo);
      }
    }

    const stockData = stockSheet.getDataRange().getValues();
    const allocData = allocSheet ? allocSheet.getDataRange().getValues() : [[]];
    const aHeaders  = allocData[0] || [];
    const aInstIdx  = aHeaders.indexOf("instituteId");
    const aItemIdx  = aHeaders.indexOf("itemId");
    const aDispIdx  = aHeaders.indexOf("dispatchedQty");
    const aPlannedIdx = aHeaders.indexOf("plannedQty");

    // Safe date parse — splits YYYY-MM-DD to avoid UTC timezone shift in IST
    const dispDate = safeParseDate(p.date);
    const today2   = new Date(); today2.setHours(0,0,0,0);
    if (dispDate.getTime() > today2.getTime()) {
      throw new Error("Dispatch date cannot be in the future.");
    }

    // Validate all items first — nothing written until all pass
    const stockUpdates = [];
    for (const item of items) {
      const qty = validatePositiveInt(item.qty, "Item quantity");
      let found = false;
      for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][0].toString() === item.itemId.toString()) {
          // Check expiry using safe local parse
          const expiryVal = stockData[i][7] ? stockData[i][7].toString().trim() : "";
          if (expiryVal && /^\d{4}-\d{2}-\d{2}$/.test(expiryVal)) {
            const expDate  = safeParseDate(expiryVal);
            const todayExp = new Date(); todayExp.setHours(0,0,0,0);
            if (expDate.getTime() < todayExp.getTime()) {
              throw new Error("Cannot dispatch expired item: " + item.itemName);
            }
          }
          const curQty = parseInt(stockData[i][4]) || 0;
          if (qty > curQty) throw new Error("Insufficient stock for: " + item.itemName + ". Available: " + curQty);
          const newQty  = curQty - qty;
          const newDisp = (parseInt(stockData[i][5]) || 0) + qty;
          stockUpdates.push({rowIndex: i, newQty, newDisp, item, qty});
          stockData[i][4] = newQty;
          stockData[i][5] = newDisp;
          found = true;
          break;
        }
      }
      if (!found) throw new Error("Stock item not found: " + item.itemName);
    }

    // Write stock updates
    stockUpdates.forEach(u => {
      stockSheet.getRange(u.rowIndex + 1, 5, 1, 2).setValues([[u.newQty, u.newDisp]]);
    });

    // Build and write record rows in one batch
    const baseTs = Date.now().toString();
    const newRecordRows = stockUpdates.map((u, idx) => [
      baseTs + '_' + idx,
      p.voucherNo, safeDateStr(p.date), p.instId, p.instName,
      u.item.itemId, u.item.itemName, u.qty, u.item.unit||"",
      p.remarks||"", p.receivedBy||"", p.receivedPhone||"",
      p.issuedBy||"", "active", u.item.serialType||"", u.item.serialNotes||""
    ]);

    if (newRecordRows.length > 0) {
      const startRow = recordSheet.getLastRow() + 1;
      const range    = recordSheet.getRange(startRow, 1, newRecordRows.length, newRecordRows[0].length);
      range.setValues(newRecordRows);
      // Force date column to text to prevent Sheets auto-converting
      const dateColIdx = recHeaders.indexOf('date');
      if (dateColIdx >= 0) {
        recordSheet.getRange(startRow, dateColIdx + 1, newRecordRows.length, 1).setNumberFormat('@STRING@');
      }
    }

    // Update allocation dispatchedQty (FIFO)
    if (allocData.length > 1 && aInstIdx >= 0) {
      stockUpdates.forEach(u => {
        let remaining = u.qty;
        for (let i = 1; i < allocData.length && remaining > 0; i++) {
          if ((allocData[i][aInstIdx]||'').toString() === p.instId.toString() &&
              (allocData[i][aItemIdx]||'').toString() === u.item.itemId.toString()) {
            const planned    = parseInt(allocData[i][aPlannedIdx])||0;
            const dispatched = parseInt(allocData[i][aDispIdx])||0;
            const canFill    = planned - dispatched;
            if (canFill <= 0) continue;
            const fill = Math.min(canFill, remaining);
            allocSheet.getRange(i+1, aDispIdx+1).setValue(dispatched + fill);
            allocData[i][aDispIdx] = dispatched + fill;
            remaining -= fill;
          }
        }
      });
    }

    return { success: true, voucherNo: p.voucherNo };
  });
}

function cancelDispatch(p) {
  return withLock(() => {
    validateRequired(p, ["voucherNo"]);
    clearCache();

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet  = ss.getSheetByName(SHEET_STOCK);
    const recordSheet = ss.getSheetByName(SHEET_RECORDS);
    const allocSheet  = ss.getSheetByName(SHEET_ALLOCATIONS);
    const recData     = recordSheet.getDataRange().getValues();
    const headers     = recData[0];
    const vIdx    = headers.indexOf("voucherNo");
    const iIdx    = headers.indexOf("itemId");
    const instIdx = headers.indexOf("instId");
    const qIdx    = headers.indexOf("qty");
    const stIdx   = headers.indexOf("status");
    const stockData = stockSheet.getDataRange().getValues();
    const allocData = allocSheet ? allocSheet.getDataRange().getValues() : [[]];
    const aHeaders  = allocData[0]||[];
    const aInstIdx  = aHeaders.indexOf("instituteId");
    const aItemIdx  = aHeaders.indexOf("itemId");
    const aDispIdx  = aHeaders.indexOf("dispatchedQty");

    let found = false;
    for (let i = 1; i < recData.length; i++) {
      if (recData[i][vIdx].toString() === p.voucherNo &&
          recData[i][stIdx].toString() !== "cancelled") {
        const itemId = recData[i][iIdx].toString();
        const instId = recData[i][instIdx] ? recData[i][instIdx].toString() : '';
        const qty    = parseInt(recData[i][qIdx])||0;

        recordSheet.getRange(i+1, stIdx+1).setValue("cancelled");

        // Restore stock
        for (let j = 1; j < stockData.length; j++) {
          if (stockData[j][0].toString() === itemId) {
            const curQty  = parseInt(stockData[j][4])||0;
            const curDisp = parseInt(stockData[j][5])||0;
            stockSheet.getRange(j+1, 5).setValue(curQty + qty);
            stockSheet.getRange(j+1, 6).setValue(Math.max(0, curDisp - qty));
            stockData[j][4] = curQty + qty;
            stockData[j][5] = Math.max(0, curDisp - qty);
            break;
          }
        }

        // Restore allocation dispatchedQty (reverse FIFO)
        if (allocData.length > 1 && aInstIdx >= 0 && instId) {
          let toRestore = qty;
          for (let k = allocData.length-1; k >= 1 && toRestore > 0; k--) {
            if ((allocData[k][aInstIdx]||'').toString() === instId &&
                (allocData[k][aItemIdx]||'').toString() === itemId) {
              const dispatched = parseInt(allocData[k][aDispIdx])||0;
              const restore    = Math.min(dispatched, toRestore);
              allocSheet.getRange(k+1, aDispIdx+1).setValue(dispatched - restore);
              allocData[k][aDispIdx] = dispatched - restore;
              toRestore -= restore;
            }
          }
        }
        found = true;
      }
    }
    return found ? { success: true } : { error: "Voucher not found or already cancelled" };
  });
}

function updateSerialNotes(p) {
  return withLock(() => {
    if (!p.recordId) throw new Error("Record ID required.");
    clearCache();
    const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx   = headers.indexOf("id");
    const snIdx   = headers.indexOf("serialNotes");
    if (snIdx < 0) throw new Error("serialNotes column not found.");
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx].toString() === p.recordId.toString()) {
        sheet.getRange(i+1, snIdx+1).setValue(p.serialNotes||"");
        return { success: true };
      }
    }
    return { error: "Record not found" };
  });
}
