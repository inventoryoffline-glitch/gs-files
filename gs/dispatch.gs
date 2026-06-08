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
      // Pre-set number format to string for the date column BEFORE setting values
      const dateColIdx = recHeaders.indexOf('date');
      if (dateColIdx >= 0) {
        recordSheet.getRange(startRow, dateColIdx + 1, newRecordRows.length, 1).setNumberFormat('@');
      }
      const range = recordSheet.getRange(startRow, 1, newRecordRows.length, newRecordRows[0].length);
      range.setValues(newRecordRows);
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
            if (canFill > 0) {
              const take = Math.min(remaining, canFill);
              allocSheet.getRange(i + 1, aDispIdx + 1).setValue(dispatched + take);
              remaining -= take;
            }
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

    const recData    = recordSheet.getDataRange().getValues();
    const recHeaders = recData[0];
    const vIdx       = recHeaders.indexOf("voucherNo");
    const stIdx      = recHeaders.indexOf("status");
    const itemIdIdx  = recHeaders.indexOf("itemId");
    const qtyIdx     = recHeaders.indexOf("qty");
    const instIdIdx  = recHeaders.indexOf("instId");

    const stockData  = stockSheet.getDataRange().getValues();
    const allocData  = allocSheet ? allocSheet.getDataRange().getValues() : [[]];
    const aHeaders   = allocData[0] || [];
    const aInstIdx   = aHeaders.indexOf("instituteId");
    const aItemIdx   = aHeaders.indexOf("itemId");
    const aDispIdx   = aHeaders.indexOf("dispatchedQty");

    let foundAny = false;
    for (let i = 1; i < recData.length; i++) {
      if (recData[i][vIdx].toString() === p.voucherNo && (recData[i][stIdx]||"active") === "active") {
        const itemId = recData[i][itemIdIdx];
        const qty    = parseInt(recData[i][qtyIdx]) || 0;
        const instId = recData[i][instIdIdx];

        // 1. Revert Stock
        for (let j = 1; j < stockData.length; j++) {
          if (stockData[j][0].toString() === itemId.toString()) {
            const curQty  = parseInt(stockData[j][4]) || 0;
            const curDisp = parseInt(stockData[j][5]) || 0;
            stockSheet.getRange(j + 1, 5, 1, 2).setValues([[curQty + qty, Math.max(0, curDisp - qty)]]);
            break;
          }
        }

        // 2. Revert Allocation
        if (allocData.length > 1 && aInstIdx >= 0) {
          for (let k = 1; k < allocData.length; k++) {
            if (allocData[k][aInstIdx].toString() === instId.toString() &&
                allocData[k][aItemIdx].toString() === itemId.toString()) {
              const curDisp = parseInt(allocData[k][aDispIdx]) || 0;
              allocSheet.getRange(k + 1, aDispIdx + 1).setValue(Math.max(0, curDisp - qty));
              // Don't break, might be multiple allocations (though unlikely in same voucher)
            }
          }
        }

        // 3. Mark Record as Cancelled
        recordSheet.getRange(i + 1, stIdx + 1).setValue("cancelled");
        foundAny = true;
      }
    }

    if (!foundAny) throw new Error("No active records found for voucher: " + p.voucherNo);
    return { success: true };
  });
}
