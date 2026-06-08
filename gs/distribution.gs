// ============================================================
// distribution.gs — Templates, Allocations, Batch Operations
// ============================================================

function createTemplate(p) {
  return withLock(() => {
    validateRequired(p, ["templateName"]);
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEMPLATES);
    const id    = Date.now().toString();
    sheet.appendRow([id, p.templateName.trim(), p.purpose||"", p.targetDate||"", p.status||"Draft", new Date().toISOString()]);
    return { success: true, templateId: id };
  });
}

function deleteTemplate(p) {
  return withLock(() => {
    validateRequired(p, ["templateId"]);
    clearCache();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const tSheet = ss.getSheetByName(SHEET_TEMPLATES);
    const tData  = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (tData[i][0].toString() === p.templateId.toString()) {
        tSheet.deleteRow(i+1); break;
      }
    }

    const aSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
    const aData  = aSheet.getDataRange().getValues();
    const tIdx   = aData[0].indexOf("templateId");
    for (let i = aData.length-1; i >= 1; i--) {
      if (aData[i][tIdx].toString() === p.templateId.toString()) {
        aSheet.deleteRow(i+1);
      }
    }
    return { success: true };
  });
}

function updateTemplateStatus(p) {
  return withLock(() => {
    validateRequired(p, ["templateId","status"]);
    clearCache();
    const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEMPLATES);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx   = headers.indexOf("templateId");
    const stIdx   = headers.indexOf("status");
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx].toString() === p.templateId.toString()) {
        sheet.getRange(i+1, stIdx+1).setValue(p.status);
        return { success: true };
      }
    }
    return { error: "Template not found" };
  });
}

function duplicateTemplate(p) {
  return withLock(() => {
    validateRequired(p, ["templateId"]);
    clearCache();
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName(SHEET_TEMPLATES);
    const aSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
    const tData  = tSheet.getDataRange().getValues();
    const headers= tData[0];
    const idIdx  = headers.indexOf("templateId");
    const nameIdx= headers.indexOf("templateName");
    const stIdx  = headers.indexOf("status");
    const crIdx  = headers.indexOf("createdAt");

    let origRow = null;
    for (let i = 1; i < tData.length; i++) {
      if (tData[i][idIdx].toString() === p.templateId.toString()) { origRow = [...tData[i]]; break; }
    }
    if (!origRow) return { error: "Template not found" };

    const newId       = Date.now().toString();
    origRow[idIdx]    = newId;
    origRow[nameIdx]  = p.newName || "Copy of " + origRow[nameIdx];
    origRow[stIdx]    = "Draft";
    origRow[crIdx]    = new Date().toISOString();
    tSheet.appendRow(origRow);

    const aData = aSheet.getDataRange().getValues();
    if (aData.length > 1) {
      const aHeaders = aData[0];
      const aTIdx    = aHeaders.indexOf("templateId");
      const aIdIdx   = aHeaders.indexOf("allocationId");
      const dIdx     = aHeaders.indexOf("dispatchedQty");
      const newRows  = [];
      for (let i = 1; i < aData.length; i++) {
        if (aData[i][aTIdx].toString() === p.templateId.toString()) {
          const newRow  = [...aData[i]];
          newRow[aIdIdx]= Date.now().toString() + i;
          newRow[aTIdx] = newId;
          newRow[dIdx]  = 0;
          newRows.push(newRow);
        }
      }
      if (newRows.length > 0) {
        aSheet.getRange(aSheet.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
    }
    return { success: true, newTemplateId: newId };
  });
}

function saveAllocation(p) {
  return withLock(() => {
    validateRequired(p, ["templateId","instituteId","itemId","plannedQty"]);
    const qty = validatePositiveInt(p.plannedQty, "Planned quantity");
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ALLOCATIONS);
    const id    = Date.now().toString();
    sheet.appendRow([id, p.templateId, p.instituteId, p.instituteName||"", p.itemId, p.itemName||"", p.unit||"", qty, 0]);
    return { success: true, allocationId: id };
  });
}

function batchSaveAllocations(p) {
  return withLock(() => {
    if (!p.templateId)  throw new Error("templateId required");
    if (!p.allocations) throw new Error("allocations JSON required");

    let rows;
    try { rows = JSON.parse(p.allocations); }
    catch(e) { throw new Error("Invalid allocations JSON: " + e.message); }

    clearCache();
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName(SHEET_ALLOCATIONS);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const tIdx      = headers.indexOf("templateId");
    const aIdIdx    = headers.indexOf("allocationId");
    const iInstIdx  = headers.indexOf("instituteId");
    const iItemIdx  = headers.indexOf("itemId");
    const pQtyIdx   = headers.indexOf("plannedQty");
    const dQtyIdx   = headers.indexOf("dispatchedQty");
    const instNmIdx = headers.indexOf("instituteName");
    const itemNmIdx = headers.indexOf("itemName");
    const unitIdx   = headers.indexOf("unit");

    // Delete existing allocations for this template
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][tIdx] && data[i][tIdx].toString() === p.templateId.toString()) {
        sheet.deleteRow(i + 1);
      }
    }

    // Batch write new rows
    if (rows.length > 0) {
      const newRows = rows.map((r, idx) => {
        const row       = new Array(headers.length).fill("");
        row[aIdIdx]     = Date.now().toString() + idx;
        row[tIdx]       = p.templateId;
        row[iInstIdx]   = r.instituteId    || "";
        row[instNmIdx]  = r.instituteName  || "";
        row[iItemIdx]   = r.itemId         || "";
        row[itemNmIdx]  = r.itemName       || "";
        row[unitIdx]    = r.unit           || "";
        row[pQtyIdx]    = parseInt(r.plannedQty)    || 0;
        row[dQtyIdx]    = parseInt(r.dispatchedQty) || 0;
        return row;
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    }

    return { success: true, saved: rows.length };
  });
}

function batchDeleteAllocations(p) {
  return withLock(() => {
    if (!p.templateId) throw new Error("templateId required");
    clearCache();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ALLOCATIONS);
    const data  = sheet.getDataRange().getValues();
    const tIdx  = data[0].indexOf("templateId");
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][tIdx] && data[i][tIdx].toString() === p.templateId.toString()) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  });
}
