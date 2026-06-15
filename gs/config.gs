// ============================================================
// config.gs — UPDATED: Added Batch Routes
// ============================================================

// ... (existing constants)

function handleRequest(e) {
  const p      = (e && e.parameter) ? e.parameter : {};
  const action = p.action || "unknown";
  let result;

  try {
    initSheets();

    if      (action === "getAll")                result = getAllData();
    else if (action === "addStock")              result = addStock(p);
    else if (action === "batchAddStock")         result = batchAddStock(p); // NEW
    else if (action === "updateStock")           result = updateStock(p);
    else if (action === "updateShowLow")         result = updateShowLow(p);
    else if (action === "deleteStock")           result = deleteRow(SHEET_STOCK, p.id);
    else if (action === "removeStock")           result = removeStock(p);
    else if (action === "addInstitute")          result = addInstitute(p);
    else if (action === "batchAddInstitutes")    result = batchAddInstitutes(p); // NEW
    else if (action === "deleteInstitute")       result = deleteRow(SHEET_INSTITUTES, p.id);
    else if (action === "dispatch")              result = dispatch(p);
    // ... (rest of the actions)
    else result = { error: "Unknown action: " + action };

    writeLog(action, "success", "");

  } catch(err) {
    result = { error: err.toString() };
    writeLog(action, "error", err.toString());
  }

  // ... (rest of the handleRequest function)
}
