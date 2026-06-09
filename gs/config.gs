// ============================================================
// config.gs — Constants, Entry Points, Request Router
// ============================================================

const SHEET_STOCK       = "Stock";
const SHEET_INSTITUTES  = "Institutes";
const SHEET_RECORDS     = "Records";
const SHEET_SETTINGS    = "Settings";
const SHEET_ADJUSTMENTS = "Adjustments";
const SHEET_TEMPLATES   = "DistributionTemplates";
const SHEET_ALLOCATIONS = "TemplateAllocations";
const SHEET_NOTES       = "InstituteNotes";
const SHEET_LOGS        = "Logs";
const CACHE_KEY         = "inv_all_data_v8";
const CACHE_DURATION    = 300; // 5 minutes

// ── Entry Points ──────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const p      = (e && e.parameter) ? e.parameter : {};
  const action = p.action || "unknown";
  let result;

  try {
    initSheets();

    if      (action === "getAll")                result = getAllData();
    else if (action === "addStock")              result = addStock(p);
    else if (action === "updateStock")           result = updateStock(p);
    else if (action === "updateShowLow")         result = updateShowLow(p);
    else if (action === "deleteStock")           result = deleteRow(SHEET_STOCK, p.id);
    else if (action === "removeStock")           result = removeStock(p);
    else if (action === "addInstitute")          result = addInstitute(p);
    else if (action === "deleteInstitute")       result = deleteRow(SHEET_INSTITUTES, p.id);
    else if (action === "dispatch")              result = dispatch(p);
    else if (action === "cancelDispatch")        result = cancelDispatch(p);
    else if (action === "saveSetting")           result = saveSetting(p.key, p.value);
    else if (action === "getSettings")           result = getSettings();
    else if (action === "createTemplate")        result = createTemplate(p);
    else if (action === "deleteTemplate")        result = deleteTemplate(p);
    else if (action === "updateTemplateStatus")  result = updateTemplateStatus(p);
    else if (action === "duplicateTemplate")     result = duplicateTemplate(p);
    else if (action === "saveAllocation")        result = saveAllocation(p);
    else if (action === "batchSaveAllocations")  result = batchSaveAllocations(p);
    else if (action === "batchDeleteAllocations")result = batchDeleteAllocations(p);
    else if (action === "deleteAllocation")      result = deleteRow(SHEET_ALLOCATIONS, p.allocationId);
    else if (action === "runAudit")              result = runServerAudit();
    else if (action === "updateSerialNotes")     result = updateSerialNotes(p);
    else if (action === "saveInstituteNote")      result = saveInstituteNote(p);
    else if (action === "getInstituteNotes")      result = getInstituteNotes(p);
    else if (action === "deleteInstituteNote")    result = deleteInstituteNote(p);
    else if (action === "getAllInstituteNotes")   result = getAllInstituteNotes(p);
    else result = { error: "Unknown action: " + action };

    writeLog(action, "success", "");

  } catch(err) {
    result = { error: err.toString() };
    writeLog(action, "error", err.toString());
  }

  const json = JSON.stringify(result);

  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
