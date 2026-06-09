// ============================================================
// notes.gs — Institute Notes Management
// ============================================================

function saveInstituteNote(p) {
  return withLock(() => {
    validateRequired(p, ["instituteId", "noteText"]);
    clearCache();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const notesSheet = ss.getSheetByName(SHEET_NOTES);
    if (!notesSheet) throw new Error("Notes sheet not found. Please run initSheets() first.");

    const noteId = "note_" + Date.now() + "_" + Math.random().toString(36).substr(2, 9);
    const timestamp = new Date().toISOString();
    const noteText = (p.noteText || "").toString().trim();
    const instituteId = p.instituteId.toString();

    if (!noteText) throw new Error("Note text cannot be empty.");

    notesSheet.appendRow([
      noteId,
      instituteId,
      p.instituteName || "",
      noteText,
      timestamp,
      "active"
    ]);

    return { success: true, noteId, timestamp };
  });
}

function getInstituteNotes(p) {
  validateRequired(p, ["instituteId"]);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const notesSheet = ss.getSheetByName(SHEET_NOTES);
  if (!notesSheet) return { notes: [] };

  const data = notesSheet.getDataRange().getValues();
  const headers = data[0];
  const instIdIdx = headers.indexOf("instituteId");
  const statusIdx = headers.indexOf("status");

  const notes = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][instIdIdx].toString() === p.instituteId.toString() && 
        (data[i][statusIdx] || "active") === "active") {
      notes.push({
        id: data[i][0],
        instituteId: data[i][1],
        instituteName: data[i][2],
        noteText: data[i][3],
        timestamp: data[i][4],
        status: data[i][5]
      });
    }
  }

  // Sort by timestamp descending (newest first)
  notes.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return { notes };
}

function deleteInstituteNote(p) {
  return withLock(() => {
    validateRequired(p, ["noteId"]);
    clearCache();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const notesSheet = ss.getSheetByName(SHEET_NOTES);
    if (!notesSheet) throw new Error("Notes sheet not found.");

    const data = notesSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === p.noteId.toString()) {
        // Mark as deleted instead of removing row
        notesSheet.getRange(i + 1, 6).setValue("deleted");
        return { success: true };
      }
    }

    throw new Error("Note not found.");
  });
}

function getAllInstituteNotes(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const notesSheet = ss.getSheetByName(SHEET_NOTES);
  if (!notesSheet) return { notes: [] };

  const data = notesSheet.getDataRange().getValues();
  if (data.length <= 1) return { notes: [] };

  const headers = data[0];
  const statusIdx = headers.indexOf("status");

  const notes = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][statusIdx] || "active") === "active") {
      notes.push({
        id: data[i][0],
        instituteId: data[i][1],
        instituteName: data[i][2],
        noteText: data[i][3],
        timestamp: data[i][4],
        status: data[i][5]
      });
    }
  }

  notes.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return { notes };
}
