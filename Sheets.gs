/**
 * 📁 Sheets.gs
 * ─────────────────────────────────────────────
 * Utility functions for managing the Learning Log and Master Lesson Plan
 * in the Japanese Learning Tracker system.
 *
 * 🧰 Functions:
 * - ensureKanjiVocabColumn() → Adds the Kanji/Vocab column to Master Sheet if missing
 * - logStudyEntry() → Launches the Kanji Logger sidebar
 * - saveStudySession() → Inserts a structured study session into the current phase log
 * - writeSessionAnchorClean() → Writes and locks a clean string-formatted session date
 * - getFirstEmptyRowInColumn() → Finds the first truly empty and unformatted row
 * - cleanGhostRowsFromLearningLog() → Deletes ghost rows beyond last meaningful entry
 * - reapplyFormattingToLearningLog() → Reapplies checkbox and validation formatting
 * - getPhaseStartDate() → Retrieves the start date of the current phase safely
 * - getRealLastRow() → Gets the true last used row in a column, ignoring ghost space
 * - hardResetLearningLog() → Trims and refills the Learning Log to 1000 rows
 * - getPhaseLogSheet() → Returns the correct phase-specific log sheet (or creates it)
 * - clearTasksForCurrentPhase() → Deletes Google Tasks matching current phase schedule
 *
 * 📊 Sheet Dependencies:
 * - Dashboard (current phase in B3, start date in B4)
 * - Master Lesson Plan (Phase and Day layout)
 * - Learning Log (per-phase or default)
 *
 * ⚙️ Other Dependencies:
 * - THEMES object (for color settings)
 * - getSheetIdFromDashboard() for active spreadsheet context
 */

// ─────────────────────────────────────────────

/**
 * 🧱 Ensure Kanji/Vocab column exists in Master Sheet
 */
function ensureKanjiVocabColumn() {
  const spreadsheet = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!headers.includes(KANJI_VOCAB_COLUMN)) {
    sheet.getRange(1, headers.length + 1).setValue(KANJI_VOCAB_COLUMN);
    Logger.log(`✅ Added '${KANJI_VOCAB_COLUMN}' column to Master Sheet`);
  }
}


/**
 * 📤 Launch Kanji Logger Sidebar
 */
function logStudyEntry() {
  openKanjiLoggerSidebar();
}


/**
 * 🗓 Safely parse date from cell value
 *
 * @param {*} value - Raw date value
 * @param {number} rowIndex - Row for logging
 * @returns {Date|null}
 */
function parseDateSafely(value, rowIndex) {
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    Logger.log(`❌ Invalid date at row ${rowIndex}: ${value}`);
    return null;
  }
  return parsed;
}


/**
 * 📚 Log a full study session entry
 *
 * @param {Object} data - Structured study entry object
 */
function saveStudySession(data) {
  const sheet = getPhaseLogSheet();
  if (!sheet) {
    Logger.log("❌ Learning Log sheet not found.");
    return;
  }

  cleanGhostRowsFromLearningLog();
  const firstEmptyRow = getFirstEmptyRowInColumn(sheet, 1);
  const formattedDate = new Date(data.timestamp);

  const values = [
    formattedDate,
    data.type || "",
    data.content || "",
    data.tags || "",
    data.day || "",
    "",
    data.time || "",
    data.completed || false,
    data.difficulty || "",
    data.notes || ""
  ];

  sheet.getRange(firstEmptyRow, 1, 1, values.length).setValues([values]);
  writeSessionAnchorClean(firstEmptyRow, data.date || "");

  const lastFormattedRow = sheet.getLastRow();
  if (firstEmptyRow > lastFormattedRow) {
    reapplyFormattingToLearningLog(firstEmptyRow);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Study Session Logged!\n(Session Anchor stored as string)", "Logger", 5);
}


/**
 * 🔐 Write session anchor as locked plain string
 *
 * @param {number} row - Target row
 * @param {string} value - Date string to store
 */
function writeSessionAnchorClean(row, value) {
  const sheet = getPhaseLogSheet();
  const col = 6;
  const cell = sheet.getRange(row, col);
  cell.clearContent();
  cell.setNumberFormat("@STRING@");
  cell.setValue(value);
  Logger.log(`🧼 Locked Session Anchor at row ${row} as string: ${value}`);
}


/**
 * 🔎 Find first truly empty row in column
 *
 * @param {Sheet} sheet - Target sheet
 * @param {number} col - Column to scan
 * @returns {number}
 */
function getFirstEmptyRowInColumn(sheet, col, maxRows) {
  const lastRow = sheet.getLastRow();
  const rowCount = Math.min(maxRows || lastRow + 1, sheet.getMaxRows());
  const range = sheet.getRange(1, col, rowCount);
  const values = range.getValues();
  const validations = range.getDataValidations();
  const notes = range.getNotes();

  for (let i = 0; i < rowCount; i++) {
    const cell = range.getCell(i + 1, 1);
    const isBlank = !values[i][0];
    const isUnformatted = !validations[i][0] && !cell.isChecked() && !notes[i][0];
    if (isBlank && isUnformatted) return i + 1;
  }
  return rowCount + 1;
}


/**
 * 🧹 Clean ghost rows from Learning Log
 */
function cleanGhostRowsFromLearningLog() {
  const sheet = getPhaseLogSheet();
  const data = sheet.getDataRange().getValues();
  let lastValidRow = data.length;

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i].some(cell => cell !== "")) {
      lastValidRow = i + 1;
      break;
    }
  }

  const maxRows = sheet.getMaxRows();
  const ghostCount = maxRows - lastValidRow;
  if (ghostCount > 0) {
    sheet.deleteRows(lastValidRow + 1, ghostCount);
    Logger.log(`🧹 Deleted ${ghostCount} ghost rows.`);
  }
}


/**
 * 🎨 Reapply checkbox and validation formatting
 *
 * @param {number} startRow - Starting row (default = 2)
 */
function reapplyFormattingToLearningLog(startRow = 2) {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getMaxRows() < startRow) return;

  const rowCount = sheet.getMaxRows() - startRow + 1;

  // ➤ Reinsert checkboxes in Completed column
  const checkboxRange = sheet.getRange(startRow, 8, rowCount);
  checkboxRange.insertCheckboxes();

  // ➤ Reapply difficulty level dropdown
  const difficultyRange = sheet.getRange(startRow, 9, rowCount);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["1", "2", "3", "4", "5"], true)
    .setAllowInvalid(false)
    .build();
  difficultyRange.setDataValidation(rule);
}


/**
 * 🗓 Get the current phase's start date from Dashboard
 *
 * @returns {Date}
 */
function getPhaseStartDate() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(DASHBOARD_SHEET);
  const dateCell = sheet.getRange("B4").getValue();
  if (!dateCell || !(dateCell instanceof Date)) throw new Error("❌ Invalid date in Dashboard!B4");
  return new Date(dateCell);
}


/**
 * 🔎 Get the true last row in a given column
 *
 * @param {Sheet} sheet
 * @param {number} column
 * @returns {number}
 */
function getRealLastRow(sheet, column = 1) {
  const values = sheet.getRange(2, column, sheet.getMaxRows() - 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 1;
  }
  return values.length + 1;
}


/**
 * 🧼 Reset Learning Log to 1000 rows cleanly
 */
function hardResetLearningLog() {
  const sheet = getPhaseLogSheet();
  if (!sheet) return;

  const lastRowWithData = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastRowWithData < maxRows) {
    sheet.deleteRows(lastRowWithData + 1, maxRows - lastRowWithData);
  }

  sheet.insertRowsAfter(lastRowWithData, 1000 - lastRowWithData);
}


/**
 * 📘 Get or create the current phase's Learning Log sheet
 *
 * @returns {Sheet}
 */
function getPhaseLogSheet() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const phase = dashboard.getRange("B3").getValue();
  const sheetName = phase == 1 ? "Learning Log" : `Learning Log ${phase}`;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const baseSheet = ss.getSheetByName("Learning Log");
    if (!baseSheet) throw new Error("Base 'Learning Log' sheet not found.");

    sheet = baseSheet.copyTo(ss);
    sheet.setName(sheetName);

    const maxRows = sheet.getMaxRows();
    if (maxRows > 2) {
      sheet.deleteRows(3, maxRows - 2);
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol >= 1) {
      sheet.getRange(2, 1, 1, lastCol).clearContent();
    }

    Logger.log(`✅ Created clean phase log sheet: '${sheetName}'`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ '${sheetName}' created.\nClick again to log your entry.`,
      "New Phase Log Ready",
      4
    );
  }

  return sheet;
}

