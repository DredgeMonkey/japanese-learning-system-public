/**
 * 🧾 LoggerTools.gs
 * ─────────────────────────────────────────────
 * Handles logging of study sessions and structured learning entries
 * into the current phase's Learning Log sheet.
 * 
 * 📒 Features:
 * - `logFullStudyEntry(entry)` logs full session records with timestamps
 * - `logLearningEntry(type, content, tags, date)` logs quick entries 
 *   (e.g. kanji, vocab, grammar snippets) via the sidebar
 * 
 * 📌 Notes:
 * - All entries append to the current phase sheet (auto-detected)
 * - Uses date validation to guard against timezone/logging issues
 * - Integrates formatting + cleanup safeguards where needed
 */

// ─────────────────────────────────────────────


/**
 * 📝 Append a full study session entry to the active phase log
 *
 * @param {Object} entry - Session data passed from the sidebar UI
 * @returns {string} Status message after logging
 */
function logFullStudyEntry(entry) {
  if (!entry || typeof entry !== 'object') return;

  const sheet = getPhaseLogSheet();
  if (!sheet) return;

  // 📅 Parse and validate fields
  const timestamp = new Date();                            // Current log time
  const parsedDate = new Date(entry.date);                 // User-selected date

  const timeSpent = parseFloat(entry.time);                // Minutes studied
  const isCompleted = entry.completed === true;            // Optional checkbox flag

  // 📥 Append row with all fields mapped in order
  sheet.appendRow([
    timestamp,                         // A: Timestamp
    entry.type || "Study Session",     // B: Type (default if missing)
    entry.resource || "Unknown",       // C: Resource name
    entry.focus || "",                 // D: Study focus
    entry.day || "",                   // E: Day number
    isNaN(parsedDate) ? new Date() : parsedDate, // F: Date (validated)
    isNaN(timeSpent) ? 0 : timeSpent,  // G: Minutes spent
    isCompleted,                       // H: Completed checkbox (optional)
    entry.difficulty || "",            // I: Difficulty rating
    entry.notes || ""                  // J: Notes
  ]);

  Logger.log("✅ Entry successfully appended.");
  return "Logged successfully.";
}


/**
 * 🧠 Log a learning note (e.g. vocab, kanji, grammar) via Kanji Logger
 *
 * @param {string} type - Log type ("Vocab", "Kanji", etc.)
 * @param {string} content - Entry content (e.g. word or concept)
 * @param {string} tags - Comma-separated tags
 * @param {string} date - Date selected via date picker
 */
function logLearningEntry(type, content, tags, date) {
  let sheet = getPhaseLogSheet();

  // 🛠 Sheet load buffer to ensure active context
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  sheet = getPhaseLogSheet();
  if (!sheet) return;

  cleanGhostRowsFromLearningLog(); // 🧹 Clear blank entries for stability

  const firstEmptyRow = getFirstEmptyRowInColumn(sheet, 1);
  const timestamp = new Date();

  // 📅 Validate the selected date from UI
  const parsedDate = new Date(date);
  if (isNaN(parsedDate.getTime())) {
    Logger.log(`❌ Invalid date provided: ${date}`);
    SpreadsheetApp.getUi().alert("❌ Error: Invalid date provided.");
    return;
  }

  // 🗓 Format date as plain text (not serial)
  const formattedDate = Utilities.formatDate(parsedDate, getUserTimeZone(), "M/d/yyyy");

  const values = [
    timestamp,                 // A: Timestamp (audit)
    type || "Log",             // B: Type
    content || "",             // C: Resource
    tags || "",                // D: Tags
    "",                        // E: Day #
    formattedDate,             // F: Session Anchor (as plain text)
    "", "", false, "", ""      // G–K: Remaining fields left empty
  ];

  // 📤 Write values to sheet
  sheet.getRange(firstEmptyRow, 1, 1, values.length).setValues([values]);

  // 🧽 Reapply formatting if this row exceeds last formatted area
  const lastFormattedRow = sheet.getLastRow();
  if (firstEmptyRow > lastFormattedRow) {
    reapplyFormattingToLearningLog(firstEmptyRow);
  }

  Logger.log(`✅ Learning log written: ${type} - ${content} on ${formattedDate}`);
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Entry Logged!", "Kanji/Vocab Logger", 3);
}

