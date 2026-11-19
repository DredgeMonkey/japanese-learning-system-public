/**
 * 📁 SeasonLog.gs
 * ─────────────────────────────────────────────
 * Logging utility for serialized media study (anime, dramas, etc.)
 * 
 * This module supports logging entire *seasons* of content via sidebar UI,
 * contributing to total study hours without inflating session count metrics.
 * 
 * ⚙️ Functions:
 * - logSeasonEntry(entry) → Logs an entry to the Learning Log sheet
 * - openSeasonLoggerSidebar() → Launches the sidebar form interface
 * 
 * 📝 Logged entries are auto-tagged as "Media: [Type]" (default: Anime)
 * 📌 Session anchor (Column F) is date-controlled and formatted as M/d/yyyy
 */

// ─────────────────────────────────────────────


/**
 * 📝 Log a full season entry into the Learning Log
 * 
 * Accepts a structured entry object and writes it into the log
 * without affecting session count or difficulty metrics.
 * 
 * @param {Object} entry - Contains title, tags, mediaType, date, time, and notes
 */
function logSeasonEntry(entry) {
  if (!entry || typeof entry !== 'object') return;

  const sheet = getPhaseLogSheet();
  if (!sheet) return;

  cleanGhostRowsFromLearningLog();
  const firstEmptyRow = getFirstEmptyRowInColumn(sheet, 1);
  const timestamp = new Date();

  // Validate and format the date input
  const parsedDate = new Date(entry.date);
  if (isNaN(parsedDate.getTime())) {
    Logger.log(`❌ Invalid date provided for season log: ${entry.date}`);
    SpreadsheetApp.getUi().alert("❌ Invalid date provided.");
    return;
  }

  const formattedDate = Utilities.formatDate(parsedDate, getUserTimeZone(), "M/d/yyyy");

  // Apply default media tag if missing (e.g., "Media: Anime")
  const finalTags = entry.tags?.trim()
    ? `Media: ${entry.mediaType || "Anime"}, ${entry.tags.trim()}`
    : `Media: ${entry.mediaType || "Anime"}`;

  const values = [
    timestamp,                  // A: Audit Timestamp
    "Season Logger",            // B: Type
    entry.title || "Unknown",  // C: Resource (season name)
    finalTags,                  // D: Tags (media type + optional tags)
    "",                         // E: Day #
    formattedDate,             // F: Session Anchor (study date)
    entry.time || 0,           // G: Time Spent (minutes)
    false,                     // H: Completed (always false for seasons)
    "",                        // I: Difficulty (not applicable)
    entry.notes || ""          // J: Notes
  ];

  sheet.getRange(firstEmptyRow, 1, 1, values.length).setValues([values]);
  reapplyFormattingToLearningLog(firstEmptyRow);

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Season entry logged!", "Season Logger", 4);
  Logger.log(`✅ Season logged: ${entry.title} (${entry.time} min)`);
}


/**
 * 📤 Open the Season Logger input form as a sidebar
 * 
 * Loads the SeasonLogger.html UI file with fixed width
 */
function openSeasonLoggerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SeasonLogger")
    .setTitle("🎬 Season Logger")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

