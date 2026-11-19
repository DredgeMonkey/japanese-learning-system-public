/**
 * 🧩 MenuExtras.gs
 * ─────────────────────────────────────────────
 * Contains optional or utility-driven menu actions that don't belong 
 * in the main `Menu.gs` logic or have their own dedicated UI module.
 *
 * 📌 Current Actions:
 * - `showSystemGuideSidebar()`: Opens the full HTML system documentation.
 * - `forceRealOAuth()`: Forces a read/write operation to trigger re-auth.
 * - `manuallyTrimLearningLog()`: Snips unused rows at the bottom of Learning Log.
 *
 * 🛠️ Use this file to house specialized menu triggers or diagnostic helpers
 * that are supplemental to the core Tracker system but enhance UX or recovery.
 */

// ─────────────────────────────────────────────



/**
 * Loads the HTML sidebar file `SystemGuideSidebar.html` and displays it in the Sheets UI.
 */
function showSystemGuideSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SystemGuideSidebar")
    .setTitle("📘 Study Tracker: System Overview");
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Triggers a dummy read/write operation to force Apps Script to re-request permissions,
 * resolving issues where OAuth has expired or broken silently.
 */
function forceRealOAuth() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learning Log");
  const range = sheet.getRange("A1:B1");
  const values = range.getValues(); // READ
  sheet.getRange("C1:D1").setValues(values); // WRITE
}


/**
 * Menu action to manually trim extra rows in the Learning Log.
 */
function manuallyTrimLearningLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learning Log");
  if (sheet) trimExtraRows(sheet);
}

