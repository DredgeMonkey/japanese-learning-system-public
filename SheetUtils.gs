/**
 * 📄 SheetUtils.gs
 * ─────────────────────────────────────────────
 * Utility functions to extract key configuration values
 * from the Dashboard sheet, such as:
 *
 * • Sheet ID (B27)
 * • Calendar ID (B28)
 *
 * Includes test functions for verifying ID extraction via UI alerts.
 * Throws UI-level errors or logs where required for invalid inputs.
 */

// ─────────────────────────────────────────────


/**
 * 🆔 Get linked Sheet ID from Dashboard (cell B27)
 *
 * @returns {string} - Sheet ID (fallbacks to active spreadsheet ID if blank)
 */
function getSheetIdFromDashboard() {
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET);
  if (!dash) throw new Error("❌ Dashboard sheet not found.");
  const id = dash.getRange("B27").getValue().toString().trim();
  return id || SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * 📅 Get Calendar ID from Dashboard (cell B28)
 *
 * @returns {string} - Calendar ID for syncing tasks
 * @throws {Error} - If ID is blank or malformed
 */
function getCalendarIdFromDashboard() {
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET);
  if (!dash) throw new Error("❌ Dashboard sheet not found.");
  const id = dash.getRange("B28").getValue().toString().trim();

  const isValidId = id && id.includes("@") && id.length > 10;
  if (!isValidId) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ Missing or invalid Calendar ID in Dashboard cell B28.\nPlease check and try again.",
      "Calendar Sync Error", 6
    );
    Logger.log("❌ Calendar ID is missing or malformed.");
    throw new Error("Calendar ID is missing or invalid.");
  }

  return id;
}

/**
 * 🧪 Debug popup to verify extracted Sheet ID
 */
function testSheetIdPopup() {
  const id = getSheetIdFromDashboard();
  SpreadsheetApp.getUi().alert("✅ Sheet ID Loaded:\n" + id);
}

/**
 * 🧪 Debug popup to verify extracted Calendar ID
 */
function testCalendarIdPopup() {
  try {
    const id = getCalendarIdFromDashboard();
    SpreadsheetApp.getUi().alert("✅ Calendar ID Loaded:\n" + id);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Calendar ID Error:\n" + e.message);
  }
}

