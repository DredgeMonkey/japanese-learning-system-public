/**
 * 🎨 Theme.gs
 * ─────────────────────────────────────────────
 * Simple theme setter for the Dashboard UI.
 * 
 * This function stores the selected theme in a dedicated cell
 * and refreshes all dashboard formatting to apply the new style.
 * 
 * 🧠 Depends on:
 * - Constant: THEME_KEY_CELL → target cell to store current theme
 * - Function: applyDashboardFormatting() → redraws dashboard visuals
 */

// ─────────────────────────────────────────────


/**
 * 🎨 Apply selected theme and refresh Dashboard styling
 * 
 * @param {string} theme - The name of the theme to apply
 */
function setTheme(theme) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!sheet) return; // 🛑 Exit if Dashboard sheet not found

  sheet.getRange(THEME_KEY_CELL).setValue(theme); // 🎯 Save theme to control cell
  applyDashboardFormatting(); // 🧽 Refresh visuals with new theme
  Logger.log(`🎨 Theme set to: ${theme}`); // 📋 Log for diagnostics
}

