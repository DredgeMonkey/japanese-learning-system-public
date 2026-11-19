/**
 * 📄 Code.gs
 * ─────────────────────────────────────────────  
 * This file contains core system constants and basic utility functions shared across scripts.
 * These constants define sheet names and structure used throughout the Japanese Study Tracker.
 * 
 * The `getSheetFromDashboardOrActive()` utility checks for sheet context between active and dashboard-linked spreadsheets.
 * Use this as a safe fallback when working across multiple spreadsheet instances.
 */

// ─────────────────────────────────────────────

// 🔧 System Constants
var DASHBOARD_SHEET = "Dashboard";
var KANJI_VOCAB_COLUMN = "Kanji/Vocab Learned";
var PHASE_COLUMN = 2;
var DEBUG_LOG_SHEET = "Debug Log";

/**
 * 📘 Get Sheet by Name from Dashboard or Active Spreadsheet
 * 
 * Attempts to retrieve a sheet by name from the linked spreadsheet ID on the Dashboard (B27),
 * falling back to the active spreadsheet if unavailable.
 *
 * @param {string} sheetName - The name of the sheet to retrieve
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The requested sheet object, or null if not found
 */
function getSheetFromDashboardOrActive(sheetName) {
  const dashboardId = getSheetIdFromDashboard();
  const activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("🆔 Dashboard ID: " + dashboardId);
  Logger.log("🆔 Active Sheet ID: " + activeId);

  if (dashboardId === activeId) {
    Logger.log("📘 Returning Active Sheet");
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }

  try {
    Logger.log("📘 Returning Dashboard Sheet by ID");
    return SpreadsheetApp.openById(dashboardId).getSheetByName(sheetName);
  } catch (err) {
    Logger.log("⚠️ Failed to open by Dashboard ID. Falling back to active spreadsheet.");
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }
}

