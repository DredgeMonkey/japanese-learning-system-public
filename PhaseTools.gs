/**
 * 📁 PhaseTools.gs
 * ─────────────────────────────────────────────
 * Utility functions for managing **phase-based learning blocks**
 * in the Master Lesson Plan. This includes:
 * 
 * - Assigning new phase dates (e.g., when a phase begins)
 * - Extending current phases (by fixed or prompt-based increments)
 * - Auditing the date span and day count of any phase
 * - Ensuring enough rows exist for new assignments
 * - Reading the current phase/day as derived from the dashboard
 * 
 * Key integration points:
 * - Reads from Dashboard sheet cell B3 (current phase) and B4 (start date)
 * - Writes to columns A (Day), B (Phase), C (Date) in Master Lesson Plan
 * - Enforces a **maximum cap of 180 days per phase**
 * - Includes alert-based safeguards to avoid malformed entries
 * 
 * Used extensively by:
 * - Menu.gs (Phase Tools menu)
 * - Dashboard.gs (Phase metrics and daily logic)
 * - Logging and syncing operations (date-anchored)
 * 
 * This file consolidates all phase-related logic in one place
 * to improve code organization, prevent duplication, and simplify auditing.
 */

// ─────────────────────────────────────────────

/**
 * Assigns the current phase's start date and fills 90 sequential rows
 * in the Master Sheet with [Day, Phase, Date] values.
 * 
 * Ensures task/calendar syncing will work by populating Phase data.
 */
function assignDatesToActivePhase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!dashboard || !master) throw new Error("❌ Sheets missing.");

  const currentPhase = dashboard.getRange("B3").getValue();
  const phaseStartDate = dashboard.getRange("B4").getValue();

  if (!currentPhase) throw new Error("❌ Phase missing.");
  if (!(phaseStartDate instanceof Date)) throw new Error("❌ Invalid start date.");

  const maxScan = master.getMaxRows(); // auto-scans all physical rows in the sheet
  const data = master.getRange(2, 1, maxScan - 1, 3).getValues(); // A2:C...

  let startRow = null;
  for (let i = 0; i < data.length; i++) {
    const [day, phase, date] = data[i];
    if (!day && !phase && !date) {
      startRow = i + 2; // Adjust for header
      break;
    }
  }

  if (!startRow) {
    SpreadsheetApp.getUi().alert("❌ No available blank rows in A–C to assign new phase dates.");
    Logger.log("⚠️ No blank A–C blocks found up to row " + maxScan);
    return;
  }

  const fillRows = 90;
  const values = [];

  for (let i = 0; i < fillRows; i++) {
    const date = new Date(phaseStartDate);
    date.setDate(date.getDate() + i);
    values.push([i + 1, currentPhase, date]);
  }

  master.getRange(startRow, 1, fillRows, 3).setValues(values);
  Logger.log(`✅ Phase ${currentPhase} filled from row ${startRow} for ${fillRows} rows.`);
}


/**
 * Automatically extends the current phase in the Master Sheet by up to 30 days,
 * respecting the 180-day hard cap.
 */
function extendCurrentPhaseBy30Days() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!dashboard || !master) throw new Error("❌ Missing Dashboard or Master Sheet.");

  const currentPhase = dashboard.getRange("B3").getValue();
  const phaseStartDate = dashboard.getRange("B4").getValue();

  if (!currentPhase || !(phaseStartDate instanceof Date)) {
    SpreadsheetApp.getUi().alert("⚠️ Phase or start date is missing or invalid.");
    return;
  }

  // Step 1: Count existing days for this phase
  const data = master.getRange(2, 1, master.getLastRow() - 1, 3).getValues(); // A–C
  const currentPhaseDays = data.filter(row => row[1] == currentPhase && row[0]).length;

  if (currentPhaseDays >= 180) {
    SpreadsheetApp.getUi().alert("🚫 Cannot extend. Phase is already at the 180-day cap.");
    return;
  }

  // Step 2: Determine how many days we *can* add
  const daysToAdd = Math.min(30, 180 - currentPhaseDays);

  // Step 3: Find the last date assigned in this phase
  const lastPhaseDate = getLastPhaseDate(currentPhase);
  if (!lastPhaseDate) {
    SpreadsheetApp.getUi().alert("⚠️ Could not locate last assigned date for this phase.");
    return;
  }

  // Step 4: Find next empty row in Master
  let firstEmptyRow = master.getLastRow() + 1;
  const maxRows = master.getMaxRows();
  if (firstEmptyRow + daysToAdd > maxRows) {
    master.insertRowsAfter(master.getLastRow(), (firstEmptyRow + daysToAdd) - maxRows);
    Logger.log(`🧱 Inserted ${firstEmptyRow + daysToAdd - maxRows} new rows.`);
  }

  // Step 5: Insert phase assignments
  const values = [];
  for (let i = 1; i <= daysToAdd; i++) {
    const dayNumber = currentPhaseDays + i;
    const newDate = new Date(lastPhaseDate);
    newDate.setDate(newDate.getDate() + i);
    values.push([dayNumber, currentPhase, newDate]);
  }

  master.getRange(firstEmptyRow, 1, values.length, 3).setValues(values);
  Logger.log(`✅ Extended Phase ${currentPhase} by ${daysToAdd} days.`);
  SpreadsheetApp.getUi().alert(`✅ Phase ${currentPhase} extended by ${daysToAdd} days.`);
}


/**
 * Displays a summary alert showing day counts and date ranges
 * for each distinct phase found in the Master Sheet.
 */
function showPhaseAuditLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return SpreadsheetApp.getUi().alert("❌ Master Sheet not found.");

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const phaseMap = {};

  data.forEach(([day, phase, date]) => {
    if (!phase || !day || !(date instanceof Date)) return;
    if (!phaseMap[phase]) phaseMap[phase] = [];
    phaseMap[phase].push({ day, date });
  });

  const lines = Object.entries(phaseMap).map(([phase, entries]) => {
    const sorted = entries.sort((a, b) => a.day - b.day);
    const start = sorted[0].date.toLocaleDateString();
    const end = sorted[sorted.length - 1].date.toLocaleDateString();
    return `📘 Phase ${phase}: ${sorted.length} days — ${start} → ${end}`;
  });

  if (lines.length === 0) {
    SpreadsheetApp.getUi().alert("ℹ️ No phase data found to audit.");
  } else {
    SpreadsheetApp.getUi().alert("📋 Phase Audit Summary:\n\n" + lines.join("\n"));
  }
}


/**
 * Ensures the Master Sheet has at least `minRows` rows.
 * Adds rows only if missing.
 */
function ensureRowCapacity(minRows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) throw new Error("❌ Master Sheet not found.");

  const currentRows = sheet.getMaxRows();
  if (currentRows < minRows) {
    const toAdd = minRows - currentRows;
    sheet.insertRowsAfter(currentRows, toAdd);
    Logger.log(`🧱 Added ${toAdd} new rows to ensure ${minRows} total.`);
  }
}


/**
 * Returns how many rows exist for a given phase in the Master Sheet.
 */
function getPhaseDayCount(phase) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return 0;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.filter(row => row[1] == phase && row[0]).length;
}


/**
 * Returns the latest Date object assigned to a given phase.
 */
function getLastPhaseDate(phase) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const dates = data.filter(r => r[1] == phase && r[2] instanceof Date).map(r => r[2]);
  return dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null;
}


/**
 * Prompts user to enter custom number of days to extend current phase.
 * Caps at 180 total days.
 */
function showPhaseExtensionPrompt() {
  const ui = SpreadsheetApp.getUi();
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET);
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);

  if (!dashboard || !master) {
    ui.alert("❌ Missing Dashboard or Master Sheet.");
    return;
  }

  const currentPhase = dashboard.getRange("B3").getValue();
  const phaseStartDate = dashboard.getRange("B4").getValue();

  if (!currentPhase || !(phaseStartDate instanceof Date)) {
    ui.alert("⚠️ Phase or start date is missing or invalid.");
    return;
  }

  const response = ui.prompt(
    "📅 Extend Phase",
    "Enter how many days to add to the current phase (up to 180 max total):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const requestedDays = parseInt(response.getResponseText());
  if (isNaN(requestedDays) || requestedDays <= 0) {
    ui.alert("⚠️ Invalid input. Please enter a positive number.");
    return;
  }

  const currentPhaseDays = getPhaseDayCount(currentPhase);
  const remainingDays = 180 - currentPhaseDays;

  if (remainingDays <= 0) {
    ui.alert("🚫 Cannot extend. Phase already has 180 days.");
    return;
  }

  const daysToAdd = Math.min(requestedDays, remainingDays);
  const lastPhaseDate = getLastPhaseDate(currentPhase);
  if (!lastPhaseDate) {
    ui.alert("⚠️ Could not locate last assigned date for this phase.");
    return;
  }

  const firstEmptyRow = master.getLastRow() + 1;
  ensureRowCapacity(firstEmptyRow + daysToAdd);

  const values = [];
  for (let i = 1; i <= daysToAdd; i++) {
    const dayNumber = currentPhaseDays + i;
    const newDate = new Date(lastPhaseDate);
    newDate.setDate(newDate.getDate() + i);
    values.push([dayNumber, currentPhase, newDate]);
  }

  master.getRange(firstEmptyRow, 1, values.length, 3).setValues(values);
  Logger.log(`✅ Phase ${currentPhase} extended by ${daysToAdd} days via prompt.`);
  ui.alert(`✅ Phase ${currentPhase} extended by ${daysToAdd} days.`);
}


/**
 * Returns the current study day number for today's date in the active phase.
 * Used for dashboard logic and tagging.
 */
function getCurrentStudyDay() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!dashboard || !master) return null;

  const currentPhase = dashboard.getRange("B3").getValue();
  const today = new Date();
  const lastRow = master.getLastRow();
  if (lastRow < 2) return null;

  const data = master.getRange(2, 1, lastRow - 1, 3).getValues();
  let matchedDay = null;

  for (const [dayLabel, phase, date] of data) {
    if (String(phase).trim() !== String(currentPhase).trim()) continue;
    if (!(date instanceof Date)) continue;

    const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    if (dateOnly.getTime() === todayOnly.getTime()) {
      matchedDay = parseInt(dayLabel.toString().replace(/[^0-9]/g, ""));
      break;
    }
    if (dateOnly < todayOnly) {
      matchedDay = parseInt(dayLabel.toString().replace(/[^0-9]/g, ""));
    }
  }

  return matchedDay || null;
}


/**
 * Fetches the current active phase from Dashboard B3.
 * Clean, final version to be used across codebase.
 */
function getCurrentPhase() {
  const dashboardSheet = SpreadsheetApp.openById(getSheetIdFromDashboard()).getSheetByName(DASHBOARD_SHEET);
  if (!dashboardSheet) throw new Error("❌ Dashboard sheet not found.");
  const phaseValue = dashboardSheet.getRange("B3").getValue();
  return phaseValue.toString().trim();
}

