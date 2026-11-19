/**
 * 🩺 Healing.gs — Session Anchor Normalization
 * ─────────────────────────────────────────────
 * 📋 Purpose:
 * This script ensures that all manually-entered session anchors (Column F) in the Learning Log
 * are cleaned and safely reformatted as plain text, locking in their displayed local dates.
 *
 * 🔥 Why It's Needed:
 * Google Apps Script stores dates internally in UTC. This causes a mismatch for users
 * in timezones behind UTC (like most US timezones), making dates appear one day behind
 * unless healed.
 *
 * 🛡️ What It Does:
 * - Converts session anchor cells to plain text without changing their visible display.
 * - Safely preserves intended dates without forcing re-entry.
 * - Corrects past entries that were manually reformatted.
 *
 * 🚀 Usage Workflow:
 * 1. Log your study session as normal.
 * 2. Select the **next day** when picking a date (forward pick).
 * 3. After logging, run "🩺 Heal Session Anchors" from the Study System → Clean-Up menu.
 *
 * ✅ After healing, your dashboard and study tracking will match your intended local dates.
 *
 * ⚙️ Safety:
 * - Non-destructive: Only normalizes Column F ("Session Anchor").
 * - No changes are made to other logged data (time, tags, difficulty, etc.).
 *
 * 🧠 Reminder:
 * Healing is necessary **only for users in timezones behind UTC**. 
 */

// ─────────────────────────────────────────────

/**
 * 
 * Normalizes all Date-type values in Column F ("Session Anchor") to plain-text
 * M/D/YYYY strings in local timezone. This locks the visual display while resolving
 * timezone-induced misalignment from Google's internal UTC handling.
 */
function auditAndHealSessionAnchors() {
  const sheet = getPhaseLogSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // ⛔ No data to heal

  const anchorRange = sheet.getRange(2, 6, lastRow - 1, 1); // 📌 Column F (Session Anchor)
  const anchorValues = anchorRange.getValues();

  const healedRows = [];

  for (let i = 0; i < anchorValues.length; i++) {
    const rowNum = i + 2;
    const raw = anchorValues[i][0];

    // 🧱 Skip if empty, already string, or malformed
    if (!raw || typeof raw !== "object" || !(raw instanceof Date)) continue;

    const local = Utilities.formatDate(raw, getUserTimeZone(), "M/d/yyyy");


    const cell = sheet.getRange(rowNum, 6); // F = Session Anchor
    cell.clearContent();                    // 🧼 Remove Date type (required for clean overwrite)
    cell.setNumberFormat("@STRING@");       // 🔒 Lock as plain text
    cell.setValue(local);                   // ✍️ Write local-date string
    healedRows.push(rowNum);
    Logger.log(`🛠️ Row ${rowNum} healed: ${raw} → ${local}`);
  }

  // 🍞 Notify user with toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ ${healedRows.length} session anchors healed.`,
    "UTC Fix Complete",
    6
  );
}

/**
 * 
 * Resets the Type Tag column (Column K) to mirror values from Column B ("Type").
 * Re-applies the formula down all valid rows.
 *
 * Use this after healing, bulk edits, or if formulas were overwritten by hand.
 */
function refreshTypeTagColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learning Log");
  if (!sheet) return;

  const rowCount = sheet.getLastRow();
  if (rowCount < 2) return; // ⛔ No rows to process

  // 🧪 Formula: if Type is blank, blank; else, Proper-cased
  const formula = `=IF(B2="", "", PROPER(B2))`;

  // 🎯 Reapply to top, fill down
  sheet.getRange("K2").setFormula(formula);
  const fillRange = sheet.getRange("K2:K" + rowCount);
  sheet.getRange("K2").autoFill(fillRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // 🍞 Toast success
  SpreadsheetApp.getActiveSpreadsheet().toast("🔁 Type Tags refreshed!", "Success", 3);
}

