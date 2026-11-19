/**
 * 🧰 Utils.gs
 * ─────────────────────────────────────────────
 * Small utility functions used across multiple files.
 * 
 * 📌 Current Utilities:
 * - `normalizeTitle(title)`: Standardizes task or entry titles by 
 *   removing unwanted symbols and harmonizing spacing.
 * 
 * 🔧 Expand this file as needed for additional cross-file helpers.
 */

// ─────────────────────────────────────────────


/**
 * 🔤 normalizeTitle(title)
 * Standardizes a title by:
 * - Unifying full-width tildes to space
 * - Removing unwanted characters (but keeps kana, kanji, letters, numbers, and spaces)
 * - Collapsing multiple spaces into one
 * - Trimming leading/trailing whitespace
 */
function normalizeTitle(title) {
  if (!title) return "";
  return title
    .replace(/[～〜]/g, " ")                // Replace full-width tildes with space
    .replace(/[^\wぁ-んァ-ン一-龥ー ]/g, "")  // Keep word characters + kana + kanji + space
    .replace(/\s+/g, " ")                  // Collapse multiple spaces
    .trim();                               // Trim leading/trailing spaces
}


/**
 * ⏰ getUserTimeZone()
 * Returns the user or spreadsheet time zone for consistent date handling.
 * Falls back to the script time zone when unavailable.
 */
function getUserTimeZone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
}
