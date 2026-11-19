/**
 * 🧼 UIExtended.gs
 * ─────────────────────────────────────────────
 * Handles styling logic and sheet-specific themes for all
 * non-Dashboard sheets in the Japanese Study Tracker system.
 *
 * Functions here polish the layout, enable zebra striping,
 * apply thematic backgrounds, and visually normalize various
 * custom sheets (Learning Log, Resource Guide, SmartPrompts, etc).
 *
 * 🧠 Depends on:
 * - Constant: THEMES → Theme definitions
 * - Constant: THEME_KEY_CELL → Dashboard theme selector (B30)
 * - Function: getSheetIdFromDashboard() → for sheet access
 */

// ─────────────────────────────────────────────


/**
 * 🧱 Base styling for a target sheet (font, color, grid size)
 *
 * @param {Sheet} sheet - The target sheet
 * @param {Object} theme - Theme object with fontColor and name
 */
function applyBaseStylingToSheet(sheet, theme) {
  if (!sheet) return; // 🛑 Exit early if sheet not found

  const range = sheet.getRange("A1:Z1000"); // 🎯 Target full aesthetic range

  // → Apply base typography and color
  range.setFontFamily("Segoe UI");
  range.setFontSize(10);
  range.setFontColor(theme.fontColor);
  range.setBackground(theme.name === "Dark" ? "#1a1a1a" : "#ffffff");

  // → Structural tweaks
  sheet.setRowHeights(1, sheet.getMaxRows(), 24);
  sheet.setColumnWidths(1, 26, 110);
  sheet.setFrozenRows(1); // 📌 Lock header
}

/**
 * 🦓 Apply zebra-striping to content rows (starting from row 2)
 *
 * @param {Sheet} sheet - The target sheet
 * @param {Object} theme - Active theme with name
 */
function applyZebraStriping(sheet, theme) {
  const range = sheet.getRange("A2:Z1000");
  const altColor = theme.name === "Dark" ? "#2a2a2a" : "#f7f7f7";
  const baseColor = theme.name === "Dark" ? "#1a1a1a" : "#ffffff";

  const existingBackgrounds = range.getBackgrounds();
  const newBackgrounds = [];

  for (let i = 0; i < existingBackgrounds.length; i++) {
    const isEven = i % 2 === 0;
    const rowColor = isEven ? baseColor : altColor;
    const newRow = Array(26).fill(rowColor);
    newBackgrounds.push(newRow); // 🎨 Push row with correct stripe
  }

  range.setBackgrounds(newBackgrounds); // 🖌️ Apply striping
}

/**
 * 📏 Auto-resize all columns and re-enable wrapping for key columns
 *
 * @param {Sheet} sheet - The target sheet
 */
function autoResizeAllColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), lastCol);

  fullRange.setWrap(false); // 🔧 Disable wrap before resize

  for (let col = 1; col <= lastCol; col++) {
    sheet.autoResizeColumn(col); // 🔄 Auto-fit each column
  }

  // 🧵 Restore wrapping where needed based on sheet name
  const name = sheet.getName();
  switch (name) {
    case "Learning Log":
      sheet.getRange("C:D").setWrap(true);
      break;
    case "SmartPrompts":
      sheet.getRange("E:E").setWrap(true);
      break;
    case "Resource Guide":
      sheet.getRange("D:E").setWrap(true);
      break;
    case "Master Lesson Plan":
      sheet.getRange("E:E").setWrap(true);
      break;
  }
}

/**
 * 🎨 Apply current Dashboard theme (from B30) to all sheets except Dashboard.
 * 🧼 Also trims excess rows in Learning Log to prevent row bloat from theme application.
 */
function applyThemeToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) return;

  const themeName = dashboard.getRange("B30").getValue() || "Light";
  const theme = THEMES[themeName] || THEMES["Light"];
  theme.name = themeName;

  const excluded = ["Dashboard"];
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!excluded.includes(name)) {
      applyBaseStylingToSheet(sheet, theme); // 🧱 Apply base polish

      // 🧩 Apply conditional formatting per sheet
      switch (name) {
        case "Learning Log":
          styleLearningLog(sheet, theme);
          trimExtraRows(sheet); // ✂️ Prevent blank row bloat after theming
          break;
        case "SmartPrompts":
          styleSmartPrompts(sheet, theme);
          break;
        case "Resource Guide":
          styleResourceGuide(sheet, theme);
          break;
        case "Master Lesson Plan":
          styleLessonPlan(sheet, theme);
          break;
        default:
          styleGenericDataSheet(sheet, theme);
      }

      autoResizeAllColumns(sheet); // 📏 Final fit & polish
    }
  });
}


/**
 * ✂️ Trims all empty rows after the last row containing content.
 * Only affects blank rows — data is never deleted.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
 */
function trimExtraRows(sheet) {
  const last = sheet.getLastRow();
  const max = sheet.getMaxRows();
  if (last > 0 && max > last) {
    sheet.deleteRows(last + 1, max - last);
    Logger.log(`🧹 Trimmed ${max - last} blank rows from "${sheet.getName()}"`);
  }
}


/**
 * 📝 Learning Log styling: header, tags, zebra, wrapping
 */
function styleLearningLog(sheet, theme) {
  sheet.getRange("1:1")
    .setBackground(theme.headerBg)
    .setFontColor(theme.fontColor)
    .setFontWeight("bold");

  sheet.getRange("A:A").setFontWeight("bold");
  applyZebraStriping(sheet, theme);
}

/**
 * 🧠 SmartPrompts sheet: italic prompt column + header styling
 */
function styleSmartPrompts(sheet, theme) {
  sheet.getRange("1:1")
    .setBackground(theme.headerBg)
    .setFontColor(theme.fontColor)
    .setFontWeight("bold");

  sheet.getRange("E:E").setFontStyle("italic");
  applyZebraStriping(sheet, theme);
}

/**
 * 🗂️ Resource Guide: wrap notes and bold progress columns
 */
function styleResourceGuide(sheet, theme) {
  sheet.getRange("1:1")
    .setBackground(theme.headerBg)
    .setFontColor(theme.fontColor)
    .setFontWeight("bold");

  sheet.getRange("E:E").setFontWeight("bold");
  sheet.getRange("F:F").setFontWeight("bold");
  applyZebraStriping(sheet, theme);
}

/**
 * 📚 Master Lesson Plan: align prompts + apply zebra styling
 */
function styleLessonPlan(sheet, theme) {
  sheet.getRange("1:1")
    .setBackground(theme.headerBg)
    .setFontColor(theme.fontColor)
    .setFontWeight("bold");

  sheet.getRange("C:F").setHorizontalAlignment("left");
  applyZebraStriping(sheet, theme);
}

/**
 * 🧾 Generic fallback style for any non-specified sheet
 */
function styleGenericDataSheet(sheet, theme) {
  sheet.getRange("1:1")
    .setBackground(theme.headerBg)
    .setFontColor(theme.fontColor)
    .setFontWeight("bold");

  applyZebraStriping(sheet, theme);
}

/**
 * 💡 Show SmartPrompt inspiration sidebar
 */
function openSmartPromptSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SmartpromptSidebar")
    .setTitle("💡 SmartPrompt")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 📘 Show Study Session Logger sidebar
 */
function logStudySession() {
  const html = HtmlService.createHtmlOutputFromFile("StudySessionLogger")
    .setTitle("📘 Log Study Session")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 🈲 Show Kanji/Vocab Logger sidebar
 */
function openKanjiLoggerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("KanjiLogger")
    .setTitle("📘 Kanji/Vocab Logger")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 🗃️ Open Resource Guide tab or show alert if missing
 */
function openResourceTracker() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const sheet = ss.getSheetByName("Resource Guide");
  if (sheet) sheet.activate();
  else SpreadsheetApp.getUi().alert("❌ Resource Guide sheet not found.");
}

