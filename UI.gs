/**
 * 📁 UI.gs
 * ─────────────────────────────────────────────
 * Dashboard styling and visual polish for the Japanese Learning Tracker.
 *
 * This module handles dynamic theme switching (Light/Dark), font and layout 
 * formatting, sparkline notes, ID/email zones, and optional header resizing.
 * 
 * 🧼 Functions:
 * - applyDashboardFormatting() → Main UI polish based on selected theme
 * - applyStreakHighlight(themeName) → Conditionally styles streak cell (B12)
 * - openThemeSidebar() → Opens sidebar to allow user to change theme
 * - resizeDashboardHeaderRow() → Resets header row height to default size
 * 
 * 🎨 Themes:
 * Pulled from the `THEMES` constant map based on Dashboard B30
 * 
 * 📌 Target Sheet: "Dashboard"
 */

// ─────────────────────────────────────────────


/**
 * 🧼 Apply full visual polish and theming to the Dashboard sheet
 * Dynamically styles zones, fonts, colors, and conditional logic
 */
function applyDashboardFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!sheet) return;

  const themeName = sheet.getRange("B30").getValue() || "Light";
  const theme = THEMES[themeName] || THEMES["Light"];

  // 🧽 Background color reset
  sheet.getRange("A1:Z100").setBackground(themeName === "Dark" ? "#1a1a1a" : "#ffffff");

  // 🔤 Font color for major content zones
  const textRanges = [
    "A1:D1", "A3:B4", "A6:B23",
    "A25:B25", "A27:B28", "A30:B30", "C10"
  ];
  textRanges.forEach(range => {
    sheet.getRange(range).setFontColor(theme.fontColor);
  });

  // 🧠 Header row title
  sheet.setRowHeight(1, 45);
  sheet.getRange("A1:D1").merge().setValue(theme.dashboardTitle);
  sheet.getRange("A1:D1").setBackground(theme.headerBg);
  sheet.getRange("B1")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 🟨 Phase zone background
  sheet.getRange("A3:B4").setBackground(theme.phaseZone);

  // 📋 Metric layout (borders, alignment)
  const metricRange = sheet.getRange("A6:B23");
  metricRange.setBorder(true, true, true, true, true, true);
  metricRange.setVerticalAlignment("middle");

  sheet.getRange("A6:A23")
    .setFontWeight("bold")
    .setHorizontalAlignment("right")
    .setFontColor(theme.fontColor);

  sheet.getRange("B6:B23")
    .setHorizontalAlignment("left")
    .setFontColor(theme.fontColor);

  // 🔵 Hyperlink override (most studied resource)
  sheet.getRange("B18")
    .setFontColor("#1a73e8")
    .setFontWeight("bold");

  // 💬 Sparkline note (set once only)
  const sparkCell = sheet.getRange("C10");
  const existing = sparkCell.getValue();
  if (!existing || !existing.toString().includes("←")) {
    sparkCell.setValue("← Habit Radar: █ = studied, ░ = not").setFontStyle("italic");
  }
  sparkCell.setFontColor(theme.fontColor);

  // 📨 Email zone
  sheet.getRange("A25:B25").setBackground(theme.emailZone);

  // 🔐 ID zone
  sheet.getRange("A27:B28").setBackground(theme.idZone);
  sheet.getRange("B27:B28")
    .setFontFamily("Courier New")
    .setFontSize(9)
    .setWrap(true);

  // 🔎 Explicit font color overrides for edge cases
  ["B25", "B27", "B28", "B30"].forEach(cell => {
    sheet.getRange(cell).setFontColor(theme.fontColor);
  });

  // ✅ Highlight current streak dynamically
  applyStreakHighlight(themeName);
}


/**
 * ✅ Highlight streak cell (B12) if current streak > 7
 *
 * @param {string} themeName - "Dark" or "Light" to determine green contrast
 */
function applyStreakHighlight(themeName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const range = sheet.getRange("B12");

  const bgColor = themeName === "Dark" ? "#336633" : "#d9ead3";

  // 🚫 Clear any existing rules targeting B12 to avoid duplicates
  const rules = sheet.getConditionalFormatRules().filter(r =>
    !r.getRanges().some(rg => rg.getA1Notation() === "B12")
  );

  // ➕ Apply new rule if regex-extracted value > 7
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=VALUE(REGEXEXTRACT(B12, "\\d+")) > 7`)
    .setBackground(bgColor)
    .setBold(true)
    .setRanges([range])
    .build();

  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}


/**
 * 🎨 Open the Theme Sidebar in the UI for theme selection
 */
function openThemeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ThemeSidebar")
    .setTitle("Theme Settings")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * 📏 Reset the dashboard header row height to default (45px)
 */
function resizeDashboardHeaderRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!sheet) return;
  sheet.setRowHeight(1, 45);
  Logger.log("✅ Dashboard header row resized to 45px.");
}

