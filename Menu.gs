/**
 * 🧭 Menu.gs — UI Menu Configuration
 * ─────────────────────────────────────────────
 * 📋 Purpose:
 * Creates all custom menu items shown in the Google Sheets UI. These menus trigger the key functions
 * across the study tracker system, enabling user interaction without needing to run scripts manually.
 *
 * 🔧 Features:
 * - Centralizes all available UI actions for logging, syncing, reporting, cleanup, styling, and diagnostics
 * - Automatically runs on sheet open via `onOpen()`
 * - Integrates multiple menu sections for modular navigation
 *
 * 🧠 Key Menus:
 * - 🧾 Study Loggers
 * - 📋 Task Syncing
 * - 📅 Calendar Syncing
 * - 🌀 Phase Tools
 * - 🧼 Clean-Up Tools
 * - 🧪 Developer & Diagnostics
 * - 📊 Dashboard Tools
 * - 🎨 Theme Tools
 * - 🎌 Study System (Master Menu)
 * - 🃏 Flashcards
 * - 📘 System Guide
 * - 🛠️ Tooltips & Help (injected by separate function)
 * - 🧭 Real OAuth
 */

// ─────────────────────────────────────────────

/**
 * This is the main trigger that builds all custom menus in the sheet UI when the file is opened.
 * Menus are grouped logically for user clarity and interaction efficiency.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 🧾 Logging Interfaces
  const studyLoggers = ui.createMenu("🧾 Study Loggers")
    .addItem("📝 Log Quick Entry", "logStudyEntry")
    .addItem("📘 Log Study Session", "logStudySession")
    .addItem("🎬 Season Logger", "openSeasonLoggerSidebar")
    .addItem("💡 Get SmartPrompt", "openSmartPromptSidebar");

  // 📋 Task Syncing (by chunks or prompted)
  const taskSyncMenu = ui.createMenu("📋 Task Syncing")
    .addItem("🧾 Create Task Lists from Master Sheet", "ensureTaskListsFromMasterSheet")
    .addSeparator()
    .addItem("Sync Tasks: Days 1–30", "syncTaskChunk1")
    .addItem("Sync Tasks: Days 31–60", "syncTaskChunk2")
    .addItem("Sync Tasks: Days 61–90", "syncTaskChunk3")
    .addItem("Sync Tasks: Days 91–120", "syncTaskChunk4")
    .addItem("Sync Tasks: Days 121–150", "syncTaskChunk5")
    .addItem("Sync Tasks: Days 151–180", "syncTaskChunk6")
    .addSeparator()
    .addItem("🔁 Sync Task Chunk (Prompt)", "showTaskChunkPrompt");

  // 📅 Calendar Syncing (by chunks or prompted)
  const calendarSyncMenu = ui.createMenu("📅 Calendar Syncing")
    .addItem("Sync Calendar: Days 1–30", "syncCalChunk1")
    .addItem("Sync Calendar: Days 31–60", "syncCalChunk2")
    .addItem("Sync Calendar: Days 61–90", "syncCalChunk3")
    .addItem("Sync Calendar: Days 91–120", "syncCalChunk4")
    .addItem("Sync Calendar: Days 121–150", "syncCalChunk5")
    .addItem("Sync Calendar: Days 151–180", "syncCalChunk6")
    .addSeparator()
    .addItem("🔁 Sync Calendar Chunk (Prompt)", "showCalendarChunkPrompt");

  // 🌀 Phase Management Tools
  const phaseToolsMenu = ui.createMenu("🌀 Phase Tools")
    .addItem("🗓️ Assign Dates to Active Phase", "assignDatesToActivePhase")
    .addItem("➕ Extend Phase by 30 Days", "extendCurrentPhaseBy30Days")
    .addItem("📅 Extend Phase (Prompt)", "showPhaseExtensionPrompt")
    .addSeparator()
    .addItem("🧾 Show Phase Audit Log", "showPhaseAuditLog");

  // 🧼 Cleanup & Data Healing
  const cleanupMenu = ui.createMenu("🧽 Clean-Up")
    .addItem("🧨 Wipe All Task Lists", "clearAllTaskLists")
    .addItem("🗑 Clean Calendar Events", "wipeStudyEventsFromCalendar")
    .addItem("🩺 Heal Session Anchors", "auditAndHealSessionAnchors")
    .addItem("🔁 Refresh Type Tags", "refreshTypeTagColumn")
    .addItem("✂️ Trim Learning Log Rows", "manuallyTrimLearningLog");

  // 🧪 Developer Tools & Diagnostics
  const diagnosticsMenu = ui.createMenu("🧪 Developer & Diagnostics")
    .addItem("📬 Send Weekly Report", "sendWeeklyReport")
    .addSeparator()
    .addItem("🧪 Test Sheet ID Connection", "testSheetIdPopup")
    .addItem("🧪 Test Calendar ID Connection", "testCalendarIdPopup");

  // 📊 Dashboard Management & Reporting
  ui.createMenu("📊 Dashboard")
    .addItem("🧠 Update Dashboard", "updateDashboard")
    .addItem("✨ Format Dashboard", "applyDashboardFormatting")
    .addSubMenu(
      ui.createMenu("🧾 Reporting")
        .addItem("📤 Send Daily Report", "sendDailyReport")
        .addItem("📬 Send Weekly Report", "sendWeeklyReport")
    )
    .addToUi();

  // 🎨 Theme Styling Controls
  ui.createMenu("🎨 Theme Tools")
    .addItem("🌞 Toggle Dashboard Light/Dark Theme", "openThemeSidebar")
    .addItem("🗂 Apply Theme to All Sheets", "applyThemeToAllSheets")
    .addToUi();

  // 🎌 Main Access Menu for All Study Tools
  ui.createMenu("🎌 Study System")
    .addSubMenu(studyLoggers)
    .addSubMenu(taskSyncMenu)
    .addSubMenu(calendarSyncMenu)
    .addSubMenu(phaseToolsMenu)
    .addSubMenu(cleanupMenu)
    .addSubMenu(diagnosticsMenu)
    .addToUi();

  // 🃏 Flashcard Exporter
  ui.createMenu("🃏 Flashcards")
    .addItem("📥 Export Anki CSV (Current Phase)", "exportAnkiCsvSidebar")
    .addToUi();

  // 📘 System Guide Sidebar Access
  ui.createMenu("📘 System Guide")
    .addItem("Open System Overview", "showSystemGuideSidebar")
    .addToUi();

  // ❔ Help Tooltips Menu (Injected via external function)
  addTooltipMenu(); // 🔁 This is a full standalone menu created elsewhere

  // 🧭 Real OAuth Menu for Permission Refresh
  ui.createMenu("🧭 Real OAuth")
    .addItem("🔐 Force Real OAuth", "forceRealOAuth")
    .addToUi();

}

