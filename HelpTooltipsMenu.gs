/**
 * ❔ HelpTooltipsMenu.gs
 * ─────────────────────────────────────────────
 * Injects the full “Tooltips & Help” menu into the spreadsheet UI.
 * 
 * 📌 Description:
 * - Creates a themed help menu covering all core systems, sync tools, 
 *   loggers, dashboard metrics, email reports, and customization features.
 * - Uses grouped `.addItem()` calls under thematic headings with separators.
 * - Skips gracefully if UI context is unavailable (e.g., on triggers).
 * 
 * 🧠 Related Files:
 * - `HelpTooltips.gs`: Contains all individual tooltip function definitions.
 * 
 * 🛠 Maintainer Notes:
 * - Update here only if new tooltip sections or structure changes are needed.
 * - Tooltip functions should still live in `HelpTooltips.gs` or logical subfiles.
 */



/**
 * ✅ addTooltipMenu()
 * Adds the “Tooltips & Help” dropdown to the spreadsheet UI, organized by theme.
 * Safe to execute even in non-edit contexts — gracefully skips if UI is not available.
 */
function addTooltipMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const tooltipsMenu = ui.createMenu("❔ Tooltips & Help")

  // 🎌 Core System Flow & Setup — foundational system logic
  .addItem("📘 Full System Overview", "showSystemGuideSidebar")
  .addItem("🧠 System Flow Overview", "showSystemInternalsTooltip")
  .addItem("📄 Sheet ID & Phase Settings", "showSheetPhaseHelpTooltip")
  .addItem("📅 Assign Phase Dates", "showDateAssignTooltip")
  .addItem("📈 Extend Phase Dates", "showExtendPhaseDatesTooltip")
  .addItem("ℹ Auto-Creating Phase Logs", "showNewPhaseLogTooltip")
  .addSeparator()

  // 🧾 Task & Calendar Syncing — Google Tasks & Calendar workflows
  .addItem("🧾 Create Task Lists (Required)", "showCreateTaskListTooltip")
  .addItem("📋 Syncing Tasks (Correctly)", "showTaskSyncTooltip")
  .addItem("📅 Syncing Calendar Events", "showCalendarSyncTooltip")
  .addItem("🔁 Syncing by Chunk (Prompt)", "showChunkSyncTooltip")
  .addItem("📦 What Are Chunks?", "showChunkOverviewTooltip")
  .addSeparator()

  // 🧽 Clean-Up Tools — system hygiene and correction flows
  .addItem("🧹 Clean-Up & Resync Workflow", "showCleanupTooltip")
  .addItem("📉 Troubleshooting Metrics", "showDashboardErrorHelpTooltip")
  .addItem("✂️ Trim Log: Tooltip", "showTrimLearningLogTooltip")
  .addSeparator()

  // 📊 Dashboard Insights — help interpreting dashboard metrics
  .addItem("📊 Dashboard Overview", "showDashboardOverviewTooltip")
  .addItem("🔄 When to Update Dashboard", "showDashboardRefreshTooltip")
  .addItem("📈 What’s a Streak?", "showStreakTooltip")
  .addItem("🔥 How Sparkline Works", "showSparklineTooltip")
  .addItem("📊 How Metrics Work", "showMetricsOverviewTooltip")
  .addItem("🈷️ Quick Study Logger", "showQuickStudyLoggerTooltip")
  .addItem("🎚️ Difficulty Ratings", "showDifficultyRatingTipTooltip")
  .addSeparator()

  // 📝 Study Logging & SmartPrompt — interface for input & suggestion tools
  .addItem("📘 What’s a Quick Entry vs Full Session?", "showQuickEntryHelpTooltip")
  .addItem("📝 Using the Study Loggers", "showLoggerTooltip")
  .addItem("✏️ Study Session Logger", "showStudySessionLoggerTooltip")
  .addItem("🈷️ Kanji & Vocab Logger", "showKanjiLoggerTooltip")
  .addItem("🎬 Season Logger", "getTooltip_SeasonLogger") 
  .addItem("📋 Type Tag Reference", "showTooltip_TypeTagReference")    
  .addItem("🎲 SmartPrompt Suggestions", "showSmartPromptTooltip")
  .addSeparator()

  // 📧 Email Reports — summaries, delivery, and troubleshooting
  .addItem("📬 Email Report Types", "showEmailTypeTooltip")
  .addItem("📘 Report Contents", "showEmailContentsTooltip")
  .addItem("📧 Email Troubleshooting", "showEmailTroubleshootingTooltip")
  .addSeparator()

  // 🛠️ Theme & Styling — design customization tips
  .addItem("🎨 Dashboard Styling", "showDashboardStylingTooltip")
  .addItem("🎨 Format Dashboard Layout", "showTooltip_FormatDashboard")
  .addItem("🎨 Sheet Styling System", "showSheetStylingTooltip")
  .addItem("🎨 Theme System", "showThemeTooltip")
  .addSeparator()

  // 🧪 System Maintenance — tech health and diagnostics
  .addItem("🧪 Test Sheet/Calendar Connection", "showTestConnectionTooltip")
  .addItem("🩺 Heal Session Anchors", "showHealTooltip")
  .addItem("🔁 Refresh Type Tags", "showTooltip_RefreshTypeTagColumn")
  .addItem("🗓 UTC Timezone Warning", "showDateWarningTooltip")
  .addItem("🔐 Real OAuth", "showTooltip_RealOAuth")
  .addSeparator()

  // 🔐 Final Safeguards — data integrity and sync protection
  .addItem("💡 Data Integrity & Sync Order Tips", "showSystemTipsTooltip");


    tooltipsMenu.addToUi(); // 🚀 Inject menu into UI
  } catch (e) {
    // ⚠️ Common if script runs in context without UI (e.g., timed triggers)
    Logger.log("⚠️ UI not available in this context. Tooltip menu skipped.");
  }
}

