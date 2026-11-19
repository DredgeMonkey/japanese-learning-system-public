/**
 * 🌐 APIRouter.gs
 * ─────────────────────────────────────────────
 * External entry point for HTTP GET requests via Web App.
 * 
 * 📌 Supports the following `trigger` parameters:
 * - syncTasks:            ➤ Runs `syncDynamicCurriculumTasks()`
 * - cleanCalendar:        ➤ Runs `cleanCalendarEvents()`
 * - openResourceTracker:  ➤ Opens Resource Guide sheet
 * - openKanjiLogger:      ➤ Opens Kanji Logger sidebar
 * 
 * 🎯 Used for lightweight automation or external triggering (e.g., bookmarks, apps).
 * Returns plain text response confirming action.
 */

// ─────────────────────────────────────────────

/**
 * 🔁 HTTP GET Handler — Routes based on trigger param
 *
 * @param {Object} e - Event parameter from HTTP GET
 * @returns {TextOutput} Plain text output for response
 */
function doGet(e) {
  const trigger = e.parameter.trigger;
  if (!trigger) return ContentService.createTextOutput("❌ No trigger specified");

  try {
    switch (trigger) {
      case "syncTasks":
        syncDynamicCurriculumTasks(); // 🔄 Sync tasks to Google Tasks
        return ContentService.createTextOutput("✅ syncDynamicCurriculumTasks triggered");

      case "cleanCalendar":
        cleanCalendarEvents(); // 🧹 Clear study-related calendar events
        return ContentService.createTextOutput("✅ cleanCalendarEvents triggered");

      case "openResourceTracker":
        openResourceTracker(); // 📖 Open the Resource Guide sheet
        return ContentService.createTextOutput("✅ Resource Tracker sheet opened");

      case "openKanjiLogger":
        openKanjiLoggerSidebar(); // ✍️ Open Kanji/Vocab logger UI
        return ContentService.createTextOutput("✅ Kanji Logger sidebar opened");

      default:
        return ContentService.createTextOutput("❌ Unknown action: " + trigger);
    }
  } catch (err) {
    return ContentService.createTextOutput("❌ Error: " + err.message); // ⚠️ Return error message
  }
}

