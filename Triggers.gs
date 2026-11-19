/**
 * ⏰ Triggers.gs
 * ─────────────────────────────────────────────
 * Manages time-based trigger registration for automation flows.
 * 
 * This module defines and sets Google Apps Script triggers that
 * keep the system in sync and updated without manual intervention.
 * 
 * 🔧 Setup Flow:
 * - Clears all existing triggers to avoid duplicates
 * - Registers:
 *   • Realtime sync every 5 minutes
 *   • Weekly report every Sunday @ 8 AM
 *   • Daily streak updater @ midnight
 * 
 * ⚙️ Trigger Targets:
 * - runRealTimeSyncAndUpdate()    → handles periodic sync of progress + tasks
 * - sendWeeklyReport()            → emails dashboard summary every week
 * - refreshDailyStreakMetric()    → writes latest streak to Dashboard daily
 */

// ─────────────────────────────────────────────


/**
 * 🚀 Create and register all automation triggers
 *
 * Clears all current triggers to avoid duplication,
 * then registers system sync, weekly report, and streak refresh.
 */
function setupTriggers() {
  Logger.log("🔄 Setting up automation triggers...");

  // 🔁 Clear all existing triggers first (avoids duplication on redeploy)
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // ⏱ Realtime sync: every 5 minutes
  // ScriptApp.newTrigger("runRealTimeSyncAndUpdate").timeBased().everyMinutes(5).create();

  // 📬 Weekly report: Sunday 8 AM
  ScriptApp.newTrigger("sendWeeklyReport").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(8).create();

 // 📅 Daily streak update: Midnight (0:00)
  // Keeps the dashboard streak current without full refresh
  ScriptApp.newTrigger("refreshDailyStreakMetric").timeBased().everyDays(1).atHour(0).create();
  Logger.log("✅ Automation triggers set up successfully!");
}

