/**
 * 🔁 SyncTools.gs
 * ─────────────────────────────────────────────
 * Handles synchronization between Google Tasks, Google Calendar,
 * and the internal Master Lesson Plan sheet.
 * 
 * 🧠 Ensures real-time task status propagation from Google Tasks → Sheet
 * and provides full reset-resync utility for recovery/debug.
 *
 * 📌 Sync Flow:
 * - updateTaskCompletionInSheets() → Pull completed tasks back into Sheet
 * - runFullCleanupAndResync() → Deletes all synced data and fully recreates
 * 
 * ⚙️ Dependencies: getCurrentPhase(), normalizeTitle(), updateDashboard(),
 *                  deleteAllTasks(), wipeStudyEventsFromCalendar(),
 *                  syncDynamicCurriculumTasks(), syncTasksWithCalendar()
 */

// ─────────────────────────────────────────────


/**
 * 🔄 Sync completed Google Tasks into Master Lesson Plan sheet
 *
 * Matches task titles (normalized) and updates sheet status/date
 * if a task is marked completed on Google Tasks.
 */
function updateTaskCompletionInSheets() {
  const sheet = SpreadsheetApp.openById(getSheetIdFromDashboard()).getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const currentPhase = getCurrentPhase();
  const allTaskLists = Tasks.Tasklists.list().items;
  if (!allTaskLists) return;

  const taskStatusMap = {};

  // 🧭 Build a lookup map: task title (normalized) → status
  allTaskLists.forEach(list => {
    const tasks = Tasks.Tasks.list(list.id, { showCompleted: true, showHidden: true }).items;
    if (tasks) {
      tasks.forEach(task => taskStatusMap[normalizeTitle(task.title)] = task.status);
    }
  });

  // 🔍 Check each row in Master Lesson Plan for update eligibility
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    const phase = sheet.getRange(i, 2).getValue();
    if (phase !== currentPhase) continue;

    const taskTitle = normalizeTitle(sheet.getRange(i, 3).getValue());
    const sheetStatus = sheet.getRange(i, 7).getValue();
    if (!taskTitle) continue;

    // ✅ If marked completed in Tasks but not in Sheets → update it
    if (taskStatusMap[taskTitle] === "completed" && sheetStatus !== "Completed") {
      Logger.log(`✅ Task Completed: ${taskTitle} → Updating in Sheets.`);
      sheet.getRange(i, 7).setValue("Completed");
      sheet.getRange(i, 8).setValue(new Date());
    }
  }

  updateDashboard(); // 🔄 Refresh progress metrics
  Logger.log("🎉 Google Tasks ↔ Sheets Sync Complete!");
}

/**
 * 💣 Full reset and resync of Tasks + Calendar
 *
 * Deletes all Google Tasks and Calendar study events,
 * then re-syncs curriculum tasks and events from scratch.
 */
function runFullCleanupAndResync() {
  Logger.log("🚀 Starting full cleanup and resync...");

  deleteAllTasks(); // 🧨 Remove all Google Tasks
  Logger.log("✅ All tasks deleted.");

  wipeStudyEventsFromCalendar(); // 📆 Remove all synced events
  Logger.log("🗑 Deleted all study-related events.");

  syncDynamicCurriculumTasks(true); // 🛠️ Resync tasks from Master Sheet
  Logger.log("🔄 Synced curriculum tasks.");
  
  syncTasksWithCalendar(); // 📅 Recreate events from synced task dates
  Logger.log("✅ Full cleanup and resync completed!");
}

