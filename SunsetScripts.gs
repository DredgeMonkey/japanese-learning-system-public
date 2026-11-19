/**
 * 🌅 SunsetScripts.gs
 * ─────────────────────────────────────────────
 * 🛑 Deprecated logic modules for legacy task syncing and metric tracking.
 * 
 * These scripts are retained *only* for historical, developmental, and
 * audit tracing purposes. All logic within this file has been replaced
 * by modularized and safer implementations in active script files.
 * 
 * ❌ DO NOT MODIFY or reintegrate into production workflows.
 * ✅ Use for reference, forensic auditing, or system evolution documentation.
 * 
 * 📜 Contents:
 * - countPendingTasks() → Legacy task counter (replaced by dashboard metric logic)
 * - TASK_CATEGORIES → Obsolete static Tasklist mapping
 * - syncTasksToTaskLists() → Deprecated task creation logic
 * - clearTasksForCurrentPhase() → Legacy bulk task deletion system
 */

// ─────────────────────────────────────────────

/**
 * 🧮 countPendingTasks
 * ─────────────────────────────────────────────
 * Legacy metric to count incomplete rows in the Master Lesson Plan
 * matching the current phase. Replaced by dashboard-native metrics.
 *
 * @returns {number} Count of pending tasks
 */
function countPendingTasks() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard()); // Uses dynamic Dashboard ID
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const currentPhase = getCurrentPhase();

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); // A2:G
  const pending = data.filter(row => row[0] == currentPhase && row[6] !== "Completed");

  return pending.length;
}


// ─────────────────────────────────────────────

/**
 * 🗂️ TASK_CATEGORIES
 * ─────────────────────────────────────────────
 * Static TaskList IDs used for pre-2024 category-based syncing.
 * No longer used. Retained for backward mapping and forensic review.
 */
var TASK_CATEGORIES = {
    "Hiragana & Katakana Mastery": "a2lQUXVsYnR3dDVZbjJLcQ",
    "Kanji Progression": "Q2F4bmVUMERIbjY4UGpFXw",
    "Grammar & Sentence Structure": "WnBjaHRpTzhqTktXTDJMMQ",
    "Listening & Speaking": "bmR1SVFmS085U282Y0ZqcA",
    "Reading Practice": "c1NnNG9mQkJpY3YxMTY1VA",
    "Writing & Expression": "VFRkQTRTTVVtTi11WS00Yw",
    "Real-World Application": "VWM3ZWxoNEJCR01WQ2tobQ",
    "Review & Testing": "R0ljYUJXS1k3d3czU3RXVA"
};


// ─────────────────────────────────────────────

/**
 * 🔄 syncTasksToTaskLists
 * ─────────────────────────────────────────────
 * Legacy function that synced tasks into separate category-based
 * Google Task lists. Replaced by phase-based chunking and unified
 * task batching in createPhaseChunkedTasks().
 */
function syncTasksToTaskLists() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const masterSheet = ss.getSheetByName("Master Lesson Plan");

  if (!masterSheet) {
    Logger.log("❌ Master Lesson Plan sheet not found.");
    return;
  }

  const headers = masterSheet.getRange(1, 4, 1, masterSheet.getLastColumn() - 3).getValues()[0]; // D1 onward
  const data = masterSheet.getDataRange().getValues(); // All rows

  const taskLists = Tasks.Tasklists.list().items;
  const taskListMap = {};
  taskLists.forEach(list => taskListMap[list.title] = list.id);

  let createdCount = 0;
  let skippedCount = 0;

  data.forEach(row => {
    const dayLabel = row[0];
    const phase = row[1];
    const date = row[2];
    
    if (phase && date instanceof Date) {
      for (let col = 3; col < row.length; col++) {
        const lesson = row[col];
        const category = headers[col - 3];

        if (!lesson || !category || !taskListMap[category]) continue;

        const taskTitle = `${dayLabel}: ${lesson}`;
        const listId = taskListMap[category];

        try {
          const task = { title: taskTitle, due: new Date(date).toISOString() };
          Tasks.Tasks.insert(task, listId);
          createdCount++;
          Logger.log(`✅ Created task: ${taskTitle}`);
        } catch (e) {
          Logger.log(`⚠️ Failed to create task: ${taskTitle} — ${e.message}`);
          skippedCount++;
        }
      }
    }
  });

  Logger.log(`✅ Task sync complete. Created: ${createdCount}, Skipped: ${skippedCount}`);
}


// ─────────────────────────────────────────────

/**
 * 🧹 Delete all Google Tasks for the current phase
 */
function clearTasksForCurrentPhase() {
  const currentPhase = getCurrentPhase();
  const sheet = SpreadsheetApp.openById(getSheetIdFromDashboard()).getSheetByName(MASTER_SHEET_NAME);
  const taskLists = Tasks.Tasklists.list().items;

  if (!sheet || !taskLists || sheet.getLastRow() < 2) {
    Logger.log("❌ Error: Sheet or Task Lists not found or empty.");
    return;
  }

  const taskListMap = {};
  taskLists.forEach(list => taskListMap[list.title] = list.id);

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const titlesToDelete = {};

  // ➤ Filter and group tasks to delete by category
  rows.forEach(row => {
    const [dayLabel, phase, date, ...categories] = row;
    if (String(phase).trim() !== String(currentPhase).trim()) return;

    categories.forEach((taskTitleRaw, j) => {
      const headers = sheet.getRange(1, 4 + j, 1, 1).getValues()[0][0];
      const category = headers;

      if (!taskTitleRaw || !category || !taskListMap[category]) return;

      const title = `${dayLabel}: ${taskTitleRaw}`;
      if (!titlesToDelete[category]) titlesToDelete[category] = [];
      titlesToDelete[category].push(title);
    });
  });

  // ➤ Execute deletions with throttle logic
  let deletedCount = 0;
  let delay = 400;
  const delayStep = 100;
  const delayStepInterval = 15;
  const maxDelay = 1500;

  Object.entries(titlesToDelete).forEach(([category, titles]) => {
    const listId = taskListMap[category];
    const existingTasks = Tasks.Tasks.list(listId).items || [];

    existingTasks.forEach(task => {
      if (titles.includes(task.title)) {
        try {
          Tasks.Tasks.remove(listId, task.id);
          deletedCount++;
          Logger.log(`🗑 Deleted: ${task.title}`);
          Utilities.sleep(delay);
          if (deletedCount % delayStepInterval === 0 && delay < maxDelay) delay += delayStep;
        } catch (e) {
          Logger.log(`⚠️ Could not delete '${task.title}': ${e.message}`);
        }
      }
    });
  });

  Logger.log(`✅ Cleared all tasks for Phase '${currentPhase}'`);
}


// 🗃️ Legacy Task Chunk Wrappers (Now Deprecated in favor of callTaskChunk() in ChunkUtils.gs)
// These are preserved for rollback or historical reference

/*
function syncPhaseChunk1to30() {
  syncDynamicCurriculumChunk(1, 30);
}

function syncPhaseChunk31to60() {
  syncDynamicCurriculumChunk(31, 60);
}

function syncPhaseChunk61to90() {
  syncDynamicCurriculumChunk(61, 90);
}

function syncPhaseChunk91to120() {
  syncDynamicCurriculumChunk(91, 120);
}

function syncPhaseChunk121to150() {
  syncDynamicCurriculumChunk(121, 150);
}

function syncPhaseChunk151to180() {
  syncDynamicCurriculumChunk(151, 180);
}
*/

