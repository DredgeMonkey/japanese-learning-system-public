/**
 * 📁 Tasks.gs
 * ─────────────────────────────────────────────
 * Task and calendar syncing logic for the Japanese Learning Tracker.
 * 
 * This script syncs daily study tasks from the Master Lesson Plan to:
 * - 📅 Google Calendar (for daily event visualization)
 * - ✅ Google Tasks (for category-based to-do tracking)
 * 
 * 🔧 Functions:
 * - syncTasksWithCalendar() → Syncs current phase tasks to Calendar
 * - syncDynamicCurriculumTasks() → Smart sync to Google Tasks with duplicate filtering
 * - normalizeTitle() (Utils.gs) → Normalizes titles for comparison
 * - ensureTaskListsFromMasterSheet() → Creates missing task lists from headers
 * - syncDynamicCurriculumChunk(start, end) → Partial range sync (per chunk)
 * - syncTaskChunk#() → Wrapper functions calling chunks (1–6)
 * - deleteAllTasks() → Deletes all tasks across all lists
 * - clearAllTaskLists() → Removes all Google Task lists
 * - createTaskListsFromMasterSheet() → Bulk-creates task lists from headers
 * 
 * 🧠 Depends on:
 * - Master Lesson Plan sheet (A: Day, B: Phase, C: Date, D+: Tasks by category)
 * - Dashboard sheet (Phase = B3, Start Date = B4)
 */

// ─────────────────────────────────────────────


/**
 * 📅 Sync current phase tasks with Google Calendar
 */
function syncTasksWithCalendar() {
  Logger.log("🔄 Syncing current phase tasks with Google Calendar...");

  const sheet = SpreadsheetApp.openById(getSheetIdFromDashboard()).getSheetByName(MASTER_SHEET_NAME);
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard());
  const currentPhase = getCurrentPhase();

  if (!sheet || !calendar) {
    Logger.log("❌ ERROR: Sheet or Calendar not found.");
    return;
  }

  const data = sheet.getDataRange().getValues(); // full range
  const headers = data[0]; // row 1
  const rows = data.slice(1); // from row 2 on

  let createdCount = 0;
  let skippedCount = 0;

  rows.forEach((row, i) => {
    const [dayLabel, phase, date] = row;

    if (String(phase).trim() !== String(currentPhase).trim()) return;
    if (!(date instanceof Date)) {
      Logger.log(`⚠️ Skipping row ${i + 2} — invalid date`);
      return;
    }

    for (let col = 3; col < row.length; col++) {
      const lesson = row[col];
      const category = headers[col];

      if (!lesson || !category) continue;

      const taskTitle = `${dayLabel}: ${lesson}`;
      const existingEvents = calendar.getEventsForDay(date).filter(e => e.getTitle() === taskTitle);

      if (existingEvents.length > 0) {
        Logger.log(`🔁 Event already exists: ${taskTitle}`);
        skippedCount++;
        continue;
      }

      try {
        calendar.createAllDayEvent(taskTitle, date, {
          description: `Japanese Study Task (${category})`,
          timeZone: getUserCalendarTimeZone()
        });
        Logger.log(`📅 Created event: ${taskTitle}`);
        createdCount++;
      } catch (e) {
        Logger.log(`❌ Failed to create event: ${taskTitle} — ${e.message}`);
      }
    }
  });

  Logger.log(`✅ Calendar sync complete for Phase "${currentPhase}". Created: ${createdCount}, Skipped: ${skippedCount}`);
}

/**
 * ✅ Sync current phase tasks to Google Tasks with normalization and duplicate filtering
 */
function syncDynamicCurriculumTasks(force = false) {
  const props = PropertiesService.getScriptProperties();
  const lastRun = props.getProperty("lastSync") || 0;
  const now = Date.now();

  if (!force && now - Number(lastRun) < 1000 * 60 * 5) {
    Logger.log("⏳ Skipping sync — ran too recently.");
    return;
  }
  props.setProperty("lastSync", now);

  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET);

  if (!sheet || !dashboardSheet) {
    Logger.log("❌ Invalid sheet or dashboard reference.");
    return;
  }

  const currentPhase = dashboardSheet.getRange("B3").getValue();
  const phaseStartDate = getPhaseStartDate();
  if (!currentPhase || !phaseStartDate) {
    Logger.log("❌ Missing phase or start date.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).filter(row => row[1] == currentPhase);

  if (rows.length === 0) {
    Logger.log(`❌ No rows found for Phase "${currentPhase}".`);
    return;
  }

  const tasks = Tasks.Tasks;
  const existingLists = Tasks.Tasklists.list().items || [];
  const taskListMap = {};
  existingLists.forEach(list => taskListMap[list.title] = list.id);

  const existingTaskTitlesMap = {};
  for (const [category, listId] of Object.entries(taskListMap)) {
    const existingTasks = Tasks.Tasks.list(listId).items || [];
    existingTaskTitlesMap[listId] = new Set(existingTasks.map(t => normalizeTitle(t.title)));
  }

  let createdCount = 0;
  let skippedCount = 0;
  const createdTitles = new Set();

  let dynamicDelay = 400;
  const delayStep = 100;
  const delayStepInterval = 15;
  const maxDelay = 1500;

  let insertCounter = 0;

  rows.forEach((row, i) => {
    const dayLabel = row[0];
    const dayOffset = i;
    const taskDate = new Date(phaseStartDate);
    taskDate.setDate(phaseStartDate.getDate() + dayOffset);

    for (let col = 3; col < row.length; col++) {
      const taskTitleRaw = row[col];
      const category = headers[col];

      if (!taskTitleRaw || !category) continue;

      const title = `${dayLabel}: ${taskTitleRaw}`; // 🔄 RESTORED original format
      const normalizedTitle = normalizeTitle(title);

      if (!taskListMap[category]) {
        Logger.log(`📂 Creating new task list: "${category}"`);
        const newList = Tasks.Tasklists.insert({ title: category });
        taskListMap[category] = newList.id;
        Logger.log(`✅ Task list created: ${category}`);
      }

      const listId = taskListMap[category];

      if (createdTitles.has(normalizedTitle)) {
        Logger.log(`🛑 Skipped in-session duplicate: ${normalizedTitle}`);
        skippedCount++;
        continue;
      }

      if (existingTaskTitlesMap[listId]?.has(normalizedTitle)) {
        Logger.log(`🔁 Already exists in list: ${normalizedTitle}`);
        skippedCount++;
        continue;
      }

      try {
        Utilities.sleep(dynamicDelay);

        tasks.insert({
          title: title, // 🔄 Use the unmodified display title
          due: taskDate.toISOString()
        }, listId);

        insertCounter++;
        createdTitles.add(normalizedTitle);
        Logger.log(`✅ Created: ${title}`);
        createdCount++;

        if (insertCounter % delayStepInterval === 0 && dynamicDelay < maxDelay) {
          dynamicDelay += delayStep;
        }

      } catch (e) {
        Logger.log(`❌ Failed to insert: ${title} — ${e.message}`);
      }
    }
  });

  Logger.log(`🚀 Task sync complete for Phase "${currentPhase}". Created: ${createdCount}, Skipped: ${skippedCount}`);
}


/**
 * 🆕 Create any missing task lists from Master Lesson Plan headers
 */
function ensureTaskListsFromMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Master Lesson Plan");
  const headers = master.getRange(1, 4, 1, master.getLastColumn() - 3).getValues()[0];

  const existingLists = Tasks.Tasklists.list().items.map(l => l.title);
  headers.forEach(header => {
    if (header && !existingLists.includes(header)) {
      Tasks.Tasklists.insert({ title: header });
      Logger.log(`🆕 Created task list: ${header}`);
    }
  });
}


/**
 * 🔢 Sync partial chunk of tasks by day number range
 */
function syncDynamicCurriculumChunk(startDay, endDay) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const master = ss.getSheetByName("Master Lesson Plan");

  const currentPhase = dashboard.getRange("B3").getValue();
  const phaseStartDate = dashboard.getRange("B4").getValue();

  const lastRow = master.getLastRow();
  const lastCol = master.getLastColumn(); // ✅ Dynamic column count
  const data = master.getRange(2, 1, lastRow - 1, lastCol).getValues(); // ✅ A2 to full width
  const headers = master.getRange(1, 1, 1, lastCol).getValues()[0]; // ✅ Get dynamic headers

  const taskLists = Tasks.Tasklists.list().items || [];
  const taskListMap = {};
  taskLists.forEach(list => taskListMap[list.title] = list.id);

  let createdCount = 0;
  let delay = 400;
  const delayStep = 100;
  const delayStepInterval = 15;
  const maxDelay = 1500;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [dayLabel, phase, date, ...categories] = row;
    const dayNumber = parseInt(dayLabel, 10);

    if (phase != currentPhase || !date) continue;
    if (dayNumber < startDay || dayNumber > endDay) continue;

    for (let j = 0; j < categories.length; j++) {
      const taskTitleRaw = categories[j];
      const category = headers[3 + j]; // ✅ Dynamically align to the correct header

      if (!taskTitleRaw || !category || !taskListMap[category]) continue;

      const title = `Day ${dayNumber}: ${taskTitleRaw}`;
      const task = {
        title: title,
        due: new Date(date).toISOString()
      };

      try {
        Tasks.Tasks.insert(task, taskListMap[category]);
        createdCount++;
        Logger.log(`📌 Created: ${title}`);
        Utilities.sleep(delay);
        if (createdCount % delayStepInterval === 0 && delay < maxDelay) {
          delay += delayStep;
        }
      } catch (e) {
        Logger.log(`⚠️ Failed to create task "${title}": ${e.message}`);
      }
    }
  }

  Logger.log(`✅ Synced Phase ${currentPhase}: Days ${startDay}–${endDay}`);
}


// 🗃️ Modern wrappers for 30-day task chunks
function syncTaskChunk1() { callTaskChunk(1); }
function syncTaskChunk2() { callTaskChunk(2); }
function syncTaskChunk3() { callTaskChunk(3); }
function syncTaskChunk4() { callTaskChunk(4); }
function syncTaskChunk5() { callTaskChunk(5); }
function syncTaskChunk6() { callTaskChunk(6); }


/**
 * 🗑 Delete all tasks across all Google Task lists
 */
function deleteAllTasks() {
    Logger.log("🗑 Deleting all existing tasks...");

    var allTaskLists = Tasks.Tasklists.list().items;
    if (!allTaskLists || allTaskLists.length === 0) {
        Logger.log("⚠️ No task lists found.");
        return;
    }

    allTaskLists.forEach(list => {
        try {
            var tasks = Tasks.Tasks.list(list.id, { showCompleted: true, showHidden: true }).items;
            if (tasks) {
                tasks.forEach(task => {
                    Tasks.Tasks.remove(list.id, task.id);
                    Logger.log(`🗑 Deleted Task: ${task.title}`);
                });
            }
        } catch (error) {
            Logger.log(`⚠️ Error deleting tasks from list '${list.title}': ${error.message}`);
        }
    });

    Logger.log("✅ All tasks deleted. Ready for fresh sync.");
}


/**
 * 🗑 Delete all task lists
 */
function clearAllTaskLists() {
  const taskLists = Tasks.Tasklists.list().items;

  if (!taskLists || taskLists.length === 0) {
    Logger.log("❌ No task lists found.");
    return;
  }

  taskLists.forEach(list => {
    try {
      // Delete each task list
      Tasks.Tasklists.remove(list.id);
      Logger.log(`🗑 Deleted task list: ${list.title}`);
    } catch (error) {
      Logger.log(`⚠️ Error deleting task list '${list.title}': ${error.message}`);
    }
  });

  Logger.log("✅ All task lists deleted.");
}


/**
 * 🆕 Create task lists from Master Lesson Plan sheet headers
 */
function createTaskListsFromMasterSheet() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const masterSheet = ss.getSheetByName("Master Lesson Plan");

  if (!masterSheet) {
    Logger.log("❌ Master Lesson Plan sheet not found.");
    return;
  }

  const headers = masterSheet.getRange(1, 4, 1, masterSheet.getLastColumn() - 3).getValues()[0]; // Column D onward (headers)

  const existingLists = Tasks.Tasklists.list().items.map(list => list.title);
  
  headers.forEach(header => {
    if (header && !existingLists.includes(header)) {
      try {
        // Create the task list
        const newList = Tasks.Tasklists.insert({ title: header });
        Logger.log(`✅ Created task list: ${header}`);
      } catch (error) {
        Logger.log(`⚠️ Error creating task list '${header}': ${error.message}`);
      }
    }
  });
  
  Logger.log("✅ Task lists created from Master Sheet.");
}
