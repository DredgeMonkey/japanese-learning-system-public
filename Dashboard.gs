/**
 * 📊 Dashboard.gs
 * ───────────────────────────────────────────── 
 * 🌸 Purpose:
 * Powers the Japanese Study Dashboard by calculating, updating, and formatting real-time learning metrics.
 * This script reads from various sources (log sheets, phase data, sync state) to update key performance indicators (KPIs).
 * 
 * 🔧 Key Functions:
 * - `updateDashboard(caller)`: Core function to update all live metrics in Column B and log executions
 * - `setDashboardMetric()`: Targeted update of a single metric by label
 * - `refreshDailyStreakMetric()`: Trigger-based update for the streak value
 * - `setDashboardNotes()`: Tooltip-style notes and explanations in Column C
 * - `runRealTimeSyncAndUpdate()`: Full system sync + dashboard refresh
 * 
 * 🧠 Notes:
 * - Relies on helper functions from Metrics.gs, Tasks.gs, and Calendar.gs
 * - Metric label names in Column A must match exactly (no extra whitespace)
 * - Updates are non-destructive and visually enhanced with `applyDashboardFormatting()`
 * 
 * 📌 Used by: UI Menu, Phase Tools, Sync Commands
 */

// ─────────────────────────────────────────────

/**
 * 🔄 Update all tracked dashboard metrics in Column B
 * Uses label matching in Column A to update associated values
 *  * Logs each call to the "Debug Log" sheet with optional caller name
 *
 * @param {string} [caller] - Optional identifier of the calling function
 */
function updateDashboard(caller) {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard()); // 📄 Ensure we're pointing to correct instance
  let logSheet = ss.getSheetByName(DEBUG_LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(DEBUG_LOG_SHEET);
    logSheet.getRange("A1:B1").setValues([["Timestamp", "Caller"]]);
  }
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange("A1:B1").setValues([["Timestamp", "Caller"]]);
  }
  logSheet.appendRow([new Date(), caller || ""]);
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dashboard) {
    Logger.log("❌ ERROR: Dashboard sheet not found!");
    return;
  }

  const labels = dashboard.getRange("A2:A" + dashboard.getLastRow()).getValues().flat(); // 📎 All labels to match against
  const metrics = {
    "Days Studied": trackDaysStudied(),
    "Days Studied This Week": trackDaysStudiedThisWeek(),
    "Completion Rate": calculateCompletionRate(),
    "Progress Bar": calculateProgressBar(),
    "Sparkline (Sessions/Day)": generateSparklineString(),
    "Daily Streak": trackStreak(),
    "Longest Streak Ever": trackLongestStreak(),
    "Kanji/Vocab Learned": trackKanjiVocab(),
    "Total Study Hours": calculateStudyTime(),
    "Average Time per Session": calculateAverageSessionTime(),
    "Study Session Count": trackStudySessionCount(),
    "Most Studied Resource": getMostStudiedResource(),
    "Top 3 Tags": getTopThreeTags(),
    "Difficulty Average": getDifficultyAverage(),
    "Last Studied Day": getLastStudiedDay(),
  };

  let updated = 0;

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i]?.toString().trim();
    if (metrics.hasOwnProperty(label)) {
      dashboard.getRange(i + 2, 2).setValue(metrics[label]); // 📌 Column B = Metric values
      Logger.log(`✅ Dashboard updated: "${label}" → ${metrics[label]}`);
      updated++;
    }
  }

  applyDashboardFormatting(); // 🧼 Visual refresh
  setDashboardNotes(); // 🧾 Update explanations

  Logger.log(`🎯 Dashboard update complete. Metrics updated: ${updated}`);
}



/**
 * 🎯 Update a single dashboard metric by name (Column A match)
 * Typically used for real-time value refreshes or manual overrides
 */
function setDashboardMetric(metricLabel, value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  const labels = sheet.getRange("A2:A").getValues().flat();

  for (let i = 0; i < labels.length; i++) {
    if (labels[i] === metricLabel) {
      sheet.getRange(i + 2, 2).setValue(value); // Column B
      return;
    }
  }

  Logger.log(`⚠️ Metric label '${metricLabel}' not found in Dashboard`);
}


/**
 * 🚀 Run full sync (tasks, completion, dashboard update)
 * Used in menu to give a "one-click" refresh of all core data views
 */
function runRealTimeSyncAndUpdate() {
  syncDynamicCurriculumTasks();        // 🧠 Push active tasks
  updateTaskCompletionInSheets();      // ✅ Refresh task checkboxes
  updateDashboard("runRealTimeSyncAndUpdate");      // 📊 Update live dashboard metrics
}


/**
 * 💬 Add tooltip-style notes to Column C next to dashboard metrics
 * Clarifies meaning and logic behind each tracked item
 */
function setDashboardNotes() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET);
  const notes = {
    "Sparkline (Sessions/Day)": "📊 Think of it as a habit radar:\n██████████████ → you're crushing it\n█░█░░█░░░░░███ → tighten up your streak game",
    "Progress Bar": "🧭 Shows % of completed Day #s in the current phase, visualized as 10-block progress bar.",
    "Longest Streak Ever": "🔥 Highest streak ever tracked (1-day grace period allowed).",
    "Daily Streak": "📆 Current streak with 1-day buffer. Updates when dashboard is refreshed.",
    "Difficulty Average": "🎚️ Average difficulty (1–5) across all logged sessions.",
    "Top 3 Tags": "🏷️ Most frequently used tags from Column D in the Learning Log.",
    "Most Studied Resource": "📘 Most frequent content (Column C) you've worked with.",
    "Study Session Count": "✏️ Number of entries with type = 'Study Session'.",
  };

  const labels = dashboard.getRange("A2:A" + dashboard.getLastRow()).getValues().flat();

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i]?.toString().trim();
    if (notes.hasOwnProperty(label)) {
      dashboard.getRange(i + 2, 3).setNote(notes[label]); // 📎 Column C = Notes
    }
  }

  Logger.log("🧼 Dashboard explainer notes refreshed.");
}

/**
 * ⏱ Refresh only the dashboard's "Daily Streak" metric.
 *
 * Called by a nightly trigger so the streak value stays
 * accurate even if the full dashboard isn't updated.
 */
function refreshDailyStreakMetric() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dashboard) {
    Logger.log("❌ Dashboard sheet not found for streak refresh.");
    return;
  }

  const labels = dashboard.getRange("A2:A" + dashboard.getLastRow()).getValues().flat();
  for (let i = 0; i < labels.length; i++) {
    if (labels[i]?.toString().trim() === "Daily Streak") {
      dashboard.getRange(i + 2, 2).setValue(trackStreak());
      Logger.log("✅ Daily streak metric refreshed.");
      return;
    }
  }
  Logger.log("⚠️ 'Daily Streak' label not found in Dashboard.");
}
