/**
 * 📁 Reporting.gs
 * ─────────────────────────────────────────────
 * Email reporting utilities for the Japanese Learning Tracker.
 * 
 * This module provides automated or manual delivery of daily and weekly
 * progress summaries based on dynamic dashboard metrics.
 * 
 * 📨 Functions:
 * - sendWeeklyReport() → Sends full summary of all dashboard KPIs.
 * - sendDailyReport()  → Sends concise motivational daily snapshot.
 * - getDashboardMetric(label) → Utility for label-based dashboard lookups.
 * 
 * 📤 Data Source: Dashboard sheet (labels in Column A, values in Column B)
 * 🧠 Dependency: updateDashboard() must be called before reporting
 * 📌 Email Address: Pulled from Dashboard cell B25
 */

// ─────────────────────────────────────────────

/**
 * 🔍 Look up a value from Column B by matching a label in Column A
 *
 * @param {string} label - Text label to search for (e.g., "Days Studied")
 * @returns {*} The value found in Column B, or a dash if not found
 */
function getDashboardMetric(label) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const labels = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat();

  for (let i = 0; i < labels.length; i++) {
    if (labels[i]?.toString().trim() === label) {
      return sheet.getRange(i + 2, 2).getValue(); // Column B = value
    }
  }

  Logger.log(`⚠️ Metric "${label}" not found in Dashboard.`);
  return "—";
}


/**
 * 💌 Send a full weekly progress email with all key metrics
 * Pulls fresh data from the Dashboard after triggering updateDashboard()
 */
function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const email = dashboard.getRange("B25").getValue();

  if (!email || !email.includes("@")) {
    Logger.log("❌ Invalid or missing email address in Dashboard!B25.");
    return;
  }

  updateDashboard("sendWeeklyReport"); // Ensure all values are up to date

  // 📊 Fetch values cleanly using labels
  const metrics = {
    "Days Studied": getDashboardMetric("Days Studied"),
    "Days Studied This Week": getDashboardMetric("Days Studied This Week"),
    "Current Streak": getDashboardMetric("Daily Streak"),
    "Longest Streak": getDashboardMetric("Longest Streak Ever"),
    "Kanji/Vocab Learned": getDashboardMetric("Kanji/Vocab Learned"),
    "Study Session Count": getDashboardMetric("Study Session Count"),
    "Total Study Hours": getDashboardMetric("Total Study Hours"),
    "Average Time per Session": getDashboardMetric("Average Time per Session"),
    "Top Tags": getDashboardMetric("Top 3 Tags"),
    "Most Studied Resource": getDashboardMetric("Most Studied Resource"),
    "Difficulty Average": getDashboardMetric("Difficulty Average"),
  };

  const subject = "📘 Your Weekly Japanese Progress Report"; // 💌 Compose report
  const body = `
こんにちは！

Here’s your weekly progress summary:

📅 Days Studied: ${metrics["Days Studied"]}
📈 This Week: ${metrics["Days Studied This Week"]}
🔥 Streak: ${metrics["Current Streak"]} | 🏆 Longest: ${metrics["Longest Streak"]}
🧠 Kanji/Vocab: ${metrics["Kanji/Vocab Learned"]}
✏️ Sessions: ${metrics["Study Session Count"]} | ⏱️ Total Time: ${metrics["Total Study Hours"]}
⏳ Avg Session: ${metrics["Average Time per Session"]}
📚 Most Used: ${metrics["Most Studied Resource"]}
🏷️ Top Tags: ${metrics["Top Tags"]}
🎚️ Avg Difficulty: ${metrics["Difficulty Average"]}

Keep going strong. Every day counts 💪

— Japanese Learning Tracker ✨
`;

  MailApp.sendEmail(email, subject, body);
  Logger.log("✅ Weekly report sent to: " + email);
}


/**
 * 💬 Send a short daily summary email with core stats
 * Intended to give encouragement and show incremental progress
 */
function sendDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const email = dashboard.getRange("B25").getValue();

  if (!email || !email.includes("@")) {
    Logger.log("❌ Invalid or missing email address in Dashboard!B25.");
    return;
  }

  updateDashboard("sendDailyReport"); // Pull fresh metrics

  const metrics = {
    "Current Streak": getDashboardMetric("Daily Streak"),
    "Kanji/Vocab Learned": getDashboardMetric("Kanji/Vocab Learned"),
    "Total Study Hours": getDashboardMetric("Total Study Hours"),
    "Average Time per Session": getDashboardMetric("Average Time per Session"),
    "Most Studied Resource": getDashboardMetric("Most Studied Resource"),
    "Top Tags": getDashboardMetric("Top 3 Tags"),
  };

  const subject = "🌅 Your Daily Japanese Learning Snapshot";
  const body = `
おはようございます！

Here’s your latest learning snapshot:

🔥 Streak: ${metrics["Current Streak"]}
🧠 Kanji/Vocab Learned: ${metrics["Kanji/Vocab Learned"]}
⏱️ Total Time: ${metrics["Total Study Hours"]} | ⏳ Avg: ${metrics["Average Time per Session"]}
📚 Most Used: ${metrics["Most Studied Resource"]}
🏷️ Top Tags: ${metrics["Top Tags"]}

Stay focused and steady — progress is made one day at a time 🛤️

— Japanese Learning Tracker ✨
`;

  MailApp.sendEmail(email, subject, body);
  Logger.log("📬 Daily report sent to: " + email);
}

