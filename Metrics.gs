/**
 * 📊 METRICS MODULE — All dashboard calculations live here
 * ─────────────────────────────────────────────
 * This script is responsible for calculating and formatting study metrics:
 * – Day tracking (total, this week, last studied, etc.)
 * – Streaks (current, longest)
 * – Completion % and progress bars
 * – Study session analytics (time, average, count)
 * – Kanji/Vocab tracking and tag analytics
 * – Sparkline generation for visual summaries
 * 
 * 🔒 PURE MODULE: Do not include any sync, UI, or sheet formatting logic here.
 * Each function should return data only and never mutate sheets.
 * 
 * 📎 Dependencies:
 * – `getPhaseLogSheet()` (assumed global util)
 * – `getSheetIdFromDashboard()` (assumed global util)
 * – Constants: DASHBOARD_SHEET, MASTER_SHEET_NAME
 */

// ─────────────────────────────────────────────

/**
 * Returns today's date formatted in the script time zone.
 * Used for reliable day comparisons across metrics.
 * @return {string}
 */
function getTodayString() {
  const tz = getUserTimeZone();
  return Utilities.formatDate(new Date(), tz, "M/d/yyyy");
}

/**
 * Normalizes a value into a Date at midnight in the spreadsheet time zone.
 * @param {Date|string} raw
 * @return {Date}
 */
function toLocalDate(raw) {
  const tz = getUserTimeZone();
  const rawDate = new Date(raw);
  const yyyy = Utilities.formatDate(rawDate, tz, 'yyyy');
  const MM = Utilities.formatDate(rawDate, tz, 'MM');
  const dd = Utilities.formatDate(rawDate, tz, 'dd');
  const offset = Utilities.formatDate(rawDate, tz, 'XXX');

  const isoString = `${yyyy}-${MM}-${dd}T00:00:00${offset}`;
  return new Date(isoString);
}

/**
 * Calculates the total number of unique days studied in the current phase.
 * Checks for valid numeric values in the 'Day' column.
 * @returns {number}
 */
function trackDaysStudied() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const dayCol = sheet.getRange("E2:E" + sheet.getLastRow()).getValues().flat();
  const uniqueDays = new Set();

  for (const day of dayCol) {
    const num = Number(day);
    if (!isNaN(num) && num > 0) uniqueDays.add(num);
  }

  return uniqueDays.size;
}


/**
 * Tracks how many unique days were studied this week (Sun–Today).
 * @returns {number}
 */
function trackDaysStudiedThisWeek() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const dateCol = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();
  const today = toLocalDate(new Date());

  const sunday = new Date(today);
  sunday.setDate(today.getDate() - today.getDay());

  const uniqueDays = new Set();

  for (let raw of dateCol) {
    if (!raw) continue;
    const parsed = toLocalDate(raw);
    if (isNaN(parsed)) continue;

    parsed.setHours(0, 0, 0, 0);
    if (parsed >= sunday && parsed <= today) uniqueDays.add(parsed.getTime());
  }

  return uniqueDays.size;
}


/**
 * Calculates the current active study streak in days.
 * A one-day grace period is allowed **only** for today's entry. If the most
 * recent logged date is yesterday, the streak continues, but any earlier gap
 */

function trackStreak() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "0 Days";

  const dayCol = sheet.getRange("E2:E" + sheet.getLastRow()).getValues().flat();
  const days = [];

  for (const raw of dayCol) {
    const num = Number(raw);
    if (!isNaN(num) && num > 0) days.push(num);
  }

  const uniqueDays = [...new Set(days)].sort((a, b) => b - a);
  if (uniqueDays.length === 0) return "0 Days";

  Logger.log("\uD83D\uDCC6 Days evaluated for current streak:");
  uniqueDays.forEach(d => Logger.log(`- Day ${d}`));

  const dayMs = 1000 * 60 * 60 * 24;
  let streak = 0;
  let expected = uniqueDays[0];

  for (const day of uniqueDays) {
    if (day === expected) {
      streak++;
      expected--;
    } else {
      break;
    }
  }

  if (streak >= 30) return `🌸 ${streak} Days`;
  if (streak >= 20) return `🌊 ${streak} Days`;
  if (streak >= 10) return `🔥 ${streak} Days`;

  return `${streak} Days`;
}




/**
 * Calculates the longest streak ever achieved during the current phase.
 * Longest streak calculation strictly requires consecutive days. Any missing
 * day resets the count.
 * @returns {string}
 */
function trackLongestStreak() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "0 Days";

  const dayCol = sheet.getRange("E2:E" + sheet.getLastRow()).getValues().flat();
  const days = [];

  for (const raw of dayCol) {
    const num = Number(raw);
    if (!isNaN(num) && num > 0) days.push(num);
  }

  const sorted = [...new Set(days)].sort((a, b) => a - b);
  if (!sorted.length) return "0 Days";

  Logger.log("\uD83D\uDCC6 Days evaluated for longest streak:");
  sorted.forEach(d => Logger.log(`- Day ${d}`));

  let longest = 1, streak = 1;

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] - sorted[i - 1] === 1) {
      streak++;
    } else {
      longest = Math.max(longest, streak);
      streak = 1;
    }
  }

  longest = Math.max(longest, streak);
  if (longest >= 30) return `🌸 ${longest} Days`;
  if (longest >= 20) return `🌊 ${longest} Days`;
  if (longest >= 10) return `🔥 ${longest} Days`;
  return `${longest} Days`;
}



/**
 * Returns raw completion percent for current phase based on unique days logged
 * in Column E (Day) or Column F (Date). Ignores the 'Completed' checkbox.
 * @returns {number}
 */
function getRawCompletionPercent() {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const logSheet = getPhaseLogSheet();
  const lessonSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!dashboard || !logSheet || !lessonSheet || logSheet.getLastRow() < 2) return 0;

  const phaseStr = String(Number(dashboard.getRange("B3").getValue()));
  const lessonData = lessonSheet.getRange(2, 1, lessonSheet.getLastRow() - 1, 3).getValues();
  const phaseLessons = lessonData.filter(r => String(r[1]) === phaseStr);
  const totalPhaseDays = phaseLessons.map(r => String(r[0]));
  if (totalPhaseDays.length === 0) return 0;

  const dateMap = {};
  const tz = getUserTimeZone();
  phaseLessons.forEach(([day, , date]) => {
    if (date instanceof Date) {
      const key = Utilities.formatDate(date, tz, "yyyy-MM-dd");
      dateMap[key] = String(day);
    }
  });

  const logData = logSheet.getRange(2, 5, logSheet.getLastRow() - 1, 2).getValues();
  const completedDays = new Set();

  logData.forEach(([day, date]) => {
    const num = Number(day);
    if (!isNaN(num) && num > 0) {
      completedDays.add(String(num));
    } else if (date instanceof Date) {
      const key = Utilities.formatDate(date, tz, "yyyy-MM-dd");
      if (dateMap[key]) completedDays.add(String(Number(dateMap[key])));
    }
  });

  return completedDays.size
    ? (Array.from(completedDays).filter(d => totalPhaseDays.includes(d)).length / totalPhaseDays.length) * 100
    : 0;
}


/**
 * Wraps raw completion percent in emoji milestone markers.
 * @returns {string}
 */
function calculateCompletionRate() {
  const percent = getRawCompletionPercent();
  const rounded = percent.toFixed(2);

  if (percent >= 90) return `🌸 ${rounded}%`;
  if (percent >= 70) return `🌊 ${rounded}%`;
  if (percent >= 50) return `🔥 ${rounded}%`;
  return `${rounded}%`;
}


/**
 * Returns a text-based progress bar based on phase completion percent.
 * @returns {string}
 */
function calculateProgressBar() {
  const percent = getRawCompletionPercent();
  if (isNaN(percent)) return "░░░░░░░░░░ 0%";

  const filledBars = Math.round(percent / 10);
  return `${"█".repeat(filledBars)}${"░".repeat(10 - filledBars)} ${percent.toFixed(0)}%`;
}


/**
 * Returns total study time in hours from column G (minutes).
 * @returns {string}
 */
function calculateStudyTime() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "0.00 hrs";

  const data = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
  const totalMinutes = data.reduce((sum, [m]) => sum + (typeof m === "number" ? m : 0), 0);
  return `${(totalMinutes / 60).toFixed(2)} hrs`;
}


/**
 * Counts all Kanji and Vocabulary type entries.
 * @returns {string}
 */
function trackKanjiVocab() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "0 entries";

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const count = data.filter(row => ["Kanji", "Vocabulary"].includes(row[1])).length;
  return `${count} ${count === 1 ? "entry" : "entries"}`;
}


/**
 * Counts total Study Session entries only.
 * @returns {number}
 */
function trackStudySessionCount() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  return data.filter(t => String(t).trim().toLowerCase() === 'study session').length;
}


/**
 * Averages session time from valid 'Study Session' entries.
 * @returns {string}
 */
function calculateAverageSessionTime() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "0.00 min";

  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 6).getValues(); // B–G

  let total = 0;
  let count = 0;

  for (const row of data) {
    const type = String(row[0]).trim().toLowerCase(); // Column B = Type
    const time = row[5]; // Column G = Time

    if (type !== "study session") continue;

    if (typeof time === "number" && time > 0) {
      Logger.log(`- Type: ${type}, Time: ${time} min`);
      total += time;
      count++;
    }
  }

  Logger.log(`\n🧮 Total Time: ${total} min`);
  Logger.log(`🧮 Session Count (excluding season logger): ${count}`);
  Logger.log(`🧮 Average Time: ${count ? (total / count).toFixed(2) : 0} min`);

  return count ? `${(total / count).toFixed(2)} min` : "0.00 min";
}


/**
 * Logs all unique types found in column B (Type).
 * Used for debugging tag consistency.
 */
function auditTypesInLog() {
  const sheet = getPhaseLogSheet();
  const types = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues().flat();
  const unique = [...new Set(types.map(t => String(t).trim()))];
  Logger.log("🧾 Unique logged types:\n" + unique.join("\n"));
}


/**
 * Returns most frequent resource logged in column C.
 * @returns {string}
 */
function getMostStudiedResource() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "–";

  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();
  const counts = {};
  data.forEach(val => { if (val) counts[val] = (counts[val] || 0) + 1; });
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  return sorted.length ? `${sorted[0][0]} (${sorted[0][1]})` : "–";
}


/**
 * Returns the top 3 most common tags in column D.
 * @returns {string}
 */
function getTopThreeTags() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "–";

  const data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues().flat();
  const tags = {};

  data.forEach(e => {
    if (!e) return;
    e.split(/[, ]+/).map(t => t.trim().toLowerCase()).forEach(t => {
      if (t) tags[t] = (tags[t] || 0) + 1;
    });
  });

  return Object.entries(tags)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([t, c]) => `${t} (${c})`)
    .join(", ") || "–";
}


/**
 * Returns the average difficulty rating (column I, 1–5 scale).
 * @returns {string}
 */
function getDifficultyAverage() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "–";

  const data = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues().flat();
  const scores = data.map(v => parseInt(v)).filter(n => !isNaN(n) && n >= 1 && n <= 5);

  return scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : "–";
}


/**
 * Gets the highest numbered day studied from column E.
 * @returns {string}
 */
function getLastStudiedDay() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "–";

  const data = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().flat();
  const nums = data.map(d => parseInt(d)).filter(n => !isNaN(n) && n > 0);
  return nums.length ? `Day ${Math.max(...nums)}` : "–";
}


/**
 * Returns 14-day sparkline using █ for studied days and ░ for gaps.
 * @returns {string}
 */
function generateSparklineString() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return "░".repeat(14);

  const data = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();
  const tz = getUserTimeZone();
  const today = toLocalDate(new Date());

  const dates = new Set(data.map(d => {
    const date = toLocalDate(d);
    if (isNaN(date)) return null;
    date.setHours(0, 0, 0, 0);
    return date.toISOString().split("T")[0];
  }).filter(Boolean));

  return Array.from({ length: 14 }, (_, i) => {
    const day = new Date(today);
    day.setDate(today.getDate() - (13 - i));
    const key = Utilities.formatDate(day, tz, "yyyy-MM-dd");
    return dates.has(key) ? "█" : "░";
  }).join("");
}


/**
 * Developer tool: Logs average time calculations with session breakdown.
 */
function debugSessionTimeAverages() {
  const sheet = getPhaseLogSheet();
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); // A–G
  let total = 0, count = 0;

  Logger.log("📊 Sessions contributing to average time:");

  for (const row of data) {
    const type = String(row[1]).trim().toLowerCase(); // Column B = Type
    const time = row[6]; // Column G = Time

    if (type === "season logger") continue;
    if (typeof time === "number" && time > 0) {
      Logger.log(`- Type: ${type}, Time: ${time} min`);
      total += time;
      count++;
    }
  }

  Logger.log(`\n🧮 Total Time: ${total} min`);
  Logger.log(`🧮 Session Count (excluding season logger): ${count}`);
  Logger.log(`🧮 Average Time: ${(total / count).toFixed(2)} min`);
}


/**
 * Developer tool: Logs raw and normalized type values by row.
 */
function debugTypeValues() {
  const sheet = getPhaseLogSheet();
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  const typeSet = new Set();

  data.forEach((v, i) => {
    const norm = String(v).trim().toLowerCase();
    typeSet.add(norm);
    Logger.log(`Row ${i + 2}: Raw="${v}", Normalized="${norm}"`);
  });

  Logger.log("🧾 Unique normalized types:");
  [...typeSet].forEach(t => Logger.log(`- ${t}`));
}

/**
 * Debug helper: logs UTC vs local time and today's date string.
 */
function debugTimezoneInfo() {
  const tz = getUserTimeZone();
  Logger.log(`UTC now: ${new Date().toISOString()}`);
  Logger.log(`Local now (${tz}): ${Utilities.formatDate(new Date(), tz, "M/d/yyyy HH:mm:ss")}`);
  Logger.log(`Today String: ${getTodayString()}`);
}

