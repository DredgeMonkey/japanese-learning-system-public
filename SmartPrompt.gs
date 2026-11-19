/**
 * 💡 SmartPrompt.gs
 * ─────────────────────────────────────────────
 * Generates randomized daily prompts for study inspiration
 * based on current study day and selected user mood.
 * 
 * • Pulls data from SmartPrompts sheet (rows: day range, mood, prompt)
 * • Mood matching supports "Any", "Low", "Normal", or "High"
 * • Used by SmartPrompt Sidebar UI
 *
 * 📋 Data Columns:
 * - Column A: Day Range (e.g. "1–30")
 * - Column B: Mood (Any, Low, Normal, High)
 * - Column E: Prompt Text
 * 
 * ⚙️ Dependencies: getCurrentStudyDay(), getSheetIdFromDashboard()
 */

// ─────────────────────────────────────────────


/**
 * 🎯 Get prompt matching current day number and mood
 *
 * @param {number} dayNum – Current study day (1–90)
 * @param {string} mood – Selected mood ("Low", "Normal", "High", or "Any")
 * @returns {string} – A random matching prompt, or fallback if none
 */
function getSmartPromptForDay(dayNum, mood) {
  const ss = SpreadsheetApp.openById(getSheetIdFromDashboard());
  const sheet = ss.getSheetByName("SmartPrompts");

  // → Exit early if SmartPrompts sheet is missing or empty
  if (!sheet || sheet.getLastRow() < 2) return "No prompt found for today’s mood and day range.";

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const matching = [];

  for (const [rangeStr, promptMood, , , prompt] of data) {
    if (!rangeStr || !prompt) continue;

    const [minStr, maxStr] = rangeStr.split("–").map(s => parseInt(s.trim()));
    const inRange = dayNum >= minStr && dayNum <= maxStr;
    const moodMatch = promptMood === "Any" || promptMood === mood;

    // ✅ If both range and mood match, add to candidate list
    if (inRange && moodMatch) matching.push(prompt);
  }

  // 🎲 Randomly select from matching prompts
  return matching.length ? matching[Math.floor(Math.random() * matching.length)] : "No prompt found for today’s mood and day range.";
}

/**
 * 📆 Get SmartPrompt for today's study day
 *
 * @param {string} mood – Selected mood for filtering
 * @returns {string} – Matched SmartPrompt or error fallback
 */
function getTodaySmartPrompt(mood) {
  const dayNum = getCurrentStudyDay();

  // → Handle edge case: day not set
  if (!dayNum) return "Could not determine today’s study day.";
  return getSmartPromptForDay(dayNum, mood);
}

