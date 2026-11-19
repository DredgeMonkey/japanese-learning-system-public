/**
 * 📁 ChunkUtils.gs
 * ───────────────────────────────────────────── 
 * 📦 Purpose:
 * Provides centralized utilities to calculate and trigger chunked syncs
 * for both Google Tasks and Calendar entries. This avoids code duplication
 * and allows clean numerical chunk targeting (e.g., Days 1–30 = Chunk 1).
 *
 * 🔧 Key Features:
 * - Unified chunk range calculation logic
 * - Abstracted task/calendar sync function wrappers
 * - Prompt-driven manual sync interface
 * - Testing helpers for multi-chunk operations
 *
 * 📊 Each chunk represents 30 sequential days of the curriculum.
 * Designed to integrate with both `syncDynamicCurriculumChunk()` and
 * `syncCalendarEventsByDayRange()` defined elsewhere in the system.
 *
 * ✅ Supports up to 6 chunks (180 days), configurable via `chunkSize`.
 */

// ─────────────────────────────────────────────

/**
 *
 * Returns the start and end day numbers for a given chunk index.
 * Default chunk size is 30 (e.g., Chunk 2 → Days 31–60).
 *
 * @param {number} chunkNum - 1-based chunk index (1–6)
 * @param {number} chunkSize - Optional custom chunk size
 * @returns {{start: number, end: number}} range object
 */
function getChunkRange(chunkNum, chunkSize = 30) {
  const start = (chunkNum - 1) * chunkSize + 1; // 🔹 Inclusive lower bound
  const end = chunkNum * chunkSize;             // 🔹 Inclusive upper bound
  return { start, end };
}

/**
 *
 * Triggers a task sync operation for a specified 30-day chunk.
 * Equivalent to `syncPhaseChunk1to30()` but more dynamic.
 *
 * @param {number} chunkNum - 1-based chunk index
 */
function callTaskChunk(chunkNum) {
  const { start, end } = getChunkRange(chunkNum);
  syncDynamicCurriculumChunk(start, end);
}

/**
 *
 * Triggers a calendar sync operation for a specified 30-day chunk.
 * Equivalent to `syncCalendarEventsChunk1to30()` but cleaner.
 *
 * @param {number} chunkNum - 1-based chunk index
 */
function callCalendarChunk(chunkNum) {
  const { start, end } = getChunkRange(chunkNum);
  syncCalendarEventsByDayRange(start, end);
}


/**
 *
 * Test helper: sequentially runs task sync for chunks 1–3.
 * Can be modified for broader coverage.
 */
function testAllTaskChunks() {
  for (let i = 1; i <= 3; i++) {
    callTaskChunk(i);// 🔹 Replace range if needed (e.g., up to 6)
  }
}

/**
 *
 * Test helper: sequentially runs calendar sync for chunks 1–3.
 * Mirrors the behavior of testAllTaskChunks().
 */
function testAllCalendarChunks() {
  for (let i = 1; i <= 3; i++) {
    callCalendarChunk(i);
  }
}

/**
 *
 * UI prompt allowing users to manually trigger task sync
 * for any 30-day chunk (1–6). Useful for selective backfills or reruns.
 */
function showTaskChunkPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Sync Task Chunk",
    "Enter a number from 1 to 6 to sync that 30-day chunk (e.g., '2' for Days 31–60):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const chunk = parseInt(response.getResponseText());
    if (chunk >= 1 && chunk <= 6) {
      callTaskChunk(chunk); // 🔹 Dynamic trigger
      ui.alert(`✅ Task Chunk ${chunk} (Days ${(chunk - 1) * 30 + 1}–${chunk * 30}) synced.`);
    } else {
      ui.alert("⚠️ Invalid input. Please enter a number between 1 and 6.");
    }
  }
}

/**
 *
 * UI prompt allowing users to manually trigger calendar sync
 * for any 30-day chunk (1–6). Mirrors the task version above.
 */
function showCalendarChunkPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Sync Calendar Chunk",
    "Enter a number from 1 to 6 to sync that 30-day chunk (e.g., '3' for Days 61–90):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const chunk = parseInt(response.getResponseText());
    if (chunk >= 1 && chunk <= 6) {
      callCalendarChunk(chunk); // 🔹 Dynamic trigger
      ui.alert(`✅ Calendar Chunk ${chunk} (Days ${(chunk - 1) * 30 + 1}–${chunk * 30}) synced.`);
    } else {
      ui.alert("⚠️ Invalid input. Please enter a number between 1 and 6.");
    }
  }
}

