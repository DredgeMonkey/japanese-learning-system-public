/**
 * 📅 Calendar.gs
 * ─────────────────────────────────────────────  
 * 🌸 Purpose:
 * Integrates study sessions with the user's Google Calendar using dynamic calendar ID from the dashboard.
 * Enables creation, syncing, and deletion of all-day events tied to Japanese study tasks.
 * 
 * 🔧 Key Functions:
 * - Create or update calendar events with duplicate prevention
 * - Bulk sync events by chunked day ranges
 * - Wipe or selectively remove study-related events
 * - Maintain timezone accuracy across operations
 * 
 * 🧠 Notes:
 * - Chunk-based syncing helps avoid API throttling
 * - Calendar ID is dynamically pulled from the Dashboard (B28)
 * - Safe-guarded against non-date entries and invalid values
 * 
 * 📌 Used by: Task and calendar sync tools in the Dashboard menu
 */

// ─────────────────────────────────────────────

/**
 * 🔁 Get user's default calendar time zone
 * Dynamically pulls the calendar’s configured time zone via ID from Dashboard (B28),
 * falling back to the script time zone if unavailable.
 */
function getUserCalendarTimeZone() {
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard()); // ✅ Dynamic
  return calendar ? calendar.getTimeZone() : getUserTimeZone();
}


/**
 * 📅 Create an all-day calendar event (with duplicate prevention)
 * 
 * Checks for an existing event with the same title on the same day and deletes it before creating a new one.
 * Ensures the event is timezone-accurate and labeled as a Japanese study task.
 */
function upsertStudyEvent(taskTitle, taskDate) {
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard()); // ✅ PATCHED
  const timeZone = getUserCalendarTimeZone();
  const eventDate = convertToTimeZone(new Date(taskDate), timeZone);

  if (isNaN(eventDate)) {
    Logger.log(`⚠️ Invalid date for task: ${taskTitle}`);
    return;
  }

  // 🧽 Clean up duplicates
  const existing = calendar.getEventsForDay(eventDate).filter(e => e.getTitle() === taskTitle);
  existing.forEach(e => e.deleteEvent());
  
  // 🆕 Create new event
  calendar.createAllDayEvent(taskTitle, eventDate, {
    description: "Japanese Study Task",
    timeZone: timeZone
  });

  Logger.log(`✅ Created event: ${taskTitle} on ${eventDate.toDateString()}`);
}

/**
 * 🧼 Wipe all events in date range (brute force cleanup)
 * 
 * Deletes all events between a preset date range that match any study-related keywords in title or description.
 * Used before re-syncing or resetting the system.
 */
function wipeStudyEventsFromCalendar() {
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard()); // ✅ PATCHED
  const startDate = new Date("2024-01-01");
  const endDate = new Date("2025-12-31");
  const events = calendar.getEvents(startDate, endDate);

  let deleted = 0;
  for (const event of events) {
    const title = event.getTitle();
    const desc = event.getDescription();
    const shouldDelete = title.match(/Day|Anki|Genki|WaniKani|Study|Shadowing|Kana|Pimsleur|Shujinkou|Tae Kim/i) || desc.includes("Japanese Study Task");
    if (shouldDelete) {
      Utilities.sleep(300); // 🕒 Delay for rate limiting
      event.deleteEvent();
      deleted++;
    }
  }

  Logger.log(`🗑 Deleted ${deleted} study-related events.`);
}


/**
 * 🧽 Remove study event for a specific title and day
 * 
 * Deletes only the matching event on the specified date if found.
 */
function removeStudySession(taskTitle, taskDate) {
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard()); // ✅ PATCHED
  const eventDate = new Date(taskDate);
  const events = calendar.getEventsForDay(eventDate);
  const filtered = events.filter(e => e.getTitle() === taskTitle);

  if (!filtered.length) {
    Logger.log(`⚠️ No match found for ${taskTitle} on ${taskDate}`);
    return;
  }

  filtered.forEach(e => e.deleteEvent());
  Logger.log(`🧼 Removed event: ${taskTitle} on ${taskDate}`);
}


/**
 * 🧠 Helper: Convert date to all-day, timezone-accurate
 * 
 * Ensures consistent all-day formatting across time zones using ISO string normalization.
 */
function convertToTimeZone(date, timeZone) {
  const formatted = Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
  return new Date(formatted + "T00:00:00");
}


// 🔁 Calendar Chunk Sync Routing Shortcuts (UI Menu-Driven)
function syncCalChunk1() { callCalendarChunk(1); }
function syncCalChunk2() { callCalendarChunk(2); }
function syncCalChunk3() { callCalendarChunk(3); }
function syncCalChunk4() { callCalendarChunk(4); }
function syncCalChunk5() { callCalendarChunk(5); }
function syncCalChunk6() { callCalendarChunk(6); }


/**
 * 📅 Shared helper function: Sync events by day number range
 * 
 * Iterates through the Master Lesson Plan sheet, syncing calendar events for each lesson
 * within the specified day range, ensuring events match the current phase only.
 * 
 * Features incremental delay throttling to prevent rate-limit errors.
 */
function syncCalendarEventsByDayRange(startDay, endDay) {
  Logger.log(`📅 Syncing calendar events for days ${startDay}–${endDay}...`);

  const sheet = SpreadsheetApp.openById(getSheetIdFromDashboard()).getSheetByName(MASTER_SHEET_NAME);
  const calendar = CalendarApp.getCalendarById(getCalendarIdFromDashboard());
  const currentPhase = getCurrentPhase();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  let created = 0;
  let skipped = 0;
  let delay = 300;
  const delayStep = 100;
  const delayInterval = 15;
  const maxDelay = 1500;

  for (const row of rows) {
    const [dayLabel, phase, date, ...lessons] = row;
    const dayNum = parseInt(String(dayLabel).replace(/\D/g, ""), 10);

    // 🔍 Phase & range filtering
    if (String(phase).trim() !== String(currentPhase).trim()) continue;
    if (!(date instanceof Date)) continue;
    if (dayNum < startDay || dayNum > endDay) continue;

    for (let i = 0; i < lessons.length; i++) {
      const lesson = lessons[i];
      const category = headers[i + 3]; // offset headers for lesson columns

      if (!lesson || !category) continue;

      const title = `${dayLabel}: ${lesson}`;
      const existing = calendar.getEventsForDay(date).filter(e => e.getTitle() === title);

      if (existing.length > 0) {
        Logger.log(`🔁 Already exists: ${title}`);
        skipped++;
        continue;
      }

      try {
        calendar.createAllDayEvent(title, date, {
          description: `Japanese Study Task (${category})`,
          timeZone: getUserCalendarTimeZone()
        });

        Logger.log(`📅 Created: ${title}`);
        created++;
        Utilities.sleep(delay);
        if (created % delayInterval === 0 && delay < maxDelay) delay += delayStep;

      } catch (e) {
        Logger.log(`❌ Failed to create event '${title}': ${e.message}`);
      }
    }
  }

  Logger.log(`✅ Calendar Sync Complete for Days ${startDay}–${endDay}. Created: ${created}, Skipped: ${skipped}`);
}


