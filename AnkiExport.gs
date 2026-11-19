/**
 * 📁 AnkiExport.gs
 * ───────────────────────────────────────────── 
 * 📦 Purpose:
 * Provides an interactive sidebar for exporting vocabulary entries as CSV files
 * compatible with Anki flashcard decks. Automatically formats content with
 * "Romaji | Kana | English" extraction and ensures clean download via browser.
 *
 * 🧠 Behavior:
 * - Filters entries marked as "Vocabulary" from the active phase's Learning Log
 * - Parses entries using delimiter pipe `|` format
 * - Outputs a CSV file with proper escaping and quoting
 * - Generates a sidebar with a downloadable CSV blob for browser-based download
 *
 * ✅ Dependencies:
 * - `getPhaseLogSheet()` to access the correct phase
 * - `getCurrentPhase()` to label the export
 *
 * 🔐 Notes:
 * - Avoids Drive or external APIs by using direct HTML5 blob download
 * - Only exports entries with valid triple-pipe format
 *
 * 🧪 Output Example:
 *   "Romaji","Kana","English","Tags"
 *   "arigatou","ありがとう","thank you","greetings"
 */

// ─────────────────────────────────────────────

/**
 *
 * Opens a sidebar allowing the user to download a CSV-formatted file of
 * Vocabulary entries for the current phase, formatted for Anki import.
 *
 * 🧮 Format: "Romaji | Kana | English"
 * ✅ Filters: Only Vocabulary-type entries with valid triple-segment format
 * 📦 Output: CSV with quoted fields and embedded download button
 * 
 * Alerts the user if no valid entries are found.
 */
function exportAnkiCsvSidebar() {
  const sheet = getPhaseLogSheet();
  if (!sheet) return SpreadsheetApp.getUi().alert("❌ Could not access current phase learning log."); // Guard clause

  const data = sheet.getDataRange().getValues();
  const csvRows = [["Romaji", "Kana", "English", "Tags"]]; // CSV header row
  const currentPhase = getCurrentPhase(); // Used for labeling download

  for (let i = 1; i < data.length; i++) { // Skip header
    const type = data[i][1];             // Column B: Type
    const content = data[i][2];          // Column C: Resource (entry with pipes)
    const tags = data[i][3];             // Column D: Tags

    if (type !== "Vocabulary" || !content.includes("|")) continue; // 🔹 Skip non-Vocab or malformed

    const parts = content.split("|").map(p => p.trim());
    if (parts.length !== 3) continue; // Only accept proper triple pipe entries

    const [romaji, kana, english] = parts;
    csvRows.push([romaji, kana, english, tags || ""]); // Add to output, tags optional
  }

  if (csvRows.length <= 1) {
    SpreadsheetApp.getUi().alert("⚠️ No valid vocabulary entries found for export.");
    return;
  }

  // Escape quotes and format CSV string
  const csvContent = csvRows.map(row =>
    row.map(val => `"${val.replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const template = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_blank">
        <style>
          body { font-family: Segoe UI, sans-serif; padding: 20px; }
          button { padding: 8px 12px; font-size: 14px; }
        </style>
      </head>
      <body>
        <h2>📥 Anki CSV Export</h2>
        <p>Your deck for <strong>Phase ${currentPhase}</strong> is ready.</p>
        <button onclick="downloadCsv()">⬇ Download CSV</button>
        <script>
          function downloadCsv() {
            const csv = ${JSON.stringify(csvContent)};
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = "anki_export_phase_${currentPhase}.csv";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
          }
        </script>
      </body>
    </html>
  `).setTitle("📥 Anki Export").setWidth(400);

  SpreadsheetApp.getUi().showSidebar(template);
}

