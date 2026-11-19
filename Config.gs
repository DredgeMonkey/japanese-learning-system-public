/**
 * ⚙️ Config.gs
 * ─────────────────────────────────────────────  
 * 🌸 Purpose:
 * Stores global configuration constants related to UI theming and layout preferences.
 * Enables centralized control of theme behavior and styling across the Dashboard interface.
 * 
 * 🔧 Key Elements:
 * - Location of the selected theme value in the Dashboard sheet
 * - Predefined light and dark themes for consistent UI application
 * 
 * 🧠 Notes:
 * - The theme selection is read from Dashboard cell B30 (`THEME_KEY_CELL`)
 * - Each theme defines custom colors for interface zones and font readability
 * - `applyDashboardFormatting()` in Theme.gs reads this configuration
 * 
 * 📌 Used by: Theme.gs, ThemeSidebar.html
 */

// ─────────────────────────────────────────────

/**
 * 🎨 Cell in the Dashboard where the current theme (Light/Dark) is stored
 * Used as the key for retrieving the user's selected visual theme
 */
const THEME_KEY_CELL = "B30"; // 🔑 Holds "Light" or "Dark" as selected by user

/**
 * 📑 Primary sheet name for the Master Lesson Plan
 */
const MASTER_SHEET_NAME = "Master Lesson Plan";

/**
 * 🧶 Theme configuration object
 * Defines UI color zones, font styling, and dashboard title emoji/text
 * Each sub-object maps directly to theming logic in Theme.gs
 */
const THEMES = {
  Light: {
    headerBg: "#e8f0fe",        // 🔷 Soft blue for headers
    phaseZone: "#fef7e0",       // 🟡 Light yellow for phase boxes
    emailZone: "#fff9db",       // 🟤 Slightly darker yellow for email fields
    idZone: "#fff3cd",          // 🟠 Light amber for Sheet/Calendar ID zones
    fontColor: "#000000",       // 🖤 Standard black for text
    metricHighlight: "#d9ead3", // 🟢 Pastel green for key metrics
    dashboardTitle: "🌸 Japanese Learning Dashboard 🌸"
  },
  Dark: {
    headerBg: "#1f1f1f",        // ⚫ Deep gray for dark mode headers
    phaseZone: "#2a2a2a",       // ⚫ Slightly lighter for phase background
    emailZone: "#2e2e2e",       // ⚫ Mid-gray for email input zone
    idZone: "#3a3a3a",          // ⚫ Lightest gray for IDs to distinguish zones
    fontColor: "#eaeaea",       // 🤍 Soft white text for readability
    metricHighlight: "#4caf50", // ✅ Vivid green for metric cells (WCAG friendly)
    dashboardTitle: "🌙 Japanese Learning Dashboard 🌙"
  }

};

