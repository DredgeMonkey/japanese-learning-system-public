# System architecture

## Application shape

```text
User in Google Sheets
  ├─ Dashboard / Master Lesson Plan / Learning Log / Resource Guide
  ├─ Custom menus and HTML sidebars
  │    ├─ Study-session / kanji-vocab / season loggers
  │    ├─ SmartPrompt and theme chooser
  │    └─ Bilingual Help and System Guide
  └─ Apps Script (V8)
       ├─ Sheets, phase and logger services
       ├─ Metrics and dashboards (including streak calculations)
       ├─ Anki CSV export (user-initiated)
       ├─ Google Tasks and Calendar synchronization (optional)
       ├─ MailApp reporting (optional)
       └─ Triggers, recovery, debugging and historical utilities
```

## Module responsibilities

- **Entry / UI:** `Code.gs`, `Menu.gs`, `MenuExtras.gs`, `UI.gs`, `UIExtended.gs`, HTML sidebars.
- **Learning records:** `LoggerTools.gs`, `SeasonLog.gs`, `Sheets.gs`, `SheetUtils.gs`.
- **Curriculum:** `PhaseTools.gs`, `ChunkUtils.gs`, `SmartPrompt.gs`.
- **Reporting:** `Dashboard.gs`, `Metrics.gs`, `Reporting.gs`.
- **Integrations:** `Tasks.gs`, `Calendar.gs`, `SyncTools.gs`, `APIRouter.gs`, `Triggers.gs`.
- **Operations and historical context:** `Healing.gs`, `SunsetScripts.gs`, `HelpTooltips.gs`, `HelpTooltipsMenu.gs`.

## Engineering constraints and decisions

The source extensively annotates the evolution of chunked/batch synchronization, date and timezone handling, and maintenance operations. Those functions were developed in response to real Apps Script execution and Workspace API constraints. Their presence is intentional; a public fork should preserve behavior until new implementations are tested in an isolated environment.

## Public-release boundaries

1. The source snapshot is equivalent to the private production source in nonblank code lines, but this distribution intentionally removes permissive deployment settings from `appsscript.json`.
2. No live production spreadsheet records, identity, secrets, Google authorization data, or deployment history is included.
3. Versioned deployments, installed triggers, permissions and live Google Sheet data are distinct from checked-in source; this repository does not deploy automatically.
4. The spreadsheet template is **not** in this release; source alone is not turnkey.
5. The code still includes deliberately broad recovery functions. Review `SECURITY.md` before invoking them.
