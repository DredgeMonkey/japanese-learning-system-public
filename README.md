# Japanese Learning System 🇯🇵

An extensively documented Google Sheets + Apps Script application for structured Japanese-language study, learning journals, and optional Google Workspace synchronization.

**Project status:** Existing, actively used personal application. This public repository is a **source release** derived from a verified production snapshot, not a hosted demo or a ready-to-copy spreadsheet template. Live user records, credentials, and account-specific configuration are not included.

## What it does

- **Study planning:** phase-based curriculum and a Master Lesson Plan.
- **Learning log:** study sessions, kanji/vocabulary entries, and long-form media/season immersion.
- **Dashboard:** streaks, logged time, study-session counts, completion and activity metrics.
- **HTML interfaces:** custom Google Sheets menus, logger sidebars, study prompts, help panels and theme selection.
- **Anki-compatible export:** generate vocabulary CSV for manual import into Anki.
- **Optional Google Workspace integrations:** chunk-oriented Google Tasks / Calendar sync and email reports.
- **Recovery tools:** explicitly documented maintenance utilities developed in response to Apps Script and Google API constraints.

## Architecture

The spreadsheet is the primary user interface and data store. Apps Script runs the business logic and HTML sidebars; Google Tasks, Google Calendar, and MailApp are optional external integrations. The implementation includes extensive module- and function-level comments, bilingual help text, and preserved historical code.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) and [docs/RELEASE_NOTES.md](docs/RELEASE_NOTES.md).

## Source tree

- `Code.gs`, `Menu*.gs`, `UI*.gs`, `Dashboard.gs` — initialization, menus and dashboard.
- `LoggerTools.gs`, `SeasonLog.gs`, `Metrics.gs`, `Sheets.gs`, `SheetUtils.gs` — study logging and metrics.
- `PhaseTools.gs`, `Tasks.gs`, `Calendar.gs`, `ChunkUtils.gs`, `SyncTools.gs`, `Triggers.gs` — lesson planning, optional synchronization and scheduling.
- `AnkiExport.gs`, `SmartPrompt.gs`, `Theme.gs`, `HelpTooltips*.gs` — export, prompts, styling and guidance.
- `*.html` — custom Sheets interfaces and help sidebars.
- `SunsetScripts.gs` — clearly marked historical/deprecated routines.
- `appsscript.json` — Apps Script V8 runtime and Google Tasks advanced service.

## Using or adapting the source

This is **not yet a one-click installation**. It was developed as a working script bound to a structured Google spreadsheet. A sanitized, redistributable spreadsheet template and step-by-step installation guide have not yet been validated or included.

To adapt it safely, study the expected tab names and Dashboard configuration documented in source, create a **separate test spreadsheet and Apps Script project**, configure your own integrations, and test each capability before granting production Google-account access. Do not point a test copy at your everyday Calendar or Tasks lists.

**Do not run the cleanup/reset functions against an account containing unrelated Tasks or Calendar records.** Some maintenance tools intentionally perform broad deletions and are intended only for controlled recovery in an environment where that scope is understood. See [SECURITY.md](SECURITY.md).

## Important deployment notice

This public source template deliberately **omits** the original production project's `webapp` and `executionApi` manifest deployment settings. Source file `APIRouter.gs` includes GET handlers that can initiate external synchronization. **Do not deploy it as an anonymously accessible web app without first implementing appropriate access controls and reviewing every route.** The live private implementation has not been changed.

## Privacy and data flows

All learning records are stored in the user's Google spreadsheet; however optional integrations **do** transfer relevant data to Google Tasks, Google Calendar, and MailApp email recipients specified in configuration. Anki CSV is exported for user-controlled import. Review Google's OAuth permission prompts before authorizing a copy.

No live account identifiers, credentials, spreadsheet contents, or OAuth configuration are intentionally included in this public source release. Please report potential exposures privately using [SECURITY.md](SECURITY.md).

## Project history and credit

**Created and developed by Spencer Armand.** This release preserves the original modular structure, detailed source annotations, and engineering decisions of the working Japanese Learning System. Many workarounds are deliberate accommodations to Google Apps Script/Workspace API constraints, not invitations to remove functioning behavior without testing.

## License

MIT; see [LICENSE](LICENSE). Copyright (c) 2025–2026 Spencer Armand.
