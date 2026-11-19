# Security and operational boundaries

## Account access

This project uses Google Apps Script services that may require access to Google Sheets, Tasks, Calendar, and MailApp. Review requested OAuth scopes in your own Google project before approving them. Never publish `.clasp.json`, `.clasprc.json`, tokens, credentials, or copies of actual learning records.

## High-impact maintenance functions

The source intentionally includes broad cleanup/recovery tools. In particular:

- `Tasks.gs`: `deleteAllTasks()` enumerates and deletes tasks from available Google Task lists; `clearAllTaskLists()` removes task lists.
- `Calendar.gs`: `wipeStudyEventsFromCalendar()` searches a broad calendar date range using title/description matching and deletes matching events.
- `SyncTools.gs`: `runFullCleanupAndResync()` invokes destructive cleanup prior to rebuilding sync state.
- `MenuExtras.gs`: `forceRealOAuth()` writes to Learning Log headers and is not a safe general-purpose OAuth reset.

These functions were developed as pragmatic recovery mechanisms. **Do not call them on an account with important unrelated data.** They are not part of initial setup. Test only with dedicated sandbox resources and explicit understanding of the effects.

## Web endpoints

`APIRouter.gs` contains `doGet` routes including `syncTasks`. The source handler does not implement application-level authentication. A web app deployed to run as its owner while allowing anonymous visitors would be unsafe. This public template intentionally has no `webapp` or `executionApi` permission blocks; any deployment must be reviewed and explicitly configured by the deployer.

## No real-world sample data

The source release does not include a published spreadsheet template, personal IDs, secrets, or calendar/task exports. Create and use fictional data in demonstrations.

## Responsible disclosure

If you find a vulnerability, do not file public issues containing secrets, personal Google account identifiers, or private spreadsheet contents. Contact the repository maintainer through a private channel before sharing details publicly. Repository-specific private reporting instructions can be added when GitHub security advisories are enabled.
