# First public-source release — published

## Provenance

- Source downloaded read-only from the live Google Apps Script project on 2026-11-19 with CLASP 3.4.1.
- Reconciled against private `DredgeMonkey/japanese-learning-system` application source at `993027af9a344ca8b163a9f50af4f40807bf295e`; the private repo was then updated with its real manifest in commit `1b1f07ad96f73eb4cea5548c0b9d96e63fa7ffbb`.
- Public package contains the **same 34 application source modules**, renamed from CLASP `.js` to `.gs` where appropriate, with the normalized sidebar extension.

## Intentional public-distribution changes

- No private Git history is imported.
- The public `appsscript.json` retains V8 runtime, project time zone and Google Tasks advanced service, but omits `webapp` and `executionApi` blocks to avoid publishing permissive deployment defaults.
- README, LICENSE, SECURITY, architecture, release notes, CONTRIBUTING, and `.gitignore` are new documentation/packaging files.
- Application source logic and inline annotations have not been refactored.

## Limitations

- This is not a public hosted demo or one-click Google Sheet installation.
- No automated Google Tasks/Calendar execution was performed during preparation.
- Copyright holder: **Spencer Armand**. MIT license included.
- GitHub repository visibility is **public**, following Spencer Armand's approval. This source distribution is separate from the private, actively used Apps Script/Sheets application.
