# Contributing

Contributions are welcome in this published source repository. Please keep changes incremental, preserve inline architectural annotations, and explain the affected Google Workspace API constraints. Do not casually replace the established synchronization/cleanup mechanisms without isolated regression tests.

For issues, provide a reproduction using a **test** Google Sheet, Calendar and Tasks list; never post your live account identifiers, tokens or private study records. Changes affecting destructive operations should include safeguards and an explicit demonstration that normal user data remains untouched. Changes to date/streak logic should consider time zones and daylight-saving transitions.

This is a source release, not a guarantee of a fully portable turnkey application. A reusable Google Sheet template and automated tests are welcome future contributions.
