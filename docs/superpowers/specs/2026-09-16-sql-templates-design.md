# SQL Template Files

## Approved workflow

Templates are user-owned `.sql` files containing the Query editor's text. Add **Load Template...** and **Save Template...** to the Query tab above the editor. Data Explorer users continue through Send to Query before saving. There is no separate template tab, database binding, metadata refresh, value substitution, or automatic execution.

The application directory contains a `Templates` folder, initially empty. Distribution packages create that empty folder without copying the developer's templates. Startup creates it if missing and preserves all existing contents. If creation fails, report a warning without blocking the application; file dialogs still allow another accessible location.

Every Save and Load dialog starts in this application-local folder, even after browsing elsewhere. Dialogs accept `.sql` files, permit other directories, and do not remember the previously selected template or folder. If the default folder has been removed, recreate it when possible before showing the dialog.

## Text and file handling

- Save the exact editor text, preserving values, comments, placeholders, whitespace, and line endings. Disable Save for a blank/whitespace-only editor. Unfinished SQL is permitted; query validation remains at execution.
- Every save uses a filename dialog. Confirm before replacing an existing `.sql` file. Editing loaded text never modifies its source file automatically.
- Write UTF-8 with a BOM for Windows text-editor compatibility. Read UTF-8 and BOM-marked Unicode files. Reject unreadable/invalidly encoded files and embedded NUL characters before modifying the editor.
- Read a selected file completely, then confirm before replacing nonblank editor text. Cancellation, refusal, or read failure preserves editor text, selected tab, connection, and displayed results. Success selects Query and uses its existing editor-staleness behavior; it never runs SQL.
- Save to a unique temporary sibling, then move a new file or replace an existing file with a sibling backup. Pre-commit failures preserve the destination. Cleanup failures after commit are warnings, not failed-save reports; keep a recoverable backup when removal fails. Temporary/backup files stay in the selected destination directory.
- File creation uses a non-overwriting move when the destination did not exist at confirmation time. A file that appears afterwards must not be silently overwritten.

## Implementation boundaries

Keep the eight runtime files. `SqlUtility.ps1` owns dialogs, controls, service calls, busy-state coordination, and editor replacement. Extend `modules/SqlUtility.Config.ps1` with UI-neutral template-directory and text-file persistence functions, without modifying schema 3 or saving template paths in JSON. Database, QueryPolicy, DataExplorer, and Excel behavior stays unchanged.

The two template buttons use a compact row above the editor so the existing execution/paging toolbar retains its width. Both are disabled while busy; Save also follows editor emptiness. Template files and their folder are user data, excluded from the runtime integrity catalog and source control. Update packaging to include only an empty folder and update the owning documentation's persistence boundaries.

## Verification

Test Unicode/text round trips, extension checks, external directories, folder initialization and preservation, new/existing saves, refused overwrites, locked files, pre-commit preservation, cleanup after failure, and committed saves with cleanup warnings. UI tests cover default paths on repeated dialogs, cancellation, read/write failures, editor-replacement confirmation, no implicit validation/SQL/database work, result staleness, blank/busy action state, and default/minimum-window layout. Package tests verify an empty Templates directory, no personal-template leakage, and unchanged protected runtime files.

Run focused tests, the Windows PowerShell 5.1 aggregate suite, catalog regeneration after CRLF normalization, and Git whitespace/status checks. Citrix dialogs, cloud-drive writes, and live desktop acceptance remain external checks.
