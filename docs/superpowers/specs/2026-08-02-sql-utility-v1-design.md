# SQL Utility Version 1 Design

## Purpose

Evolve the proven SQL connection POC into a portable internal Windows GUI for managing Windows-authenticated SQL Server connections, running restricted read-only queries, paging results, and exporting complete results to formatted Excel workbooks.

Version 1 remains a folder-based application launched through a command file. It requires no installation, administrator access, third-party modules, Microsoft Excel installation, or executable compilation.

## Scope

Version 1 includes:

- A connection-first startup view with saved server/database pairs.
- Windows-authenticated connection testing and workspace entry.
- Installation-directory JSON persistence for saved connections and two global settings.
- A Query tab for one restricted read-only `SELECT` against one table.
- Bounded unordered-result retrieval and 500-row local display pages.
- Deterministic server-side paging for queries containing `ORDER BY`.
- A dependency-free `.xlsx` export with bold headers, filtering, and a frozen header row.
- A Settings tab for the unordered-row limit and Query/Export timeout.

Version 1 excludes:

- Joins, subqueries, CTEs, set operators, multiple SQL statements, stored procedures, and data-changing SQL.
- Query/result/history persistence.
- Credentials or SQL authentication.
- Asynchronous execution, cancellation, progress bars, and background runspaces.
- Multiple result sets, multiple query tabs, and multiple active database connections.
- Automatic total-row counts for ordered queries.
- Multi-worksheet exports for results exceeding Excel's row limit.

## Platform and Runtime Constraints

- Target Windows PowerShell 5.1 and .NET Framework components normally present in the Citrix environment.
- Continue launching through a `.cmd` file that invokes `powershell.exe` with `-NoProfile`, `-STA`, and a process-only execution-policy override.
- Use `System.Windows.Forms`, `System.Drawing`, `System.Data.SqlClient`, `System.IO.Compression`, `System.Xml`, and standard PowerShell JSON commands only.
- Require no installation, PowerShell module registration, NuGet package, Office application, `sqlcmd`, Python runtime, or administrator access.
- Perform no registry or environment-variable writes.
- Perform no network writes except SQL Server connection/query traffic.
- Persistent application configuration may be written only beside the application. User-requested Excel output may be written to the location selected in the Save dialog.

## Runtime File Structure

```text
StartSqlUtility.cmd
SqlUtility.ps1
modules/
  SqlUtility.Config.ps1
  SqlUtility.QueryPolicy.ps1
  SqlUtility.Database.ps1
  SqlUtility.Excel.ps1
```

`StartSqlUtility.cmd` is the only launch entry point a user needs. `SqlUtility.ps1` explicitly dot-sources the four module files relative to `$PSScriptRoot`; it does not modify module search paths or environment variables.

Responsibilities:

- `SqlUtility.ps1`: WinForms controls, staged-view navigation, busy-state coordination, user interaction, result-page binding, paging/status rendering, and the 200-pixel grid-width cap.
- `SqlUtility.Config.ps1`: configuration defaults, schema validation, JSON loading/writing, saved-connection uniqueness, deletion, and settings validation.
- `SqlUtility.QueryPolicy.ps1`: SQL tokenization, version 1 read-only policy enforcement, table-source validation, and top-level `ORDER BY` detection.
- `SqlUtility.Database.ps1`: connection-string construction, diagnostic connection tests, unordered probing, ordered page retrieval, conversion of `SqlDataReader` schema/rows into neutral result objects, callback-based ordered-export row streaming, command timeout enforcement, and deterministic ownership/disposal of SQL resources.
- `SqlUtility.Excel.ps1`: dependency-free streaming `.xlsx` package generation from neutral schema/row input and export-limit enforcement; it never creates or owns SQL connections, commands, or readers.

The POC filenames are replaced by these production names. The proven POC remains recoverable from Git history rather than remaining as a second runnable application in the distribution folder.

## Application Layout and Navigation

The application uses one resizable window with two stages.

### Connection Stage

The connection stage is the first interface on every launch. It contains:

- An initially blank Server field.
- An initially blank Database field.
- **Test Connection** and **Connect** buttons.
- A right-side Saved connections list displaying `server — database`.
- A **Delete** button for the selected saved connection.

Selecting a saved connection fills the two fields but does not connect automatically.

**Test Connection**:

1. Validates that both fields contain non-whitespace text.
2. Opens a short-lived `SqlConnection` with Windows integrated authentication.
3. Executes a fixed read-only diagnostic command and discards its result.
4. Shows a success or failure pop-up.
5. After success, adds the pair to the saved list and attempts to persist it.

**Connect** performs the same validation, database test, and save. After database success it replaces the connection stage with the workspace stage. A configuration-save failure produces a warning but does not block workspace entry.

Saved pairs are unique case-insensitively on both server and database while preserving the casing last entered by the user. **Delete** requires confirmation and updates the file only after a successful write.

### Workspace Stage

The workspace header displays the active `server — database` and a **Change Connection** action. The workspace opens on the Query tab and also contains the Settings tab.

**Change Connection** always asks the user to confirm that the current SQL text and results will be lost. Confirmation clears query text, result data, page state, and export state, then returns to the connection stage. Saved connections and global settings remain.

The application does not keep a database connection open between actions. Connection testing, query execution, page retrieval, and ordered export each open and deterministically dispose their own connection.

## Configuration Persistence

The application uses one file beside `SqlUtility.ps1`:

```text
SqlUtility.config.json
```

Schema:

```json
{
  "schemaVersion": 1,
  "unorderedRowLimit": 1000,
  "queryExportTimeoutSeconds": 120,
  "connections": [
    {
      "server": "server-name",
      "database": "database-name"
    }
  ]
}
```

Rules:

- Missing configuration uses defaults in memory. The file is created only when a connection or settings change is successfully saved.
- `unorderedRowLimit` must be an integer from `100` through `2000`; its default is `1000`.
- `queryExportTimeoutSeconds` must be an integer from `5` through `3600`; its default is `120`.
- The file never contains credentials, query text, results, last selection, export paths, or window state.
- Writes serialize the complete validated configuration to a unique temporary file in the application directory, then replace/move it into place. The temporary file is removed after failure.
- If configuration JSON is malformed, the app shows the error and asks whether to reset it. Reset requires confirmation before overwriting with defaults. Declining reset exits without altering the file.
- An unsupported future schema version produces an error and exits without modifying the file.
- Failure to write configuration is reported in a pop-up. It does not invalidate a successful database connection or query.

`SqlUtility.config.json` and its temporary-file pattern are excluded from Git.

## Settings Tab

The Settings tab is available only in the post-connection workspace. It contains two numeric controls and an explicit **Save Settings** button:

- **Maximum unordered rows** (`100–2000`, default `1000`).
- **Query/Export timeout (seconds)** (`5–3600`, default `120`).

Help text states that the timeout covers interactive query commands and the complete Excel export operation. The 10-second SQL connection timeout remains fixed and separate.

Settings enter application state only after validation and successful configuration writing. A failed write leaves the prior saved values active.

## Query Tab

The Query tab contains:

- A multiline SQL editor at the top.
- **Execute** and **Export to Excel** actions.
- Status text and **Previous** / **Next** page controls.
- A read-only `DataGridView` filling the lower portion.

The displayed page size is fixed at 500 rows and is not a configurable setting.

Editing SQL after successful execution marks the result stale. Stale state disables paging and export until **Execute** is clicked again. Paging and export always use the exact last successfully executed query snapshot, never unexecuted editor changes.

## Version 1 Query Policy

Accepted logical shape:

```sql
SELECT [DISTINCT] columns
FROM [schema.]table
[WHERE ...]
[GROUP BY ...]
[HAVING ...]
[ORDER BY ...]
```

Normal scalar expressions, aliases, `CASE` expressions, aggregate functions, string literals, quoted identifiers, bracketed identifiers, and comments are allowed when they remain within this single-table shape.

The query policy tokenizer masks or recognizes string literals, quoted/bracketed identifiers, line/block comments, and parenthesis depth before classifying keywords. It then enforces:

- Exactly one user statement with only an optional trailing semicolon.
- `SELECT` as the first executable keyword.
- Exactly one top-level `FROM` and one named table source.
- No nested `SELECT` at any parenthesis depth.
- No `JOIN`, `APPLY`, comma-separated table source, CTE, `UNION`, `INTERSECT`, `EXCEPT`, `INTO`, `TOP`, `OFFSET`, or `FETCH`.
- No stored-procedure execution, dynamic SQL, table variables, temporary tables, or external/remote table-source functions.
- No DML, DDL, transaction, permission, or administrative keyword.

Validation returns a structured result containing validity, a user-facing error, normalized executable SQL without the optional trailing semicolon, the single table identifier, and whether a top-level `ORDER BY` exists.

This client policy is an accidental-change safeguard, not a replacement for SQL Server permissions. The signed-in Windows identity remains the authoritative security boundary.

The isolated policy design permits later versions to add explicit `INNER JOIN` and `LEFT JOIN` grammar without changing the UI, database executor, paging model, configuration, or exporter.

## Query Execution and Paging

All commands use the active server/database, Windows integrated security, the configured command timeout, and deterministic disposal. Query errors and timeouts leave the last successful result cleared and display a user-facing error.

### Result Data Boundary

`SqlUtility.Database.ps1` parses database results but does not render them. Interactive execution returns a neutral page-result object containing a `DataTable` plus page metadata: page number, displayed-row count, previous/next availability, completeness, and truncation state. For unordered queries, the database module also returns the bounded complete-or-truncated in-memory result used for local paging.

`SqlUtility.ps1` owns presentation. It binds the returned `DataTable` to the read-only `DataGridView`, updates page controls and status text from the metadata, and recalculates displayed column widths after each bind. No WinForms control is passed into the database module.

For complete ordered export, `SqlUtility.Database.ps1` owns the fresh connection, command, and reader and exposes schema and rows through a callback-based streaming boundary. `SqlUtility.Excel.ps1` consumes that neutral stream and writes workbook cells without directly opening or disposing database resources. Complete unordered export passes the cached schema and rows directly to the Excel module.

### Ordered Queries

For a validated query containing top-level `ORDER BY`:

1. Remove its optional trailing semicolon during normalization.
2. Append application-controlled `OFFSET @Offset ROWS FETCH NEXT @FetchCount ROWS ONLY`.
3. Set `@Offset` to `(pageNumber - 1) * 500` and `@FetchCount` to `501`.
4. Load at most 501 rows.
5. Display the first 500 and use the extra row only to determine whether **Next** is enabled.

**Previous** is enabled after page 1. Page navigation re-executes the query using its executed snapshot and the requested offset. No `COUNT(*)` query or total-page calculation is performed. Status shows `Page N` and the number of displayed rows.

Users are instructed to order by stable, preferably unique columns. The app cannot guarantee deterministic paging when the supplied ordering contains ties or when underlying data changes between page requests.

### Unordered Queries

For a validated query without top-level `ORDER BY`:

1. Execute the original query without rewriting it.
2. Read and retain at most `unorderedRowLimit + 1` rows from the data reader.
3. Close/cancel remaining reader work after the probe boundary.
4. If no extra row appears, mark the retained result complete.
5. If the extra row appears, remove only that extra row, keep the configured maximum, mark the result truncated, and show one notification asking the user to add `ORDER BY` for complete server-side paging.

Retained rows are cached in memory and displayed locally in fixed 500-row pages. For example, a limit of 1000 produces at most two display pages; a limit below 500 produces one partial page.

The cap bounds transferred and retained rows but cannot guarantee that SQL Server avoids expensive scans, aggregation, or sorting required before returning those rows.

### Result Grid Sizing

After every page bind, column widths are calculated from the page's header and displayed values. Each width is capped at 200 pixels. A later page may therefore recalculate widths. Long values remain accessible through normal grid selection/cell display behavior.

## Excel Export

**Export to Excel** uses a standard Save dialog with `.xlsx` as the file type. The selected destination is not persisted.

Export availability:

- Ordered result: enabled when the editor still matches the successfully executed query.
- Complete unordered result: enabled and exports all cached rows.
- Truncated unordered result: disabled because a complete result is unavailable.
- Stale, failed, or never-executed result: disabled.

Ordered export opens a fresh Windows-authenticated connection and re-executes the exact validated query snapshot without application paging. Rows stream directly into the workbook so the complete result is not loaded into memory.

The exporter creates an Open Packaging Convention `.xlsx` using only `System.IO.Compression`, `System.Xml`, and base .NET types. The workbook contains one worksheet named `Results` with:

- A bold header row.
- AutoFilter across the complete used range.
- The first row frozen.
- Numeric, Boolean, and date/time values written using appropriate Excel cell types/styles.
- Database nulls written as empty cells.
- All other values written as strings; a leading `=`, `+`, `-`, or `@` is never emitted as an Excel formula.

Excel permits 1,048,576 total worksheet rows. Version 1 allows at most 1,048,575 data rows plus the header. Encountering another row stops export with an error and leaves no incomplete destination file.

Export writes to a unique temporary file beside the chosen destination. Only a successfully completed and closed package replaces/moves to the requested filename. Failures, row-limit violations, and timeouts close SQL/file resources and remove the temporary file.

The Query/Export timeout is applied to the SQL command and checked as an overall elapsed-time limit during streaming and package writing. Exceeding it cancels database work when possible and aborts the export.

## Synchronous Busy State

Version 1 intentionally uses synchronous work to keep PowerShell implementation simple.

Before connection, query, page, settings-write, delete, or export operations, the app:

- Disables actions that could cause re-entry.
- Sets an informative status message.
- Uses the wait cursor where appropriate.
- Refreshes the form before starting work.

Finally blocks restore controls and cursors after success or failure. The window may be temporarily unresponsive during SQL execution or large file generation; bounded timeouts prevent indefinite database/export waits.

## Error and Confirmation Behavior

- Connection failures: pop-up with the SQL/network/authentication error; do not save or enter the workspace.
- Successful test: success pop-up; separately warn if persistence failed.
- Query-policy failures: concise validation pop-up/status without contacting SQL Server.
- Query/database timeout or failure: clear current results/page/export state and show the error.
- Truncated unordered result: keep the configured maximum visible and show one explanatory pop-up.
- Delete saved connection: require confirmation before writing.
- Change connection: require confirmation that current SQL and results will be lost.
- Configuration corruption: require confirmation before reset; otherwise exit unchanged.
- Configuration-write failure: preserve the prior file and report the failure.
- Export overwrite: use the Save dialog's overwrite confirmation.
- Export failure: preserve any prior destination until the completed temporary workbook can replace it; remove the temporary file.

No operational logs are written in version 1.

## Security and Privacy

- SQL Server authentication uses only the Windows identity of the PowerShell process.
- No username, password, token, or connection-string secret is accepted or persisted.
- `SqlConnectionStringBuilder` constructs all connection strings.
- Application-generated paging values use SQL parameters.
- Only the validated query snapshot reaches the database executor/exporter.
- Query text, results, page caches, and active connection selection exist only in process memory.
- Strings exported to Excel are emitted as literal cell values rather than formulas.
- The application writes only its minimal installation-directory JSON configuration and user-requested Excel output plus same-directory temporary files required for safe replacement.

## Testing Strategy

Tests remain dependency-free PowerShell scripts executed under Windows PowerShell 5.1 with STA where WinForms is involved.

### Configuration tests

- Defaults for missing configuration.
- Valid round-trip serialization.
- Setting range boundaries and invalid types.
- Case-insensitive connection deduplication.
- Confirmed deletion behavior at the configuration boundary.
- Malformed JSON and unsupported schema handling.
- Failed-write behavior leaves the prior configuration intact.

### Query policy tests

- Accepted basic selects, aliases, scalar/aggregate expressions, comments, literals, and optional trailing semicolon.
- Ordered-query detection only at the top level.
- Rejection of every excluded construct and data-changing keyword.
- Keywords inside strings, comments, and quoted identifiers do not produce false classifications.
- Table-source and multi-statement edge cases.
- Future policy extension is possible without database/UI changes.

### Database and paging tests

- Integrated-security connection strings and absence of credentials.
- Connection-test success/failure through injected runners.
- Ordered offset/fetch parameters, 501-row sentinel handling, and page navigation.
- Unordered limit-plus-one probing, complete/truncated state, and 500-row local pages.
- Resource disposal and timeout/error propagation.

### UI state tests

- Connection stage is first and fields start blank.
- Saved selection populates fields without auto-connecting.
- Successful Connect transitions to Query/Settings with Query selected.
- Change-connection confirmation and state clearing.
- Settings visibility/ranges/save behavior.
- Query dirty-state disabling of pager/export.
- Grid read-only behavior and 200-pixel column cap.
- Busy-state restoration after failures.

### Excel tests

- Generate workbooks in test-owned temporary directories.
- Inspect ZIP entries and workbook XML without requiring Excel.
- Verify worksheet name, bold header style, frozen row, AutoFilter range, typed cells, null handling, and formula-injection resistance.
- Verify complete cached unordered export.
- Verify streamed ordered export through an injected row source.
- Verify timeout, worksheet-row-limit, cleanup, and preservation of an existing destination on failure.

### Acceptance tests in Citrix

- Launch through `StartSqlUtility.cmd` from the copied folder.
- Test and persist a real Windows-authenticated server/database connection.
- Reload saved connections and settings from the application directory.
- Run representative unordered complete, unordered truncated, and ordered paged queries.
- Verify paging against a stable unique `ORDER BY`.
- Export and open representative results in Microsoft Excel to confirm formatting, filtering, and the frozen row.
- Confirm no files other than `SqlUtility.config.json`, transient same-directory safe-write files, and explicitly selected Excel outputs are created by the application.

## Explicitly Deferred

- JOIN grammar, beginning with named-table `INNER JOIN` and `LEFT JOIN`.
- Additional functional tabs.
- Background execution, cancellation, and detailed progress.
- Query history, saved queries, and result persistence.
- Total row/page counts.
- Multiple worksheets for Excel-limit overflow.
- Multiple result sets or batch SQL.
- SQL authentication and credential storage.
