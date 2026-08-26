# SQL Utility Data Explorer Design

## Purpose

Add a constrained, read-only Data Explorer tab for users who know which physical table and filter values they need but want to avoid repeatedly writing routine single-table SQL. The explorer provides on-demand bounded previews, column selection, several `AND` filters, preview-only Excel export, and a safe handoff to the existing Query tab.

The feature remains a query assistant rather than a general database-management surface. It does not edit data, browse views, page arbitrary table data, or replace the Query tab for advanced SQL.

## Approved Scope

Data Explorer provides:

- Automatic loading of visible physical user tables in the active database when the tab is first opened, with a client-side table-name filter and an explicit Refresh action.
- No column-metadata or row query when a table is merely selected.
- An explicit Preview action that retrieves column metadata when needed and returns no more than the configured preview-row limit.
- A scrollable checked column list, selected in full after the first preview, with explicit mouse/Space check changes and buffered prefix navigation as defined by the [Column Selection Navigation Design](2026-08-15-sql-utility-column-selection-navigation-design.md).
- Zero or more structured filters joined only with `AND`. The same column can appear more than once, and a filter column does not need to be selected for output.
- A single preview grid without paging, sorting controls, or an implicit count.
- Export Preview, which exports exactly the currently displayed preview rows.
- Send to Query, which generates editable SQL from the current builder choices, confirms before replacing nonblank Query-tab text, switches to the Query tab, and does not execute.
- A global `previewRowLimit` setting from `10` through `500`, default `100`.
- A global retained-result `resultDataLimitMiB` setting from `128` through `1024`, default `256`, shared with Query pages and unordered caches.

The user explicitly approved adding `modules/SqlUtility.DataExplorer.ps1`, changing the portable runtime distribution from six to seven files.

## Out of Scope

The first release excludes:

- Views, including future view support unless separately approved.
- `OR`, grouped conditions, arbitrary SQL filter expressions, subqueries, joins, and functions in filters.
- Paging, sorting controls, row counts, and complete-result export from Data Explorer.
- Persisting table-specific column selections, filters, filter history, or generated queries.
- Distinct-value lookup queries or value suggestion dropdowns.
- Background execution, cancellation UI beyond existing command timeout behavior, and concurrent database operations.
- Editable cells, inserts, updates, deletes, or any other data modification.

Advanced cases remain in the Query tab. Send to Query is the supported transition from the structured builder to manual SQL.

## User Interface

### Tab layout

The responsive pane hierarchy and scrollbar containment correction is owned by [SQL Utility Data Explorer Layout Correction Design](2026-08-14-sql-utility-data-explorer-layout-correction-design.md).

Data Explorer is a new workspace tab beside Query and Settings. Its layout has:

1. A left table pane with a table-name filter, Refresh action, bare `dbo` table names, and `schema.table` names for other schemas.
2. An upper-right builder area containing the checked output-column list and structured filter rows.
3. A shared toolbar with compact All and None actions on the left and Preview, Export Preview, and Send to Query aligned on the right.
4. A lower-right read-only preview grid and a header that identifies the source table and displayed row count.

The table and column panes scroll vertically. The output-column control's checkbox and typed-prefix behavior is owned by the [Column Selection Navigation Design](2026-08-15-sql-utility-column-selection-navigation-design.md); no instructional label or search box is added. Compact All and None actions share the toolbar row with the preview actions and update output selection without changing the displayed preview. Filter rows remain above the toolbar in their independently scrolling editor.

The preview grid reuses the Query tab's neutral `DataTable` binding, empty-result handling, read-only behavior, cell selection/copy behavior, and 300-pixel column-width cap. It does not reuse Query-tab paging, Count, Export eligibility, or editor-staleness logic.

### Initial table loading

The first activation of Data Explorer synchronously queries the active database for physical user tables visible to the signed-in identity, excluding `sys.tables.is_ms_shipped = 1`. Results are sorted by schema and table name. A table in `dbo` is displayed by its bare table name, while a table in another schema is displayed as `schema.table`; list labels omit SQL identifier brackets. The text filter performs a case-insensitive substring match over these displayed names in memory without another database query. Catalog schema/table metadata remains unchanged and generated SQL remains schema-qualified and safely bracketed.

Refresh re-runs the fixed table-catalog query. If the selected table still exists, it remains selected but its cached metadata, output selection, and filters are reset so the next Preview obtains its current structure. If it no longer exists, the table selection and builder are reset. In either case, the last successful preview snapshot remains available until a later Preview succeeds or the connection changes.

Selecting a table performs no database operation. It resets the current builder's cached metadata, output selection, and filters for that table, but it does not clear or relabel the last successful preview.

### First and later previews

On the first Preview for a selected table:

1. Retrieve the table's column metadata.
2. Populate the checked output-column list with every column selected.
3. Execute an unordered, parameterized preview for all columns and at most `previewRowLimit` rows.
4. Replace the preview snapshot only after the operation succeeds.

Later Preview actions use the current selected output columns and filters. At least one output column is required. A successful Preview atomically replaces the one preview snapshot; the previous table's preview is then lost. A metadata, validation, query, timeout, or result-data-limit failure preserves the prior successful snapshot and its Export Preview availability. A limit failure never displays the partial rows read before the limit was reached and asks the user to review selected columns and filters.

Preview SQL uses `TOP (@PreviewLimit)` with no `ORDER BY`. The preview header and documentation state that rows are unordered and can differ between executions. The SQL command and the client reader both enforce the configured maximum as defense in depth. No sentinel row, completeness probe, paging state, or total count is produced.

Large tables commonly return an unfiltered preview quickly because SQL Server can stop after producing the requested rows, but this is not guaranteed. Unindexed or non-sargable filters, absent matches, wide rows, large values, server load, and storage behavior can still cause scans or timeouts. The existing Query/Export timeout applies.

## Builder and Preview State

Data Explorer maintains two independent in-memory objects:

- **Builder state:** selected table, loaded column metadata, selected output columns, and current filter rows.
- **Preview snapshot:** source table label and the exact `DataTable` returned by the last successful Preview.

Changing table selection, output columns, filters, or preview-row setting does not alter, clear, or mark the displayed preview as stale. No out-of-date status is displayed.

Export Preview always exports the preview snapshot currently shown, even when the builder has since changed or now targets another table. Send to Query always uses current builder state, not the preview snapshot. The source-table label beside the preview prevents ambiguity between the two.

Change Connection clears the table catalog, builder state, preview snapshot, and Data Explorer control state along with the existing transient Query-tab state.

No Data Explorer table choice, metadata, output selection, filter, SQL, or preview data is persisted.

## Columns and Filters

### Column metadata

The Database module obtains neutral metadata from fixed catalog SQL over `sys.tables`, `sys.schemas`, `sys.columns`, and `sys.types`. A requested table is identified from the selected catalog result rather than arbitrary user text. The metadata includes schema name, table name, column name, ordinal, SQL type, maximum length, precision, scale, and nullability as needed for display, validation, and `SqlParameter` construction.

All physical columns remain selectable for output, including types that are not filterable. Column display follows table ordinal order and includes a concise SQL type description.

### Supported filter types and operators

Filter rows contain a column selector, an operator selector, and a value control when the operator requires a value. Filter column choices come from every filterable table column, independently of output-column checks. Duplicate filter columns are allowed and preserve row order in generated SQL.

The supported matrix is:

| SQL type family | Operators |
| --- | --- |
| `char`, `varchar`, `nchar`, `nvarchar` | equals, not equals, contains, starts with; plus is null/is not null when nullable |
| Integer, decimal, money, floating-point | `=`, `<>`, `>`, `>=`, `<`, `<=`; plus is null/is not null when nullable |
| `date`, `time`, `smalldatetime`, `datetime`, `datetime2`, `datetimeoffset` | `=`, `<>`, `>`, `>=`, `<`, `<=`; plus is null/is not null when nullable |
| `bit` | equals, not equals; plus is null/is not null when nullable |
| `uniqueidentifier` | equals, not equals; plus is null/is not null when nullable |

Binary, rowversion/timestamp, XML, spatial, hierarchy, `sql_variant`, CLR/user-defined, and legacy large-object types remain output-only in the first release. Unsupported types do not acquire implicit string conversion for filtering.

Text, numeric, date/time, and identifier input is validated and converted to a typed neutral value before any preview execution. Date/time accepts unambiguous ISO 8601 input and normal Windows-culture input; numeric input accepts normal Windows-culture formatting. Generated editor SQL always renders invariant, unambiguous literals. Bit values use a constrained true/false choice. Null operators omit the value control.

Contains and starts with are parameterized `LIKE` operations. User input is literal data, not a SQL pattern: `%`, `_`, and `[` are converted to SQL Server bracket-escaped pattern text before wildcard construction. Contains adds wildcards on both sides; starts with adds only a trailing wildcard. This avoids adding the optional `ESCAPE` clause to the existing QueryPolicy grammar. SQL Server column collation continues to determine case and accent sensitivity.

## Query Construction and Safety

### Data Explorer module

`modules/SqlUtility.DataExplorer.ps1` is UI-neutral and owns:

- Neutral builder, column, filter, parameter, and preview-command validation.
- The SQL-type-to-operator matrix.
- Identifier quoting for catalog-derived schema, table, and column names.
- Typed filter conversion and literal LIKE-pattern escaping.
- Parameterized preview-command construction.
- Readable Query-tab SQL generation with safe literal serialization.

The module has no WinForms or database access. It does not replace or weaken QueryPolicy.

### Preview command

Preview values never enter SQL as literals. The builder produces a command descriptor containing SQL text and typed parameter descriptors. Its logical shape is:

```sql
SELECT TOP (@PreviewLimit)
    [SelectedColumn1],
    [SelectedColumn2]
FROM [schema].[PhysicalTable]
WHERE [FilterColumn1] = @Filter1
  AND [FilterColumn2] LIKE @Filter2;
```

Identifiers are sourced only from returned metadata and are bracket-quoted with closing brackets doubled. Values and preview limit are `SqlParameter` values with SQL types derived from column metadata; `AddWithValue` is not used. Parameter names are application-generated and unique.

The Database module owns opening the Windows-authenticated connection, adding typed parameters, executing the command, converting the reader to a neutral `DataTable`, enforcing timeout/row limits, cancelling where possible after the bound, and deterministic disposal.

### Query-tab SQL

Send to Query creates an unrestricted single-table `SELECT` with the current output columns and `AND` filters. It deliberately omits `TOP`, paging clauses, and `ORDER BY`, both to provide a useful starting query and to remain within the existing QueryPolicy grammar.

The generated SQL uses safely bracketed identifiers and type-aware literals. String quotes, LIKE pattern metacharacters, binary-looking content, dates, numbers, GUIDs, and null predicates are rendered so the text has the same logical filter meaning as the builder. Generated SQL is readable and multi-line.

If the Query editor contains non-whitespace text, Send to Query shows a confirmation before replacement. Cancellation or generation failure leaves the editor, selected tab, Query results, and Data Explorer state unchanged. On confirmation, the action replaces the editor text and selects Query without executing it. Existing Query-tab `TextChanged` behavior remains authoritative for any displayed Query result.

When the user later executes the generated text, it passes through the unchanged `Test-SqlUtilityQuery` policy and the existing Query workflow like manually entered SQL. The client-side policy remains an accidental-change safeguard rather than a security sandbox; least-privileged SQL Server permissions under the signed-in Windows identity remain authoritative.

## Preview Export

Export Preview is enabled whenever one successful preview snapshot exists and the application is not busy. It is not affected by later builder changes or table selection.

The action prompts for an `.xlsx` destination and passes the exact displayed preview `DataTable` through the existing neutral cached-table row source to `SqlUtility.Excel.ps1`. It never re-runs SQL and never exports more rows than the snapshot contains. User-facing text says Export Preview and reports the bounded row count so it cannot be confused with complete-result Query-tab export.

All existing Excel behavior remains: formula-literal safety, binary hex fidelity, type preservation, row/text limits, timeout checks, unique destination-local transients, prior-destination preservation, and deterministic cleanup. Export failures preserve the preview snapshot and any existing destination.

## Configuration Migration

Configuration schema version 3 adds the shared retained-result data limit to the existing version 2 preview setting:

```json
{
  "schemaVersion": 3,
  "unorderedRowLimit": 1000,
  "queryExportTimeoutSeconds": 120,
  "previewRowLimit": 100,
  "resultDataLimitMiB": 256,
  "connections": []
}
```

`previewRowLimit` is an integer from `10` through `500`, inclusive, default `100`. `resultDataLimitMiB` is an integer from `128` through `1024`, inclusive, default `256`. Both appear in Settings and become active only after the existing safe Save Settings workflow succeeds. Changing either setting does not execute or modify the displayed preview.

Reading schema version 1 supplies `previewRowLimit = 100` and `resultDataLimitMiB = 256`; reading schema version 2 preserves `previewRowLimit` and supplies `resultDataLimitMiB = 256`. Both migrate in memory without creating or rewriting the file. A later successful connection or settings save persists the complete version 3 object through the existing safe-write path. Schema versions greater than 3 remain unsupported; malformed version 3 data remains corruption rather than being silently repaired.

All config-copying and saved-connection operations preserve `previewRowLimit` and `resultDataLimitMiB`. Safe replacement, cleanup-error handling, environmental-read failure behavior, and the restriction against persisted query/editor/result state remain unchanged.

## Architecture and Runtime Distribution

The runtime distribution becomes:

```text
StartSqlUtility.cmd
SqlUtility.ps1
modules/
  SqlUtility.Config.ps1
  SqlUtility.QueryPolicy.ps1
  SqlUtility.DataExplorer.ps1
  SqlUtility.Database.ps1
  SqlUtility.Excel.ps1
```

Responsibilities are:

- `SqlUtility.ps1`: WinForms construction, Data Explorer state and workflow orchestration, shared result rendering, button states, prompts, and user messages.
- `SqlUtility.Config.ps1`: schema version 3 defaults, version 1/2 migration, validation, copying, and safe persistence.
- `SqlUtility.QueryPolicy.ps1`: unchanged ownership of user/editor SQL tokenization, read-only grammar, JOIN policy, normalization, paging/count metadata, and validation.
- `SqlUtility.DataExplorer.ps1`: UI-neutral builder validation and safe preview/editor SQL construction.
- `SqlUtility.Database.ps1`: fixed catalog queries, typed command execution, neutral metadata/results, timeout, cancellation where possible, and SQL-resource disposal.
- `SqlUtility.Excel.ps1`: unchanged neutral `.xlsx` generation and safe destination replacement.

`SqlUtility.ps1` loads the new sibling module relative to `$PSScriptRoot`. No installer, compiled executable, third-party module, NuGet dependency, Office automation, administrator requirement, registry write, or environment-variable write is added. Windows PowerShell 5.1 and .NET Framework remain the runtime.

## Error Handling and Busy Behavior

All Data Explorer database and export actions use the existing synchronous busy model: disable re-entry, show concise status, restore controls and cursor in `finally`, and allow the UI to be temporarily unresponsive. Table catalog, metadata, and preview commands use the configured Query/Export timeout; connection testing remains fixed at 10 seconds.

Validation errors are shown before SQL execution and preserve the preview snapshot. Table-list failure leaves Data Explorer retryable through Refresh. Metadata or preview failure preserves the last successful preview and Export Preview. If metadata was obtained before the row command failed, it remains available for correcting the builder without changing the preview snapshot.

Schema drift between metadata retrieval and execution is reported as a normal preview failure. Refresh or reselecting the table obtains fresh metadata; the application does not attempt speculative automatic repair.

Changing connection clears all Data Explorer transient state after the existing confirmation. No operational log or new persistent diagnostic file is added.

## Testing and Verification

### New focused suite

Add `tests/Test-DataExplorer.ps1` for the UI-neutral module. Cover:

- Every supported and unsupported SQL type family and operator combination.
- Nullable null operators and required/omitted values.
- Multiple `AND` filters, repeated columns, and output/filter independence.
- Zero selected output columns, duplicate selections, catalog-derived unusual identifiers, and bracket quoting.
- Typed conversion boundaries for numeric, date/time, bit, and GUID values.
- Literal `%`, `_`, `[`, `]`, quotes, comment markers, and statement-looking text.
- Parameter descriptors and `TOP (@PreviewLimit)` SQL.
- Editor SQL equivalence, invariant literal formatting, omission of `TOP`/`ORDER BY`, and acceptance by QueryPolicy.

### Existing focused suites

Extend:

- `Test-Config.ps1`: schema 1/2 in-memory migration, schema 3 validation and round trips, inclusive/exclusive setting bounds, copy preservation, environmental/corruption behavior, and safe writes.
- `Test-Database.ps1`: physical-table catalog SQL, physical-table-only behavior, column metadata, typed parameters without `AddWithValue`, bounded sequential preview reads, result-data-limit rejection without partial results, timeout forwarding, executor injection, cancellation where possible, and deterministic disposal.
- `Test-SqlUtilityUi.ps1`: tab loading/refresh, no work on table selection, first/later Preview, all-selected default, native list setup, filter rows, button states, snapshot/builder independence, failure preservation, table changes, connection reset, Send-to-Query confirmation, and exact preview export handoff.
- `Test-Launcher.ps1`: the exact seven-file runtime distribution, relative loading of `SqlUtility.DataExplorer.ps1`, Windows PowerShell 5.1 syntax boundaries, and unchanged launcher behavior from an external working directory.
- `Test-All.ps1`: include the new focused suite.

Excel production behavior does not change. Existing Excel tests remain authoritative; UI tests prove that Export Preview supplies exactly the displayed `DataTable` and invokes no SQL.

Before completion, run the focused suites, the aggregate `Test-All.ps1`, `git diff --check`, and inspect worktree/staged status. Update README, AGENTS, the architecture diagram, portable distribution, configuration schema, workflow, security boundary, tests, and limitations alongside implementation.

## External Acceptance

Local tests do not prove target-environment behavior. External acceptance remains required for:

- Citrix launch with the new seventh runtime file present.
- Automatic/refresh table discovery and metadata visibility under the signed-in Windows identity.
- Preview behavior and timeout on representative large, wide, indexed, and unindexed live tables.
- Generated SQL execution against representative SQL Server data types and collations.
- Configuration version 1/2 migration and version 3 reload from the application directory/cloud drive.
- Result-data-limit rejection for representative `nvarchar(max)` and `varbinary(max)` preview values without replacing the prior preview snapshot.
- Export Preview creation, overwrite behavior, and opening the workbook in desktop Excel.

No live SQL Server, Citrix, cloud-drive, or desktop Excel acceptance is claimed by local verification.
