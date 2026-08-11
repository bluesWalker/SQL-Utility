# SQL Utility GUI Modernization and On-Demand Row Count Design

## Purpose

Improve SQL Utility's appearance and query-workspace ergonomics while preserving its portable Windows PowerShell 5.1 and dependency-free WinForms runtime. Add an explicit exact-row-count action without imposing count-query cost on every execution.

The work remains a focused extension of the existing application. It does not change authentication, query grammar, paging size, export semantics, configuration schema, runtime packaging, or the current draggable query/results splitter.

## Approved Scope

This change includes:

- A restrained native WinForms visual refresh.
- Segoe UI at 9 points for the application interface.
- Consolas at 10 points for the SQL editor.
- A single responsive query-action row.
- Shorter paging and export button labels.
- A 300-pixel displayed-grid column-width cap instead of 200 pixels.
- More informative row-status text using counts already known by the application.
- An explicit **Count** action for exact result-row counts when the total is otherwise unknown.

This change excludes:

- Dark mode or theme selection.
- External images, icon packs, fonts, modules, packages, or other dependencies.
- Owner-drawn controls or a custom widget framework.
- Window-size, splitter-position, or theme persistence.
- Automatic count queries during normal query execution or page navigation.
- Total-page calculation or navigation to an arbitrary page.
- Changes to the accepted SQL grammar.
- Background execution, progress reporting, or cancellation controls beyond the existing synchronous timeout and best-effort command cancellation boundaries.

## Runtime and Portability Constraints

Preserve the existing runtime contract:

- Launch only through `StartSqlUtility.cmd`.
- Run under Windows PowerShell 5.1 with `-NoLogo -NoProfile -STA -ExecutionPolicy Bypass`.
- Use only .NET Framework and Windows components expected in the target Citrix environment.
- Keep Windows integrated authentication as the only authentication mode.
- Add no installer, compiled executable, third-party dependency, Office automation, `sqlcmd`, Python runtime, registry write, environment-variable write, or administrator requirement.
- Keep the six-file runtime distribution unchanged.

## Visual Direction

Use the approved restrained native refresh rather than a custom theme.

- Enable Windows visual styles before constructing the application form.
- Set the form-level interface font to Segoe UI, 9 point, so normal child controls inherit it.
- Keep the SQL editor's explicit Consolas, 10 point font.
- Use normal Windows light surfaces and system colors for general controls.
- Use restrained blue emphasis for the primary **Execute** action and normal native treatment for secondary actions.
- Give the results grid a white data surface, quiet neutral column headers, standard readable selection colors, and light grid lines.
- Preserve native rendering for controls whose modernization would otherwise require owner drawing, including the tab control.
- Preserve DPI scaling, the current minimum form size, keyboard focus behavior, and tab navigation.

The refresh applies to the complete application, including the connection stage, query workspace, Settings tab, status area, and dialogs. It must not change the meaning, availability, or confirmation boundary of any existing action.

## Query Workspace Layout

The horizontal `SplitContainer` remains draggable and keeps its current initial distance and panel minimum sizes. No splitter behavior or persistence is added.

Replace the two-row, fixed-position query action area with one docked responsive layout. Its left-to-right logical order is:

```text
Execute | ORDER BY required for paging. | flexible space | page status | Count | < | > | Export
```

Layout rules:

- **Execute** remains left-aligned and is the primary action.
- The paging notice immediately follows **Execute** and reads exactly `ORDER BY required for paging.`
- Flexible space absorbs horizontal resizing.
- Page status and the remaining four buttons remain right-aligned.
- **Count** appears immediately after page status.
- `<`, `>`, and **Export** replace **Previous**, **Next**, and **Export to Excel** respectively.
- Accessible names or tooltips retain the full meanings `Previous page`, `Next page`, and `Export to Excel` for the shortened controls.
- The row remains single-line at the current minimum window size and under supported DPI scaling. Secondary text may ellipsize before controls overlap or wrap.
- The reduced action-row height gives the editor or results area the vertical space no longer needed by the second control row.

The action order is presentation-only. Existing execution, paging, export, stale-result, and busy-state behavior remains authoritative.

## Grid Column Sizing

After each page bind, continue to size every generated result column from its displayed header and cell contents. Then:

1. Cap the calculated width at 300 pixels.
2. Return the column to a fixed `None` autosize mode so the cap remains effective.
3. Preserve user horizontal scrolling for wider result sets.

No minimum-width, wrapping, manual resize, column reorder, or persistence behavior is added.

## Row Status Without Additional SQL

Use information already present in the neutral page result before offering an exact count query.

- A complete unordered cache reports the exact cached total, for example `Page 1 - 500 of 723`.
- A truncated unordered cache reports a lower bound based on the retained limit, for example `Page 1 - 500 of 1000+`.
- An ordered page with no explicit count reports only the current page and displayed rows, for example `Page 1 - 500`.
- A complete unordered empty result, or an empty first ordered page with no previous page, reports `Page 1 - 0 of 0` without issuing a count query. An empty later ordered page caused by concurrent data changes does not imply a zero total.

Local unordered paging updates the displayed-row component while retaining the exact or lower-bound total. Ordered paging retains an explicit total only after the user has requested and successfully completed **Count**.

## Explicit Exact Count

### Availability

**Count** never runs implicitly. It is enabled only when all of the following are true:

- The form is not busy.
- A successful result is current rather than stale.
- The application has an executed-query snapshot and a current page result.
- The total is not already known exactly from a complete unordered cache.

It remains available after an ordered or truncated-unordered count succeeds so the user may explicitly refresh a point-in-time total. Editing the SQL disables **Count** with paging and export. A new execution, query failure, or confirmed connection change clears the prior explicit total.

### Query Semantics

The count must represent the number of rows produced by the exact last successfully validated query, not the number of physical rows in its source table.

- Preserve `WHERE`, `DISTINCT`, `GROUP BY`, and `HAVING` semantics.
- Remove only the validated top-level `ORDER BY`, because ordering cannot change result cardinality and may add unnecessary sort work.
- Use `COUNT_BIG` so totals are not limited to the `int` range.
- Generate count SQL from tokenizer/parser boundaries established during successful policy validation. Do not use regex or unvalidated string replacement.
- Make the generated count query alias-safe for every accepted projection shape, including wildcards, unnamed expressions, duplicate output names, scalar expressions, aggregates, and multiple selected expressions.
- Do not broaden or reinterpret the user-facing query grammar. The application-generated wrapper is an internal execution detail applied only after the original query passes the existing policy.

Validation exposes an order-free `CountSourceSql`. After a successful page reveals the actual output-column count, QueryPolicy finalizes the alias-safe `CountSql`, which the UI stores in the last-successful-query snapshot. This supports wildcard, unnamed, and duplicate result columns while ensuring **Count**, paging, and export continue to use the exact last successful execution rather than current editor text.

### Database Execution

The Database module owns exact-count execution through a scalar SQL command:

- Construct the connection string through the existing integrated-security helper.
- Open a fresh connection for the explicit count action.
- Apply the configured Query/Export timeout.
- Convert the scalar `COUNT_BIG` result to a nonnegative 64-bit integer.
- Attempt command cancellation where possible on timeout or interruption.
- Deterministically dispose the connection and command on success and failure.

The count operation is synchronous, matching the existing v1 runtime model. The busy state prevents re-entry and shows a concise `Counting rows...` status while it runs.

### Success and Failure

On success:

- Store the point-in-time total separately from the current page data.
- Update page status to include the exact total, for example `Page 1 - 500 of 123,456`.
- Preserve the current grid, page number, pager availability, export eligibility, and executed-query snapshot.

On failure or timeout:

- Preserve the displayed grid and every valid query, paging, and export state.
- Preserve any prior successful explicit total rather than replacing it with an error value.
- Restore the form's busy state and cursor in `finally` behavior.
- Report the count failure separately from query execution so it cannot be mistaken for a lost result.
- Allow the user to retry **Count**.

Because the count uses a separate SQL execution, concurrent database changes may cause it to differ from rows fetched earlier or later. The UI and documentation treat it as a point-in-time total, not a transactionally frozen companion to every page.

## State Model

Add explicit total-count state without changing configuration persistence:

- The executed-query snapshot gains generated count SQL.
- Transient workspace state gains an optional 64-bit explicit total.
- The neutral page result continues to own displayed rows, cached rows, page number, previous/next sentinels, and complete/truncated flags.
- Known unordered totals come from `CachedData.Rows.Count`; they are not copied into configuration.
- No count, query, result, layout, font, theme, or window state is persisted.

All reset paths must clear the new state consistently. Existing stale-result behavior continues to leave displayed rows visible while disabling follow-up actions.

## Module Boundaries

### `SqlUtility.ps1`

- Initialize visual styles.
- Construct and style WinForms controls.
- Own the one-row action layout and transient explicit-count state.
- Coordinate the injected count service, busy state, messages, and page-status rendering.
- Preserve all existing connection, query, paging, export, settings, and stale-result workflows.

### `modules/SqlUtility.QueryPolicy.ps1`

- Continue owning tokenization, grammar validation, normalization, and top-level clause detection.
- Generate alias-safe count SQL only for a successfully validated query.
- Return no executable count SQL for rejected input.

### `modules/SqlUtility.Database.ps1`

- Execute the generated scalar count query with integrated authentication, timeout enforcement, validation of the scalar result, cancellation where possible, and deterministic disposal.
- Remain free of WinForms dependencies.

### Unchanged modules

- `SqlUtility.Config.ps1` receives no schema or persistence change.
- `SqlUtility.Excel.ps1` receives no workbook or export change.
- `StartSqlUtility.cmd` receives no launcher change.

## Documentation

Update `README.md` with:

- Segoe UI and native visual-style behavior.
- The one-row toolbar and shortened button meanings.
- The 300-pixel grid column cap.
- Known, lower-bound, and explicit point-in-time row-count behavior.
- The cost and separate-execution boundary of **Count**.
- The continued absence of automatic counts or total-page queries.
- Updated architecture responsibility text and external acceptance notes.

The approved v1 design remains historical context. This document owns the approved extension and supersedes its 200-pixel cap and blanket exclusion of total-row counting only to the extent described here. Automatic counting remains excluded.

## Verification Strategy

Use regression-first test-driven development under Windows PowerShell 5.1.

### Query policy tests

- Generated count SQL is empty for every rejected query.
- Simple, filtered, ordered, distinct, grouped, and having queries preserve result-row cardinality semantics.
- Only top-level `ORDER BY` is removed.
- Strings, comments, quoted identifiers, bracketed identifiers, nested parentheses, and ORDER-like text cannot alter the generated boundary.
- Wildcards, unnamed expressions, duplicate aliases/names, and multi-expression projections remain alias-safe.
- The accepted/rejected grammar does not widen.

### Database tests

- Exact count uses integrated-security connection strings and the supplied count SQL.
- Timeout reaches the scalar command.
- `COUNT_BIG` values above the `int` range are preserved.
- Null, negative, malformed, timeout, and executor-failure results are rejected clearly.
- Connections and commands are disposed on every path.

### UI tests

- Application and editor fonts have the approved families and sizes.
- The action row is single-line and follows the approved logical/control order.
- Short labels and full accessible meanings are present.
- Paging notice text is exact.
- Page status renders complete, lower-bound, ordered-unknown, zero, and explicit-total forms.
- **Count** enablement follows busy, stale, no-result, complete-cache, ordered, and truncated states.
- Count success updates only the total/status state.
- Count failure preserves grid, page, pager, export, snapshot, and prior successful total.
- Query execution, query failure, editor changes, and connection changes clear or disable count state correctly.
- Grid columns never exceed 300 pixels and remain in fixed autosize mode after binding.
- Existing query, paging, export, configuration, and busy-state tests remain green.

### Completion checks

Run focused policy, database, and UI suites during development, then run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

## External Acceptance

Local fake services and in-memory tests cannot establish real-environment behavior. The following remain pending until actually performed:

- Appearance, DPI scaling, one-line fit, and control accessibility in the target Citrix environment.
- Live SQL Server count correctness for representative simple, distinct, grouped, and large ordered queries.
- Count-query performance and timeout behavior against production-scale data and indexes.
- Continued live paging and export behavior after the UI changes.
- Opening exported workbooks in desktop Excel.

## Acceptance Criteria

The change is acceptable when:

- The application uses the approved restrained native appearance with Segoe UI and a Consolas editor.
- The query action area is one responsive line with the exact approved order and shortened labels.
- Result columns cap at 300 pixels.
- Status text uses already-known totals without additional SQL.
- Exact total counting occurs only after the user selects **Count**.
- Count SQL preserves the validated result-set cardinality for every accepted query shape and removes only top-level ordering.
- Count failures never destroy valid results or follow-up state.
- No portability, security, persistence, paging, export, grammar, or packaging invariant regresses.
- Focused and aggregate tests pass from the feature worktree.
- Citrix and live SQL Server checks are reported as pending unless actually completed.
