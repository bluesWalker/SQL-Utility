# GUI Modernization and Explicit Row Count Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Apply the approved restrained native WinForms modernization, raise the result-column width cap to 300 pixels, consolidate the query toolbar, improve page status text, and add an explicit exact row-count action without changing query safety, paging, export, packaging, or authentication behavior.

**Architecture:** Keep layout, transient count state, enablement, and messages in `SqlUtility.ps1`; keep count-source derivation and alias-safe `COUNT_BIG` SQL generation in `modules/SqlUtility.QueryPolicy.ps1`; keep scalar SQL execution and deterministic ADO.NET cleanup in `modules/SqlUtility.Database.ps1`. The count query is never automatic. Validation produces an order-free count source, and the policy layer finalizes the derived-table wrapper after a successful result exposes the actual output-column count, which is required to handle wildcard, duplicate, and unnamed projections safely.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework WinForms/System.Drawing, `System.Data.SqlClient`, existing script-based regression tests, Git.

## Global Constraints

- Preserve `StartSqlUtility.cmd`, Windows PowerShell 5.1, `-STA`, Windows integrated authentication, and the six-file runtime distribution.
- Add no dependency, installer, compiled executable, registry write, environment-variable change, SQL authentication, or persistent UI state.
- Do not change the approved query grammar, 500-row page size, unordered limit behavior, ordered `OFFSET`/`FETCH` behavior, export completeness rules, or splitter behavior.
- Never execute `COUNT_BIG` automatically. Use the existing Query/Export timeout and a fresh SQL connection.
- Treat the count as a point-in-time value for the exact last successful normalized query snapshot. Editor changes make it stale; a new execution, failed execution, or connection change clears it.
- A count failure must leave the grid, pager, export state, current result, and any previously displayed exact count unchanged.
- Preserve deterministic disposal of WinForms resources and SQL connections/commands.
- Use regression-first TDD: add a failing assertion, run the focused suite and confirm the intended failure, implement the smallest change, then rerun the focused suite.
- Do not stage, commit, merge, push, or remove the worktree without first reviewing status and diff. Each task below proposes a scoped commit; obtain the user's integration choice after all verification passes.

---

## Interface and File Map

| File | Planned contract change |
| --- | --- |
| `modules/SqlUtility.QueryPolicy.ps1` | Add `CountSourceSql` to validation results and add `New-SqlUtilityCountSql -CountSourceSql -OutputColumnCount`. |
| `modules/SqlUtility.Database.ps1` | Add `Invoke-SqlUtilityExactCount -Server -Database -CountSql -CommandTimeoutSeconds [-Executor]`, returning a non-negative `Int64`. |
| `SqlUtility.ps1` | Add `BuildCountSql` and `ExecuteCount` services, count state/action/status helpers, one-line toolbar, native styling, Segoe UI/Consolas fonts, and 300-pixel grid cap. |
| `tests/Test-QueryPolicy.ps1` | Cover top-level `ORDER BY` removal and schema-count-based alias-safe wrapper generation. |
| `tests/Test-Database.ps1` | Cover scalar executor inputs, timeout forwarding, `Int64` conversion, invalid scalar rejection, exception propagation, and disposal boundary. |
| `tests/Test-SqlUtilityUi.ps1` | Cover status variants, count enablement/lifecycle/failure preservation, toolbar order/text/accessibility, fonts/style, and 300-pixel cap. |
| `README.md` | Document the new toolbar, explicit count semantics/cost, page-status meanings, fonts, 300-pixel cap, and unchanged splitter/paging/export boundaries. |

### New and extended interfaces

```powershell
# Query policy result, valid query
[pscustomobject][ordered]@{
    IsValid = $true
    ErrorMessage = ''
    NormalizedSql = $normalizedSql
    TableIdentifier = $tableIdentifier
    HasOrderBy = [bool] $hasOrderBy
    CountSourceSql = $countSourceSql
}

function New-SqlUtilityCountSql {
    param(
        [Parameter(Mandatory = $true)][string] $CountSourceSql,
        [Parameter(Mandatory = $true)][int] $OutputColumnCount
    )
    # Returns one SELECT COUNT_BIG(*) wrapper with a complete derived-table
    # column alias list matching the successful result schema.
}

function Invoke-SqlUtilityExactCount {
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][string] $CountSql,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor = ${function:Invoke-SqlUtilityScalarExecutor}
    )
    # Returns [long].
}
```

The UI service map gains these scriptblocks:

```powershell
BuildCountSql = {
    param($CountSourceSql, $OutputColumnCount)
    New-SqlUtilityCountSql -CountSourceSql $CountSourceSql -OutputColumnCount $OutputColumnCount
}
ExecuteCount = {
    param($Server, $Database, $CountSql, $TimeoutSeconds)
    Invoke-SqlUtilityExactCount -Server $Server -Database $Database -CountSql $CountSql `
        -CommandTimeoutSeconds $TimeoutSeconds
}
```

The last-successful-query snapshot gains `CountSql`; transient UI state gains nullable `ExplicitTotalRowCount`.

---

### Task 1: Derive and build an alias-safe exact-count query

**Files:**

- Modify: `modules/SqlUtility.QueryPolicy.ps1`
- Test: `tests/Test-QueryPolicy.ps1`

- [ ] **Step 1: Add failing validation-result shape assertions**

Extend the existing valid and invalid query assertions so every result contains `CountSourceSql`. Use representative cases:

```powershell
$unordered = Test-SqlUtilityQuery -Sql 'SELECT * FROM dbo.Items;'
Assert-Equal 'SELECT * FROM dbo.Items' $unordered.CountSourceSql `
    'Unordered count source keeps the complete normalized query'

$ordered = Test-SqlUtilityQuery -Sql @'
SELECT CategoryId, COUNT(*)
FROM dbo.Items
GROUP BY CategoryId
HAVING COUNT(*) > 5
ORDER BY CategoryId;
'@
Assert-Equal @'
SELECT CategoryId, COUNT(*)
FROM dbo.Items
GROUP BY CategoryId
HAVING COUNT(*) > 5
'@ $ordered.CountSourceSql 'Ordered count source removes only top-level ORDER BY'

$invalid = Test-SqlUtilityQuery -Sql 'DELETE FROM dbo.Items'
Assert-Equal '' $invalid.CountSourceSql 'Invalid result has an empty count source'
```

Also add cases proving that `ORDER BY` text inside string literals/comments and ordering inside accepted expressions are not used as the truncation point. Retain existing grammar acceptance and rejection cases unchanged.

- [ ] **Step 2: Run the focused query-policy suite and confirm the expected failure**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
```

Expected: failure because `CountSourceSql` does not yet exist.

- [ ] **Step 3: Extend validation results and derive the order-free source from token positions**

Add `CountSourceSql = ''` to `New-SqlUtilityInvalidQueryResult`.

In `Test-SqlUtilityQuery`, retain the actual top-level `ORDER` token used to set `HasOrderBy`. After normalization, calculate the source with the token's `Start` offset instead of searching raw text:

```powershell
$countSourceSql = $normalizedSql
if ($null -ne $topLevelOrderToken) {
    $countSourceSql = $normalizedSql.Substring(0, $topLevelOrderToken.Start).TrimEnd()
}
```

If normalization removes a trailing semicolon before trailing comments, verify the stored token positions still refer to the normalized string. If they do not, recompute tokens from `$normalizedSql` and locate the already-validated top-level `ORDER` token there. Do not use regex or `LastIndexOf('ORDER BY')`.

Return `CountSourceSql` after `HasOrderBy` in both valid and invalid result objects.

- [ ] **Step 4: Add failing builder tests for wildcard, unnamed, and duplicate projections**

Test the builder independently of SQL Server:

```powershell
$countSql = New-SqlUtilityCountSql `
    -CountSourceSql 'SELECT *, Price * 1.25 FROM dbo.Items' `
    -OutputColumnCount 4

$expected = @'
SELECT COUNT_BIG(*)
FROM (
SELECT *, Price * 1.25 FROM dbo.Items
) AS [SqlUtilityCountSource] ([SqlUtilityCountColumn1], [SqlUtilityCountColumn2], [SqlUtilityCountColumn3], [SqlUtilityCountColumn4]);
'@
Assert-Equal $expected $countSql 'Count wrapper supplies one alias per actual result column'

Assert-Throws {
    New-SqlUtilityCountSql -CountSourceSql 'SELECT Id FROM dbo.Items' -OutputColumnCount 0
} 'System.ArgumentOutOfRangeException' 'Count builder rejects a zero-column schema'

Assert-Throws {
    New-SqlUtilityCountSql -CountSourceSql '   ' -OutputColumnCount 1
} 'System.ArgumentException' 'Count builder rejects an empty source'
```

Expected first run: failure because `New-SqlUtilityCountSql` does not exist.

- [ ] **Step 5: Implement the smallest schema-aware wrapper builder**

Use fixed application-owned identifiers only; do not derive aliases from user SQL or returned display names:

```powershell
$columnAliases = @(
    for ($index = 1; $index -le $OutputColumnCount; $index++) {
        '[SqlUtilityCountColumn{0}]' -f $index
    }
)

return "SELECT COUNT_BIG(*)`r`nFROM (`r`n{0}`r`n) AS [SqlUtilityCountSource] ({1});" -f `
    $CountSourceSql.Trim(), ($columnAliases -join ', ')
```

This alias list deliberately replaces derived output names. It allows wildcard expansion, unnamed expressions, and duplicate names without parsing or rewriting the select list. `DISTINCT`, `WHERE`, `GROUP BY`, and `HAVING` remain inside the derived source; only the validated top-level `ORDER BY` is absent.

- [ ] **Step 6: Run the focused suite and inspect the diff**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff -- modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
```

Expected: query-policy tests pass; no whitespace errors; no accepted/rejected grammar case changes.

- [ ] **Step 7: Commit the policy slice**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' add modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --cached --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' commit -m "feat: build exact row count queries"
```

---

### Task 2: Execute the exact count through the database boundary

**Files:**

- Modify: `modules/SqlUtility.Database.ps1`
- Test: `tests/Test-Database.ps1`

- [ ] **Step 1: Add failing tests for the public exact-count function**

Use the existing injected-executor style:

```powershell
$script:countCall = $null
$count = Invoke-SqlUtilityExactCount -Server ' server ' -Database ' db ' `
    -CountSql 'SELECT COUNT_BIG(*) FROM (...) AS q ([c1]);' `
    -CommandTimeoutSeconds 321 -Executor {
        param($ConnectionString, $CommandText, $CommandTimeoutSeconds)
        $script:countCall = [pscustomobject]@{
            ConnectionString = $ConnectionString
            CommandText = $CommandText
            CommandTimeoutSeconds = $CommandTimeoutSeconds
        }
        return [decimal] 922337203685477580
    }

Assert-Equal ([long] 922337203685477580) $count 'Exact count returns Int64'
$builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($script:countCall.ConnectionString)
Assert-Equal 'server' $builder.DataSource 'Exact count trims and forwards server'
Assert-Equal 'db' $builder.InitialCatalog 'Exact count trims and forwards database'
Assert-True $builder.IntegratedSecurity 'Exact count uses integrated authentication'
Assert-Equal 'SELECT COUNT_BIG(*) FROM (...) AS q ([c1]);' $script:countCall.CommandText `
    'Exact count forwards only policy-generated SQL'
Assert-Equal 321 $script:countCall.CommandTimeoutSeconds 'Exact count forwards query timeout'
```

Add rejection/propagation cases for `$null`, `[DBNull]::Value`, a nonnumeric scalar, a negative scalar, and an executor-thrown exception. Assert no fallback or second call occurs.

- [ ] **Step 2: Run the focused database suite and confirm the missing-function failure**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
```

Expected: failure because `Invoke-SqlUtilityExactCount` does not yet exist.

- [ ] **Step 3: Add the default scalar executor with deterministic disposal**

Implement a private database helper adjacent to `Invoke-SqlUtilityTableExecutor`:

```powershell
function Invoke-SqlUtilityScalarExecutor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ConnectionString,
        [Parameter(Mandatory = $true)][string] $CommandText,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds
    )

    $connection = [System.Data.SqlClient.SqlConnection]::new($ConnectionString)
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        try {
            $command.CommandText = $CommandText
            $command.CommandTimeout = $CommandTimeoutSeconds
            return $command.ExecuteScalar()
        }
        finally {
            if ($null -ne $command) { $command.Dispose() }
        }
    }
    finally {
        $connection.Dispose()
    }
}
```

Keep it parameter-free because all user semantics are already contained in the validated, policy-generated count SQL. Do not add UI or WinForms types to this module.

- [ ] **Step 4: Implement and validate the public count boundary**

Build the integrated-security connection string with `New-SqlUtilityConnectionString`, invoke the executor exactly once, reject null/DB null, convert through `[System.Convert]::ToInt64`, and reject negative values:

```powershell
$value = & $Executor $connectionString $CountSql $CommandTimeoutSeconds
if ($null -eq $value -or $value -eq [DBNull]::Value) {
    throw [System.Data.DataException]::new('The row-count query did not return a value.')
}

try {
    $count = [System.Convert]::ToInt64($value, [System.Globalization.CultureInfo]::InvariantCulture)
}
catch {
    throw [System.Data.DataException]::new('The row-count query returned an invalid value.', $_.Exception)
}

if ($count -lt 0) {
    throw [System.Data.DataException]::new('The row-count query returned a negative value.')
}
return [long] $count
```

Validate `CommandTimeoutSeconds -ge 1` and nonblank `CountSql` before opening a connection.

- [ ] **Step 5: Rerun the focused suite and review module boundaries**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff -- modules/SqlUtility.Database.ps1 tests/Test-Database.ps1
```

Expected: database tests pass; the new path uses one scalar execution, honors timeout, and owns all ADO.NET cleanup.

- [ ] **Step 6: Commit the database slice**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' add modules/SqlUtility.Database.ps1 tests/Test-Database.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --cached --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' commit -m "feat: execute exact row counts"
```

---

### Task 3: Add page-status semantics and explicit count state/action

**Files:**

- Modify: `SqlUtility.ps1`
- Test: `tests/Test-SqlUtilityUi.ps1`

- [ ] **Step 1: Extend the UI test service map with count services**

Update the shared fake services so `Assert-SqlUtilityServices` can require the new boundaries without breaking unrelated tests:

```powershell
BuildCountSql = {
    param($CountSourceSql, $OutputColumnCount)
    return "COUNT SOURCE=$CountSourceSql COLUMNS=$OutputColumnCount"
}
ExecuteCount = {
    param($Server, $Database, $CountSql, $TimeoutSeconds)
    return [long] 1234
}
```

Update fake validation results to include `CountSourceSql`.

- [ ] **Step 2: Add failing pure status-text tests**

Introduce tests for a helper named `Get-SqlUtilityPageStatusText -State $state -PageResult $result`:

```powershell
# Complete unordered cache
Assert-Equal 'Page 1 - 500 of 723' (Get-SqlUtilityPageStatusText -State $completeState -PageResult $completePage) `
    'Complete unordered status shows exact cache total'

# Truncated unordered cache, configured limit 1000
Assert-Equal 'Page 1 - 500 of 1000+' (Get-SqlUtilityPageStatusText -State $truncatedState -PageResult $truncatedPage) `
    'Truncated unordered status marks lower bound'

# Ordered, unknown exact total
Assert-Equal 'Page 1 - 500' (Get-SqlUtilityPageStatusText -State $orderedState -PageResult $orderedPage) `
    'Ordered status omits unknown total'

# Explicit count overrides the unknown/lower-bound suffix
$orderedState.ExplicitTotalRowCount = [long] 1234
Assert-Equal 'Page 1 - 500 of 1234' (Get-SqlUtilityPageStatusText -State $orderedState -PageResult $orderedPage) `
    'Explicit count appears in ordered status'

# Empty first page only
Assert-Equal 'Page 1 - 0 of 0' (Get-SqlUtilityPageStatusText -State $orderedStateWithoutCount -PageResult $emptyFirstPage) `
    'Empty ordered first page proves zero rows'
Assert-Equal 'Page 2 - 0' (Get-SqlUtilityPageStatusText -State $orderedStateWithoutCount -PageResult $emptyLaterPage) `
    'Empty later ordered page does not claim the total is zero'
```

Construct caches with the existing `New-TestDataTable` helper. Do not add production-only hooks.

- [ ] **Step 3: Run the UI suite and confirm the helper is missing**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: failure at the first new status assertion.

- [ ] **Step 4: Implement status precedence and bind it in `Show-SqlUtilityPage`**

Use this precedence:

1. `ExplicitTotalRowCount` when non-null.
2. Complete unordered cache row count.
3. Truncated unordered configured-limit lower bound with `+`.
4. Ordered page 1 with zero displayed rows: exact zero.
5. Otherwise no total suffix.

Format page numbers and displayed rows as plain invariant integers. Format exact totals with invariant `N0` grouping (for example `123,456`) and keep the configured lower-bound text ungrouped (for example `1000+`). Replace:

```powershell
$pageStatus.Text = 'Page {0} - {1} rows' -f ...
```

with the helper result.

- [ ] **Step 5: Add failing execution-snapshot and lifecycle tests**

For an ordered execution with a two-column result, record the `BuildCountSql` call and assert:

```powershell
Assert-Equal 'SELECT Id, Name FROM dbo.Items' $script:buildCountCall.CountSourceSql `
    'Execute forwards the validated order-free count source'
Assert-Equal 2 $script:buildCountCall.OutputColumnCount `
    'Execute finalizes count SQL from the actual result schema'
Assert-Equal 'generated count sql' $form.Tag.ExecutedQuery.CountSql `
    'Successful snapshot stores policy-generated count SQL'
Assert-Equal $null $form.Tag.ExplicitTotalRowCount `
    'A new successful execution starts without an explicit count'
```

Add assertions that a new execution, invalid/failed execution, and confirmed connection change clear `ExplicitTotalRowCount`. An editor change keeps the prior result and its displayed point-in-time total visible, matching the existing stale-result behavior, but marks the snapshot stale and disables Count, paging, and export. If the editor text is restored exactly and the existing workflow makes the snapshot fresh again, the stored total remains tied to that immutable snapshot.

When building the successful snapshot, call `BuildCountSql` only after page execution succeeds and `PageResult.Data.Columns.Count` is known. A builder failure is an execution failure and must not publish a partial snapshot.

- [ ] **Step 6: Add failing Count-button enablement and action tests**

Test these states through the named `CountButton` control:

- disabled before any result;
- enabled for a fresh ordered result;
- enabled for a fresh truncated unordered result;
- disabled for a fresh complete unordered result because the cache already gives the exact total;
- disabled while busy or stale;
- remains enabled after a successful explicit count for refresh;
- success updates `ExplicitTotalRowCount` and status without rebinding or changing the grid/result object;
- failure shows an error and preserves the result object, pager/export enablement, grid data, status, and prior exact count.

Record the execution call:

```powershell
Assert-Equal 'server' $script:countCall.Server 'Count uses active server snapshot'
Assert-Equal 'database' $script:countCall.Database 'Count uses active database snapshot'
Assert-Equal $form.Tag.ExecutedQuery.CountSql $script:countCall.CountSql `
    'Count executes only the stored policy-generated SQL'
Assert-Equal $form.Tag.Config.queryExportTimeoutSeconds $script:countCall.TimeoutSeconds `
    'Count uses the Query/Export timeout'
```

- [ ] **Step 7: Implement count enablement and `Invoke-SqlUtilityCountAction`**

Add `ExplicitTotalRowCount = $null` to form state. Add `BuildCountSql` and `ExecuteCount` to default services and required-service validation.

Centralize button refresh in a small helper (for example `Update-SqlUtilityQueryActionState`) called after page binding, stale changes, busy transitions, execution failure, count completion/failure, and connection changes. Count enablement is:

```powershell
$hasFreshResult = -not $state.IsBusy -and -not $state.IsQueryStale -and `
    $null -ne $state.ExecutedQuery -and $null -ne $state.CurrentResult
$exactCacheKnown = $hasFreshResult -and -not [bool] $state.ExecutedQuery.HasOrderBy -and `
    [bool] $state.CurrentResult.IsComplete -and -not [bool] $state.CurrentResult.IsTruncated
$countButton.Enabled = $hasFreshResult -and -not $exactCacheKnown
```

The action must:

1. Return if disabled by state guards.
2. Copy the prior `ExplicitTotalRowCount` and leave all result references untouched.
3. Set busy state and disable query actions.
4. Call `ExecuteCount` with the stored snapshot SQL and existing timeout.
5. On success only, assign the returned `Int64` and recompute page status.
6. On failure, restore/retain the prior count and show `Row Count Failed` with an error icon.
7. In `finally`, clear busy state and recompute enablement.

Do not call `Show-SqlUtilityPage` on count success if that would rebind the grid. Update only the page-status label and action enabled states.

- [ ] **Step 8: Run the focused UI suite and inspect state transitions**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
```

Expected: all prior UI workflow tests plus new count/status tests pass; no count execution occurs during Execute, paging, or export.

- [ ] **Step 9: Commit the UI behavior slice**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --cached --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' commit -m "feat: add explicit result row counts"
```

---

### Task 4: Modernize the native query workspace and consolidate the toolbar

**Files:**

- Modify: `SqlUtility.ps1`
- Test: `tests/Test-SqlUtilityUi.ps1`

- [ ] **Step 1: Add failing appearance and toolbar-structure assertions**

Assert these control contracts:

```powershell
Assert-Equal 'Segoe UI' $form.Font.Name 'Application uses Segoe UI'
Assert-Equal 9 ([int] $form.Font.SizeInPoints) 'Application uses Segoe UI 9pt'
Assert-Equal 'Consolas' $queryEditor.Font.Name 'SQL editor keeps Consolas'
Assert-Equal 10 ([int] $queryEditor.Font.SizeInPoints) 'SQL editor keeps Consolas 10pt'

Assert-Equal 'Execute' $executeButton.Text 'Execute label remains explicit'
Assert-Equal 'ORDER BY required for paging.' $pagingHelp.Text 'Paging help stays beside Execute'
Assert-Equal 'Count' $countButton.Text 'Count action is explicit'
Assert-Equal '<' $previousButton.Text 'Previous action uses compact label'
Assert-Equal '>' $nextButton.Text 'Next action uses compact label'
Assert-Equal 'Export' $exportButton.Text 'Export action uses compact label'
Assert-Equal 'Previous page' $previousButton.AccessibleName 'Previous action retains accessible meaning'
Assert-Equal 'Next page' $nextButton.AccessibleName 'Next action retains accessible meaning'
Assert-Equal 'Export to Excel' $exportButton.AccessibleName 'Export retains accessible meaning'
```

Name the one-row container `QueryActionLayout`. Assert it is a `TableLayoutPanel`, has one row, and its child controls appear in this exact left-to-right order:

```text
Execute | ORDER BY required for paging. | flexible space | page status | Count | < | > | Export
```

The help label occupies the percent-sized flexible column; all other columns are autosized. Assert the action area is shorter than the existing 64-pixel two-row panel and that every control fits at the form's existing minimum size.

Assert restrained native styling properties rather than pixel colors alone:

- Execute uses the restrained blue application accent and readable foreground.
- Other buttons keep native/light styling.
- Grid headers use quiet light background, dark foreground, and no forced high-contrast-incompatible custom painting.
- `querySplit.FixedPanel` remains `None` and `IsSplitterFixed` remains false.

- [ ] **Step 2: Add the failing 300-pixel cap assertion**

Update the existing grid-width check from 200 to 300 and ensure the test data contains one long value that would autosize beyond 300:

```powershell
foreach ($column in $resultsGrid.Columns) {
    Assert-True ($column.Width -le 300) 'Result columns are capped at 300 pixels'
}
Assert-Equal 300 $longValueColumn.Width 'A long result column reaches the new cap'
```

- [ ] **Step 3: Run the UI suite and confirm the intended layout/style failures**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: failures for old fonts/button labels/two-row panel/200-pixel cap and missing Count layout control.

- [ ] **Step 4: Enable native visual styles before form construction**

In `Start-SqlUtilityApplication`, before creating the form, call:

```powershell
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
```

Do not call either after controls exist in production. Preserve `-NoGui` test-loading behavior and the current STA launcher.

- [ ] **Step 5: Apply scoped fonts and dispose owned font objects**

Create an application font and editor font in the form factory:

```powershell
$interfaceFont = [System.Drawing.Font]::new('Segoe UI', 9.0)
$editorFont = [System.Drawing.Font]::new('Consolas', 10.0)
$form.Font = $interfaceFont
$queryEditor.Font = $editorFont
```

Include any explicit bold derivative in the same owned-font collection. Dispose only fonts created by this form after the form closes/disposes; do not dispose `SystemFonts` or fonts owned by controls/framework. Keep the existing editor multiline, scroll, accept-tab, and no-word-wrap behavior.

- [ ] **Step 6: Replace the absolute two-row toolbar with one `TableLayoutPanel`**

Use seven control columns:

1. Execute: `AutoSize`
2. Help label: `Percent, 100`
3. Page status: `AutoSize`
4. Count: `AutoSize`
5. Previous: `AutoSize`
6. Next: `AutoSize`
7. Export: `AutoSize`
Because the percent-sized help column is wider than its left-anchored text, the unused remainder is the approved flexible space between the paging notice and page status. Do not add a visible spacer control. Set `AutoEllipsis = $true` on the help label so it yields space first at minimum width. Align Execute and help left; status and navigation/export right.

Set compact labels and accessible names/descriptions. Wire the Count click event to `Invoke-SqlUtilityCountAction`. Keep existing Execute/Page/Export event handlers and keyboard behavior.

- [ ] **Step 7: Apply restrained native colors and the new column cap**

Use SystemColors for light surfaces, text, grid, and quiet headers. Reserve a muted Windows-like blue for Execute only, with a readable foreground and restrained flat border. Avoid dark mode, owner drawing, custom themes, images, and external assets.

Change only the cap expression:

```powershell
$width = [Math]::Min(300, $column.Width)
```

Keep the existing `AutoResizeColumn(...AllCells)` measurement, fill-weight reset, and user resizing behavior.

- [ ] **Step 8: Rerun focused UI tests and perform a local visual smoke check**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Then launch from the worktree only for a local visual check if the environment permits:

```powershell
.\StartSqlUtility.cmd
```

Check: one toolbar row at minimum and normal sizes; help ellipsis before controls overlap; Segoe UI chrome and Consolas editor; draggable horizontal splitter; Count status alignment; grid headers; high-DPI layout; 300-pixel cap. Do not claim Citrix acceptance from this local check.

- [ ] **Step 9: Inspect and commit the appearance slice**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --cached --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' commit -m "feat: modernize the native query workspace"
```

---

### Task 5: Synchronize user documentation

**Files:**

- Modify: `README.md`
- Modify only if needed for technical precision: `docs/superpowers/specs/2026-08-11-gui-modernization-row-count-design.md`

- [ ] **Step 1: Update the README workflow and architecture table**

Document:

- one-line toolbar labels and order;
- `<` and `>` meanings;
- `Count` is explicit and never automatic;
- complete unordered results already show an exact cache total and disable Count;
- truncated unordered status uses `limit+` until counted;
- ordered results omit the total until Count is selected;
- count uses a separate query, may be expensive on large/complex result sets, uses the configured Query/Export timeout, and is a point-in-time value that can drift if data changes;
- count preserves `DISTINCT`, `WHERE`, `GROUP BY`, and `HAVING`, and ignores only the top-level display `ORDER BY`;
- native visual styles, Segoe UI 9pt chrome, Consolas 10pt editor, 300-pixel automatic column cap;
- the result splitter remains user-draggable and no dark theme/external resource is added.

Update the architecture table so QueryPolicy owns count-source/wrapper generation, Database owns scalar count execution, and UI owns count state/rendering. Replace references to the 200-pixel cap and full old button labels.

- [ ] **Step 2: Reconcile the approved design's technical interface wording if necessary**

If implementation uses the schema-aware two-stage contract described by this plan, adjust only the design's internal interface paragraph to say:

```text
Validation exposes an order-free CountSourceSql. After a successful page reveals
the actual output-column count, QueryPolicy finalizes the alias-safe CountSql,
which the UI stores in the last-successful-query snapshot.
```

Do not change approved behavior or scope. This clarification explains how wildcard, unnamed, and duplicate result columns are supported.

- [ ] **Step 3: Review documentation diff and commit**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff -- README.md docs/superpowers/specs/2026-08-11-gui-modernization-row-count-design.md
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' add README.md docs/superpowers/specs/2026-08-11-gui-modernization-row-count-design.md
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --cached --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' commit -m "docs: describe modern query workspace and counting"
```

If the design spec required no edit, stage only `README.md`.

---

### Task 6: Run full verification and prepare handoff

**Files:**

- Verify: all production, test, and documentation changes in the worktree

- [ ] **Step 1: Run all focused suites owning changed contracts**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: all pass.

- [ ] **Step 2: Run the aggregate suite from the worktree root**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
```

Expected: configuration, query policy, database, Excel, UI, launcher, and aggregate tests all pass.

- [ ] **Step 3: Perform static and repository checks**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff --check
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' status --short
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' log --oneline --decorate -6
```

Expected: no whitespace errors; no unexpected/untracked files; only intentional commits on `codex/gui-modernization-row-count`.

- [ ] **Step 4: Review the complete branch diff against the approved design**

```powershell
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff main...HEAD --stat
git -c safe.directory='C:/Users/bluesWalker/Documents/SQL Utility/.worktrees/gui-modernization-row-count' diff main...HEAD -- SqlUtility.ps1 modules/SqlUtility.QueryPolicy.ps1 modules/SqlUtility.Database.ps1 README.md tests/Test-QueryPolicy.ps1 tests/Test-Database.ps1 tests/Test-SqlUtilityUi.ps1 docs/superpowers/specs/2026-08-11-gui-modernization-row-count-design.md
```

Confirm explicitly:

- no implicit count call;
- no query grammar widening;
- no config/schema/persistence change;
- no launcher/distribution dependency change;
- no Database-to-WinForms or QueryPolicy-to-Database boundary leak;
- complete unordered export and ordered export behavior unchanged;
- splitter remains free;
- every count/status/lifecycle state from the approved design is covered.

- [ ] **Step 5: Scan the plan implementation for placeholders and interface drift**

```powershell
rg -n "TODO|TBD|FIXME|PLACEHOLDER" SqlUtility.ps1 modules tests README.md docs/superpowers/specs/2026-08-11-gui-modernization-row-count-design.md
rg -n "CountSourceSql|New-SqlUtilityCountSql|Invoke-SqlUtilityExactCount|BuildCountSql|ExecuteCount|ExplicitTotalRowCount|CountButton" SqlUtility.ps1 modules tests
```

Expected: no newly introduced placeholder; every new service/property/function spelling is consistent across production and tests.

- [ ] **Step 6: Record external acceptance still pending**

Handoff must label these as pending unless actually performed:

- Citrix launch and DPI/font rendering;
- live SQL Server count correctness/performance under representative permissions and data sizes;
- concurrent-data drift behavior observation;
- cloud-drive configuration/export authorization;
- opening exported workbooks in desktop Excel.

- [ ] **Step 7: Ask the user how to integrate the verified feature branch**

Do not merge, push, delete the branch, or remove the worktree without the user's choice.

---

## Acceptance Matrix

| Scenario | Expected result |
| --- | --- |
| Complete unordered, 723 rows | `Page 1 - 500 of 723`; Count disabled. |
| Complete unordered, 0 rows | `Page 1 - 0 of 0`; Count disabled. |
| Truncated unordered, limit 1000 | `Page 1 - 500 of 1000+`; Count enabled. |
| Ordered first page, 500 rows | `Page 1 - 500`; Count enabled. |
| Ordered first page, 0 rows | `Page 1 - 0 of 0`; Count remains enabled because only a complete unordered cache suppresses the action. |
| Ordered later empty page | `Page N - 0`; never infer total zero. |
| Successful explicit count 12,345 | Status becomes `Page N - X of 12345`; grid and page unchanged; Count remains enabled. |
| Count fails after a prior count | Prior count/status/result/export state remain; error message shown. |
| Editor changes | Results and their prior total remain visible but stale; Count/Export/Paging disable per existing stale rules. |
| New execute/failure/connection change | Prior exact count clears. |
| Projection has `*`, unnamed expressions, or duplicate names | Count wrapper uses actual schema column count and fixed aliases; count succeeds without projection rewriting. |
| Query has top-level `ORDER BY` | Only that clause is removed from count source. |
| Query has `DISTINCT`/`WHERE`/`GROUP BY`/`HAVING` | All remain inside count source. |
| Window resized | Existing splitter remains draggable; one-row action bar does not overlap; help yields space first. |
| Long grid value | Automatic width stops at 300 pixels; user can still resize normally. |
