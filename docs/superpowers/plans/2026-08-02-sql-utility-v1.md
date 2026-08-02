# SQL Utility Version 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the proven connection POC with a portable Windows PowerShell 5.1 WinForms utility that stores non-secret connection settings, enforces a single-table read-only query policy, pages query results, and exports complete results to dependency-free `.xlsx` files.

**Architecture:** `SqlUtility.ps1` owns WinForms state and rendering while four dot-sourced files isolate configuration, query validation, SQL execution/result parsing, and workbook generation. SQL resources remain inside the database module; UI and Excel code exchange neutral `DataTable`, metadata, schema, and row values. Every operation is synchronous, bounded by the approved settings, and tested through injected scriptblock boundaries without requiring a live SQL Server.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework, `System.Windows.Forms`, `System.Drawing`, `System.Data.SqlClient`, `System.IO.Compression`, `System.Xml`, JSON via built-in PowerShell cmdlets, dependency-free PowerShell test scripts.

## Global Constraints

- Target Windows PowerShell 5.1 and .NET Framework components normally present in the Citrix environment.
- Launch only through `StartSqlUtility.cmd`, which invokes `powershell.exe` with `-NoLogo -NoProfile -STA -ExecutionPolicy Bypass`.
- Require no installer, compiled executable, administrator rights, third-party module, NuGet package, Office installation, `sqlcmd`, Python runtime, registry write, or environment-variable write.
- Use Windows integrated authentication only; never accept or persist credentials.
- Write persistent app state only to `SqlUtility.config.json` beside `SqlUtility.ps1`; write Excel output only to the destination selected by the user.
- Keep Server and Database blank on every launch even when saved connections exist.
- Allow only the approved version 1 single-table read-only `SELECT` policy.
- Keep display pages fixed at 500 rows.
- Keep unordered-row limit configurable from 100 through 2000, default 1000.
- Keep Query/Export timeout configurable from 5 through 3600 seconds, default 120; keep connection timeout fixed at 10 seconds.
- Reject exports beyond 1,048,575 data rows plus one header row and leave no incomplete destination.
- Use TDD for every implementation task and commit only after the task-specific test plus the full available suite pass.

## Planned File Structure

```text
.gitignore                              Runtime config/temp exclusions
StartSqlUtility.cmd                     Production launch entry point
SqlUtility.ps1                          WinForms, app state, workflows, rendering
modules/
  SqlUtility.Config.ps1                 JSON model, validation, atomic persistence
  SqlUtility.QueryPolicy.ps1            Tokenizer and v1 SELECT policy
  SqlUtility.Database.ps1               SQL ownership, result parsing, paging/streaming
  SqlUtility.Excel.ps1                  Neutral row stream to safe XLSX package
tests/
  Test-Helpers.ps1                      Dependency-free assertions/test data
  Test-Config.ps1                       Configuration unit tests
  Test-QueryPolicy.ps1                  Query grammar/policy unit tests
  Test-Database.ps1                     Connection, paging, disposal boundary tests
  Test-Excel.ps1                        XLSX package and safe-write tests
  Test-SqlUtilityUi.ps1                 WinForms state/workflow tests with fakes
  Test-Launcher.ps1                     Distribution and launch contract tests
  Test-All.ps1                          Ordered aggregate test runner
```

The production files `SqlConnectionPoc.ps1`, `StartSqlPoc.cmd`, and `tests/Test-SqlConnectionPoc.ps1` are removed only in the final task, after the replacement suite covers the proven connection behavior.

---

### Task 1: Configuration Model and Safe Persistence

**Files:**

- Create: `modules/SqlUtility.Config.ps1`
- Create: `tests/Test-Helpers.ps1`
- Create: `tests/Test-Config.ps1`
- Modify: `.gitignore`

**Interfaces:**

- Produces: `New-SqlUtilityDefaultConfig -> PSCustomObject`
- Produces: `ConvertTo-SqlUtilityValidatedConfig -InputObject <object> -> PSCustomObject` or throws `ArgumentException`/`NotSupportedException`
- Produces: `Read-SqlUtilityConfig -Path <string> -> PSCustomObject`; a missing file returns unsaved defaults
- Produces: `Write-SqlUtilityConfig -Path <string> -Config <object> -> validated PSCustomObject`; complete safe replacement or exception
- Produces: `Add-SqlUtilitySavedConnection -Config <object> -Server <string> -Database <string> -> new PSCustomObject`
- Produces: `Remove-SqlUtilitySavedConnection -Config <object> -Server <string> -Database <string> -> new PSCustomObject`
- Contract: returned configs always contain `schemaVersion`, `unorderedRowLimit`, `queryExportTimeoutSeconds`, and an array-valued `connections` property.

- [ ] **Step 1: Add the shared dependency-free assertion helpers**

Create `tests/Test-Helpers.ps1` with assertions that accumulate failures without terminating the first test:

```powershell
$script:Failures = 0

function Assert-True([bool] $Condition, [string] $Message) {
    if (-not $Condition) {
        $script:Failures++
        Write-Error -ErrorAction Continue "FAIL: $Message"
    }
}

function Assert-Equal($Expected, $Actual, [string] $Message) {
    Assert-True ($Expected -eq $Actual) "$Message (expected '$Expected', got '$Actual')"
}

function Assert-Throws([scriptblock] $Action, [string] $ExceptionType, [string] $Message) {
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception }
    Assert-True ($null -ne $caught) $Message
    if ($null -ne $caught -and $ExceptionType) {
        Assert-Equal $ExceptionType $caught.GetType().FullName "$Message exception type"
    }
}

function Complete-TestFile([string] $SuccessMessage) {
    if ($script:Failures -gt 0) { exit 1 }
    Write-Host $SuccessMessage
}
```

- [ ] **Step 2: Write failing configuration tests**

Create `tests/Test-Config.ps1`, dot-source the helpers and future module, and cover these exact cases in a test-owned GUID directory under `[System.IO.Path]::GetTempPath()`:

```powershell
$defaults = New-SqlUtilityDefaultConfig
Assert-Equal 1 $defaults.schemaVersion 'Default schema version'
Assert-Equal 1000 $defaults.unorderedRowLimit 'Default unordered limit'
Assert-Equal 120 $defaults.queryExportTimeoutSeconds 'Default timeout'
Assert-Equal 0 @($defaults.connections).Count 'Default saved connections'

$missing = Read-SqlUtilityConfig -Path (Join-Path $testRoot 'missing.json')
Assert-Equal 1000 $missing.unorderedRowLimit 'Missing config uses defaults'
Assert-True (-not (Test-Path (Join-Path $testRoot 'missing.json'))) 'Read does not create config'

$withPair = Add-SqlUtilitySavedConnection -Config $defaults -Server ' ServerA ' -Database ' DbA '
$deduplicated = Add-SqlUtilitySavedConnection -Config $withPair -Server 'servera' -Database 'dba'
Assert-Equal 1 @($deduplicated.connections).Count 'Connection pair is unique case-insensitively'
Assert-Equal 'servera' $deduplicated.connections[0].server 'Latest server casing is kept'
Assert-Equal 'dba' $deduplicated.connections[0].database 'Latest database casing is kept'

$roundTripPath = Join-Path $testRoot 'SqlUtility.config.json'
$saved = Write-SqlUtilityConfig -Path $roundTripPath -Config $deduplicated
$loaded = Read-SqlUtilityConfig -Path $roundTripPath
Assert-Equal ($saved | ConvertTo-Json -Depth 4) ($loaded | ConvertTo-Json -Depth 4) 'Config round-trips'

foreach ($invalidLimit in @(99, 2001, 100.5, '100')) {
    $candidate = New-SqlUtilityDefaultConfig
    $candidate.unorderedRowLimit = $invalidLimit
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig $candidate } 'System.ArgumentException' "Reject limit $invalidLimit"
}
foreach ($invalidTimeout in @(4, 3601, 5.5, '120')) {
    $candidate = New-SqlUtilityDefaultConfig
    $candidate.queryExportTimeoutSeconds = $invalidTimeout
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig $candidate } 'System.ArgumentException' "Reject timeout $invalidTimeout"
}
```

Also assert inclusive boundaries `100`, `2000`, `5`, and `3600`; blank connection fields; malformed JSON; missing required properties; schema version `2`; confirmed removal behavior; and that a simulated replacement failure preserves existing bytes and removes `*.tmp` files. Simulate the failure by passing a test-only `-ReplaceAction` scriptblock to `Write-SqlUtilityConfig` that throws before replacement.

- [ ] **Step 3: Run the configuration test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
```

Expected: FAIL because `modules/SqlUtility.Config.ps1` or its functions do not exist.

- [ ] **Step 4: Implement the validated immutable-style configuration operations**

Create `modules/SqlUtility.Config.ps1`. Keep the public functions above and use this exact safe-write structure:

```powershell
function Write-SqlUtilityConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)] $Config,
        [scriptblock] $ReplaceAction
    )

    $validated = ConvertTo-SqlUtilityValidatedConfig -InputObject $Config
    $directory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($Path))
    if (-not [System.IO.Directory]::Exists($directory)) {
        throw [System.IO.DirectoryNotFoundException]::new("Configuration directory does not exist: $directory")
    }
    $temporaryPath = Join-Path $directory ('.SqlUtility.config.{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    try {
        $json = $validated | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($temporaryPath, $json, [System.Text.UTF8Encoding]::new($false))
        if ($ReplaceAction) {
            & $ReplaceAction $temporaryPath $Path
        }
        elseif ([System.IO.File]::Exists($Path)) {
            [System.IO.File]::Replace($temporaryPath, $Path, $null)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
        return $validated
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) { [System.IO.File]::Delete($temporaryPath) }
    }
}
```

`ConvertTo-SqlUtilityValidatedConfig` must rebuild a new object instead of returning or mutating the input. Validate exact integral CLR types (`sbyte`, `byte`, `int16`, `uint16`, `int32`, `uint32`, `int64`, `uint64`) before range checks, trim connection fields, and deduplicate with the key `server.ToUpperInvariant() + "`0" + database.ToUpperInvariant()`. `Add-*` replaces an equal pair with the latest trimmed casing; `Remove-*` filters the exact case-insensitive pair. `Read-*` uses `Get-Content -Raw` and `ConvertFrom-Json`, lets malformed JSON throw, and rejects `schemaVersion` values other than `1` with `NotSupportedException`.

- [ ] **Step 5: Ignore runtime configuration artifacts**

Append these root-anchored patterns to `.gitignore`:

```gitignore
/SqlUtility.config.json
/.SqlUtility.config.*.tmp
```

- [ ] **Step 6: Run task and baseline tests**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: both exit `0`; output includes `All configuration tests passed.` and `All SQL Connection POC tests passed.`

- [ ] **Step 7: Commit the configuration boundary**

```powershell
git add .gitignore modules/SqlUtility.Config.ps1 tests/Test-Helpers.ps1 tests/Test-Config.ps1
git commit -m "feat: add app-local configuration persistence"
```

---

### Task 2: Single-Table Read-Only Query Policy

**Files:**

- Create: `modules/SqlUtility.QueryPolicy.ps1`
- Create: `tests/Test-QueryPolicy.ps1`

**Interfaces:**

- Produces: `Test-SqlUtilityQuery -Sql <string> -> PSCustomObject`
- Result properties: `IsValid` (`bool`), `ErrorMessage` (`string`), `NormalizedSql` (`string`), `TableIdentifier` (`string`), `HasOrderBy` (`bool`)
- Contract: invalid results contain blank `NormalizedSql`/`TableIdentifier` and are never sent to `SqlUtility.Database.ps1`.

- [ ] **Step 1: Write the policy matrix as failing tests**

Create `tests/Test-QueryPolicy.ps1` with table-driven accepted and rejected SQL. The accepted set must include:

```powershell
$accepted = @(
    @{ Sql = 'SELECT * FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false },
    @{ Sql = 'SELECT i.Id FROM dbo.Items AS i WHERE i.Enabled = 1 ORDER BY i.Id'; Table = 'dbo.Items'; Ordered = $true },
    @{ Sql = 'SELECT DISTINCT [Type], COUNT(*) AS [Count] FROM [dbo].[Items] GROUP BY [Type] HAVING COUNT(*) > 1 ORDER BY [Type];'; Table = '[dbo].[Items]'; Ordered = $true },
    @{ Sql = "-- SELECT FROM fake`r`nSELECT CASE WHEN Name = 'ORDER BY' THEN 1 ELSE 0 END AS Flag FROM dbo.Items WHERE Note = 'JOIN'"; Table = 'dbo.Items'; Ordered = $false },
    @{ Sql = 'SELECT "ORDER BY" AS [Value] FROM dbo.Items /* JOIN x */ ORDER BY Id'; Table = 'dbo.Items'; Ordered = $true }
)
```

The rejected set must individually cover blank SQL, two statements, CTE, nested `SELECT`, `JOIN`, `APPLY`, comma table source, `UNION`, `INTERSECT`, `EXCEPT`, `INTO`, `TOP`, `OFFSET`, `FETCH`, `EXEC`, DML, DDL, transactions, permissions, admin commands, temp/table variables, four-part/external sources, `OPENQUERY`, and malformed/unclosed string/comment/identifier delimiters. Assert every invalid result has a nonblank user-facing `ErrorMessage` and blank executable fields.

Also assert that a single optional trailing semicolon is removed even when trailing whitespace/comments follow it, keywords inside strings/comments/quoted identifiers do not trigger rejection, parenthesized scalar expressions remain allowed, and only a depth-zero `ORDER BY` sets `HasOrderBy`. Include `SELECT Id FROM dbo.Items ORDER BY Id; -- trailing comment` so ordered paging cannot accidentally append inside a line comment.

- [ ] **Step 2: Run the policy test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
```

Expected: FAIL because the query-policy module is absent.

- [ ] **Step 3: Implement a stateful tokenizer that preserves source positions**

Create private `Get-SqlUtilitySqlTokens`. Scan one character at a time with states `Normal`, `SingleQuote`, `DoubleQuote`, `Bracket`, `LineComment`, and `BlockComment`; double quote/bracket/single-quote escapes stay inside their state. Emit tokens with this shape:

```powershell
[pscustomobject]@{
    Text  = $sql.Substring($start, $length)
    Upper = $sql.Substring($start, $length).ToUpperInvariant()
    Kind  = 'Word' # Word, Identifier, Symbol, Semicolon
    Depth = $depth
    Start = $start
    End   = $start + $length
}
```

Increment/decrement depth only for parentheses in `Normal`; reject negative/unclosed depth and unclosed quote/comment states. Comments and literal contents are skipped for policy classification. Quoted/bracketed identifiers are emitted as one `Identifier` token so reserved words in them are not classified.

- [ ] **Step 4: Implement the explicit version 1 policy**

Build `Test-SqlUtilityQuery` around the tokenizer with a consistent invalid-result helper:

```powershell
function New-SqlUtilityInvalidQueryResult([string] $Message) {
    [pscustomobject]@{
        IsValid = $false; ErrorMessage = $Message; NormalizedSql = ''
        TableIdentifier = ''; HasOrderBy = $false
    }
}
```

Treat one semicolon as optional only when it is the final executable token; remove exactly its source span while retaining trailing whitespace/comments. Any earlier or additional semicolon is a second-statement violation. Then enforce first executable token `SELECT`, one depth-zero `FROM`, and no denied word at any executable token depth. Use this denied set:

```powershell
$denied = @(
    'INSERT','UPDATE','DELETE','MERGE','DROP','ALTER','CREATE','TRUNCATE',
    'EXEC','EXECUTE','DECLARE','SET','USE','GRANT','REVOKE','DENY',
    'BEGIN','COMMIT','ROLLBACK','BACKUP','RESTORE','DBCC','BULK',
    'JOIN','APPLY','UNION','INTERSECT','EXCEPT','INTO','TOP','OFFSET','FETCH',
    'OPENQUERY','OPENROWSET','OPENDATASOURCE'
)
```

Reject every `SELECT` except the first depth-zero token. Parse the source after `FROM` as either `identifier` or `identifier . identifier`, allowing quoted/bracketed identifier tokens, followed by at most one optional table alias (`alias` or `AS alias`); stop only at depth-zero `WHERE`, `GROUP`, `HAVING`, `ORDER`, or end. Any comma, opening parenthesis, variable/temp marker, more than one source-name dot, or extra source token is invalid. Require `GROUP BY` and `ORDER BY` pairs when those clause words occur at depth zero. Return the normalized original source slice and the exact source-table text without its alias; do not rewrite user expressions.

- [ ] **Step 5: Run the policy and baseline suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: all exit `0`; policy output includes `All query policy tests passed.`

- [ ] **Step 6: Commit the isolated query policy**

```powershell
git add modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git commit -m "feat: enforce single-table read-only queries"
```

---

### Task 3: Database Execution, Parsing, and Paging

**Files:**

- Create: `modules/SqlUtility.Database.ps1`
- Create: `tests/Test-Database.ps1`

**Interfaces:**

- Consumes: only validated `NormalizedSql` and `HasOrderBy` output from Task 2.
- Produces: `New-SqlUtilityConnectionString -Server -Database -> string`
- Produces: `Invoke-SqlUtilityConnectionTest -Server -Database [-Executor <scriptblock>]`
- Produces: `Invoke-SqlUtilityOrderedPage -Server -Database -Sql -PageNumber -CommandTimeoutSeconds [-Executor] -> QueryPageResult`
- Produces: `Invoke-SqlUtilityUnorderedQuery -Server -Database -Sql -RowLimit -CommandTimeoutSeconds [-Executor] -> QueryPageResult`
- Produces: `Get-SqlUtilityLocalPage -CachedData <DataTable> -PageNumber <int> -IsComplete <bool> -IsTruncated <bool> -> QueryPageResult`
- Produces: `Invoke-SqlUtilityOrderedRowStream -Server -Database -Sql -CommandTimeoutSeconds -OnSchema -OnRow -ShouldContinue [-StreamExecutor]`
- `QueryPageResult` properties: `Data`, `CachedData`, `PageNumber`, `DisplayedRowCount`, `HasPrevious`, `HasNext`, `IsComplete`, `IsTruncated`.
- Executor contract: `param($ConnectionString,$CommandText,$Parameters,$CommandTimeoutSeconds,$MaximumRows) -> DataTable`.
- Stream executor contract: `param($ConnectionString,$CommandText,$CommandTimeoutSeconds,$OnSchema,$OnRow,$ShouldContinue)`; it owns and disposes SQL resources.

- [ ] **Step 1: Write failing connection and paging tests with injected executors**

Create `tests/Test-Database.ps1`. Verify connection strings with `SqlConnectionStringBuilder`: trimmed source/catalog, `IntegratedSecurity = true`, empty user/password, application name `SQL Utility`, and `ConnectTimeout = 10`.

Use an injected executor that records its arguments and returns a generated `DataTable`:

```powershell
function New-NumberedTable([int] $Count) {
    $table = [System.Data.DataTable]::new('Results')
    [void] $table.Columns.Add('Id', [int])
    1..$Count | ForEach-Object { [void] $table.Rows.Add($_) }
    return (, $table)
}

$ordered = Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' `
    -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' -PageNumber 2 `
    -CommandTimeoutSeconds 120 -Executor $recordingExecutor
Assert-Equal 500 $ordered.Data.Rows.Count 'Ordered page hides sentinel row'
Assert-True $ordered.HasPrevious 'Ordered page two has Previous'
Assert-True $ordered.HasNext 'Ordered sentinel enables Next'
Assert-Equal 500 $recorded.Parameters.Offset 'Ordered offset'
Assert-Equal 501 $recorded.Parameters.FetchCount 'Ordered fetch count'
```

Cover ordered results of 0, 500, and 501 rows; invalid page numbers; unordered results of `limit - 1`, `limit`, and `limit + 1`; removal of only the unordered sentinel; local page 1/page 2 slicing at 500; complete/truncated flags; command-timeout forwarding; executor exception propagation; schema preservation for zero rows; and ordered streaming callback order (`OnSchema` once before `OnRow`).

- [ ] **Step 2: Run the database test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
```

Expected: FAIL because the database module is absent.

- [ ] **Step 3: Implement connection construction and real SQL executors**

Use `System.Data.SqlClient.SqlConnectionStringBuilder` and these fixed settings:

```powershell
$builder.DataSource = $Server.Trim()
$builder.InitialCatalog = $Database.Trim()
$builder.IntegratedSecurity = $true
$builder.ApplicationName = 'SQL Utility'
$builder.ConnectTimeout = 10
```

The default table executor creates `SqlConnection`, `SqlCommand`, and `SqlDataReader`; adds `@Offset`/`@FetchCount` as `SqlDbType.Int`; builds a `DataTable` schema from `reader.GetSchemaTable()`/field names and types; reads no more than `MaximumRows`; and disposes reader, command, and connection in nested `finally` blocks. On reaching a probe boundary, call `command.Cancel()` best-effort before disposal. The connection test executes `SELECT 1;` with a fixed 10-second command timeout via `ExecuteScalar()` and discards the value.

- [ ] **Step 4: Implement neutral page-result construction**

Ordered SQL is exactly:

```powershell
$commandText = $Sql + "`r`nOFFSET @Offset ROWS FETCH NEXT @FetchCount ROWS ONLY"
$parameters = @{ Offset = (($PageNumber - 1) * 500); FetchCount = 501 }
```

Remove the 501st row from the displayed table after recording `HasNext`. Ordered results use `CachedData = $null`, `IsComplete = $false`, and `IsTruncated = $false` because only the requested server page is retained. For unordered execution request `RowLimit + 1`, copy/remove the extra row, retain the bounded table in `CachedData`, and call `Get-SqlUtilityLocalPage` for page 1. Local paging must return an empty schema clone for zero rows; otherwise clone the cached schema and import rows from `($PageNumber - 1) * 500` through the lesser of that page's end or `count - 1`. Complete unordered results use `IsComplete = $true`; sentinel-detected results use `IsTruncated = $true`. Return objects with all eight interface properties on every path; never return WinForms controls or formatting instructions.

- [ ] **Step 5: Implement database-owned ordered streaming**

The default stream executor opens a fresh connection, executes the normalized unpaged SQL, calls `OnSchema` once with objects containing `Name`, `DataType`, and `Ordinal`, then for each row calls `ShouldContinue` and `OnRow` with an `object[]` where SQL null remains `[DBNull]::Value`. `Invoke-SqlUtilityOrderedRowStream` passes through the configured command timeout and never returns the reader or connection.

- [ ] **Step 6: Run database and prior suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: all exit `0`; database output includes `All database tests passed.`

- [ ] **Step 7: Commit the database/result boundary**

```powershell
git add modules/SqlUtility.Database.ps1 tests/Test-Database.ps1
git commit -m "feat: add bounded query execution and paging"
```

---

### Task 4: Dependency-Free Streaming Excel Export

**Files:**

- Create: `modules/SqlUtility.Excel.ps1`
- Create: `tests/Test-Excel.ps1`

**Interfaces:**

- Consumes: neutral schema objects (`Name`, `DataType`, `Ordinal`) and `object[]` rows; never consumes SQL objects.
- Produces: `New-SqlUtilityDataTableRowSource -DataTable <DataTable> -> scriptblock`
- Produces: `Export-SqlUtilityXlsx -DestinationPath <string> -RowSource <scriptblock> -TimeoutSeconds <int> [-MaximumDataRows <int>]`
- Row-source contract: `param($OnSchema,$OnRow,$ShouldContinue)`; invoke schema once, then rows in result order.
- Contract: destination is replaced only after a complete, valid package has closed successfully.

- [ ] **Step 1: Write failing workbook package tests**

Create `tests/Test-Excel.ps1` with a mixed-type `DataTable`: integer, decimal, Boolean, `DateTime`, normal text, formula-looking text (`=1+1`, `+cmd`, `-2+3`, `@name`), and `DBNull`. Export to a test-owned temp directory and inspect the package using `System.IO.Compression.ZipFile` plus `System.Xml.XmlDocument`.

Assert these exact entries exist:

```text
[Content_Types].xml
_rels/.rels
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/styles.xml
xl/worksheets/sheet1.xml
```

Assert workbook sheet name `Results`; header cells use bold style; `pane` has `ySplit="1"`, `topLeftCell="A2"`, and `state="frozen"`; `autoFilter` covers header through the last row/column; numeric and Boolean types are not inline strings; dates use the date style; null produces an empty cell; formula-looking strings use `inlineStr` and no `<f>` element exists.

Add tests for zero data rows, 26/27-column address conversion, cached `DataTable` row source, timeout using a delayed/injected row source, a test `MaximumDataRows = 2` overflow, prior-destination preservation, and deletion of `.tmp` siblings on every failure.

- [ ] **Step 2: Run the Excel test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
```

Expected: FAIL because the Excel module is absent.

- [ ] **Step 3: Implement the minimal Open Packaging Convention package**

Load `System.IO.Compression` and `System.IO.Compression.FileSystem`. Create a unique temporary path beside the destination and open it with `[System.IO.Compression.ZipFile]::Open(..., Create)`. Write the six entries above using UTF-8 `XmlWriter` with indentation disabled. Relationships must point workbook to `xl/workbook.xml`, workbook to `worksheets/sheet1.xml` and `styles.xml`, and content types must declare workbook, worksheet, and styles overrides.

Use style indexes `1` for bold header and `2` for date/time. Write strings as `t="inlineStr"` with an `<is><t xml:space="preserve">value</t></is>` payload; this guarantees formula-looking input is literal. Write Boolean as `t="b"` and `0`/`1`; integral/decimal/floating values as invariant numeric text; `DateTime` as invariant OLE Automation number with style `2`; and database null as an empty `<c>`.

- [ ] **Step 4: Stream rows, enforce limits, and finalize safely**

Start one `Stopwatch` before package creation. `ShouldContinue` throws `TimeoutException` once elapsed total seconds exceed `TimeoutSeconds`; call it while writing headers, before every row, and before final package replacement. Increment a data-row counter before writing and throw when it exceeds `MaximumDataRows` (default `1048575`). Write `<autoFilter ref="A1:{lastColumn}{rowCount+1}"/>` after closing `<sheetData>` so no pre-count is needed.

Close `XmlWriter`, entry streams, and `ZipArchive` before replacing. Use `[System.IO.File]::Replace($temp,$destination,$null)` for an existing destination and `Move` otherwise. A `finally` block deletes the temp when it still exists. Never delete or truncate the destination before successful package closure.

- [ ] **Step 5: Run Excel and all prior suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
```

Expected: all exit `0`; Excel output includes `All Excel export tests passed.`

- [ ] **Step 6: Commit the neutral Excel exporter**

```powershell
git add modules/SqlUtility.Excel.ps1 tests/Test-Excel.ps1
git commit -m "feat: export neutral query results to xlsx"
```

---

### Task 5: Connection Stage, Workspace Shell, and Settings UI

**Files:**

- Create: `SqlUtility.ps1`
- Create: `tests/Test-SqlUtilityUi.ps1`

**Interfaces:**

- Consumes: Tasks 1-4 through explicit dot-sourcing relative to `$PSScriptRoot`.
- Produces: `New-SqlUtilityDefaultServices -> hashtable` of injectable workflow scriptblocks.
- Produces: `New-SqlUtilityMainForm -Config -ConfigPath -Services -> Form`.
- Produces: `Start-SqlUtilityApplication -ConfigPath -Services`, which owns startup configuration recovery and `ShowDialog`.
- Produces: `Set-SqlUtilityBusy -Form -Busy -Message` and `Set-SqlUtilityStage -Form -Stage Connection|Workspace`.
- Form state in `$form.Tag`: `Config`, `ConfigPath`, `Services`, `ActiveServer`, `ActiveDatabase`, `ExecutedQuery`, `CurrentResult`, `CurrentPage`, `IsQueryStale`, `IsBusy`.
- Required service keys for this task: `TestConnection`, `WriteConfig`, `ShowMessage`, `Confirm`.
- `TestConnection` signature: `param($Server,$Database)`; throws on failure.
- `WriteConfig` signature: `param($Path,$Config) -> validated config`; throws on failure.
- `ShowMessage` signature: `param($Text,$Caption,$Icon)`.
- `Confirm` signature: `param($Text,$Caption) -> bool`.

- [ ] **Step 1: Write failing UI shell tests using fake services**

Dot-source `SqlUtility.ps1 -NoGui`, create forms without `ShowDialog`, and locate controls recursively by these stable names:

```text
ConnectionPanel, ServerTextBox, DatabaseTextBox, TestConnectionButton,
ConnectButton, SavedConnectionsList, DeleteConnectionButton,
WorkspacePanel, ActiveConnectionLabel, ChangeConnectionButton,
WorkspaceTabs, QueryTab, SettingsTab, UnorderedLimitNumeric,
QueryExportTimeoutNumeric, SaveSettingsButton, MainStatusLabel
```

Assert connection stage is visible first; server/database are blank; saved pairs appear but selection only fills fields; Test Connection calls the fake once, shows success, and saves; a save failure after successful SQL shows a warning but keeps the pair in current in-memory UI state; Connect repeats test/save then shows workspace with Query selected; connection failure neither saves nor transitions; delete requires confirmation and only updates the list after `WriteConfig` succeeds; settings ranges/labels/defaults are exact; settings enter state only after successful writing; busy state disables re-entry and is restored after fake exceptions.

Assert Change Connection cancellation preserves text/state; confirmation clears SQL/result/page/export state and returns to blank connection inputs while retaining saved pairs/settings. Call `Start-SqlUtilityApplication` with injected messages/confirmation to assert malformed JSON is overwritten with defaults only after confirmation, declining leaves its bytes unchanged, and unsupported schema version shows an error and exits unchanged without offering reset.

- [ ] **Step 2: Run the UI test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: FAIL because `SqlUtility.ps1` is absent.

- [ ] **Step 3: Build the staged WinForms shell and injectable default services**

Start `SqlUtility.ps1` with:

```powershell
param([switch] $NoGui, [string] $ConfigPath)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
. (Join-Path $PSScriptRoot 'modules\SqlUtility.Config.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.QueryPolicy.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.Database.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.Excel.ps1')
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $PSScriptRoot 'SqlUtility.config.json'
}
```

`New-SqlUtilityDefaultServices` wraps module functions and standard `MessageBox` calls. Construct a single resizable form containing two dock-fill panels; switch stages with `Visible` rather than creating another window. Put connection inputs/actions on the left and saved connections/delete on the right, displaying each saved item as `server – database`. Put active connection header above a `TabControl` in workspace. Keep both input fields empty when populating the saved list. Settings help text must say that **Query/Export timeout (seconds)** covers interactive queries and complete Excel export while connection timeout remains fixed and separate.

- [ ] **Step 4: Implement connection, saved-pair, and settings workflows**

Both Test Connection and Connect must trim/validate fields before invoking `TestConnection`; on SQL success create a candidate via `Add-SqlUtilitySavedConnection`, update in-memory config/list, then attempt `WriteConfig`. Test shows a success popup; Connect enters workspace even if persistence warns. Delete builds a candidate config, writes it first, and only then swaps state/list. Settings likewise build a validated candidate, write first, then update state.

For malformed config at startup, `Start-SqlUtilityApplication` catches `Read-SqlUtilityConfig`: ask `Confirm` to reset, write defaults only on confirmation, and otherwise return without showing a form or altering the file. For `NotSupportedException`, show an error and return without offering overwrite. When configuration succeeds, create the form, call `ShowDialog`, and dispose it in `finally`. The script calls this function only inside `if (-not $NoGui)` so dot-sourcing for tests performs no read, write, or UI launch.

- [ ] **Step 5: Run UI shell and all module suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
```

Expected: all exit `0`; UI output includes `All SQL Utility UI tests passed.`

- [ ] **Step 6: Commit the connection/settings application shell**

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git commit -m "feat: add connection and settings workspace"
```

---

### Task 6: Query Execution, Rendering, Paging, and Export Workflow

**Files:**

- Modify: `SqlUtility.ps1`
- Modify: `tests/Test-SqlUtilityUi.ps1`

**Interfaces:**

- Extends default services with `ValidateQuery`, `ExecuteOrderedPage`, `ExecuteUnordered`, `GetLocalPage`, `ExportResult`, and `PromptSavePath`.
- Produces: `Show-SqlUtilityPage -Form -PageResult`, which alone binds `DataTable` and sizes grid columns.
- Produces: `Invoke-SqlUtilityExportWorkflow -State -DestinationPath`, which composes the database row-stream with the neutral Excel exporter.
- UI controls: `SqlEditor`, `ExecuteButton`, `ExportButton`, `PreviousPageButton`, `NextPageButton`, `PageStatusLabel`, `ResultsGrid`.
- Contract: database functions never receive controls; Excel functions never receive SQL connection/command/reader objects.
- Service signatures: `ValidateQuery($Sql)`; `ExecuteOrderedPage($Server,$Database,$Sql,$PageNumber,$TimeoutSeconds)`; `ExecuteUnordered($Server,$Database,$Sql,$RowLimit,$TimeoutSeconds)`; `GetLocalPage($CachedData,$PageNumber,$IsComplete,$IsTruncated)`; `ExportResult($State,$DestinationPath)`; `PromptSavePath() -> string or $null`.

- [ ] **Step 1: Extend UI tests with ordered, unordered, stale, and export scenarios**

Use fake page-result objects with all eight Task 3 properties. Cover:

- Invalid policy: no database call, concise message, cleared result state.
- Ordered execute: page 1 bound, Previous false, Next from sentinel metadata, status `Page 1 - 500 rows`.
- Ordered Next/Previous: executor receives the exact executed SQL snapshot and requested page.
- Complete unordered execute: cached local paging, export enabled, no truncation popup.
- Truncated unordered execute: first configured maximum retained across 500-row local pages, one popup asks for `ORDER BY`, export disabled.
- Editor change after success: sets `IsQueryStale`, disables paging/export, and never changes the executed snapshot.
- Query error/timeout: clears grid/page/export state and restores buttons/cursor.
- Grid: read-only/no add/delete; widths recalculate after bind and never exceed 200 pixels.
- Export: complete unordered passes cached data; ordered passes exact unpaged query snapshot; canceling Save dialog does nothing; success/failure popups are correct.
- Change Connection confirmed: clears editor and results as approved.

- [ ] **Step 2: Run the extended UI test and verify new cases fail**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: FAIL on missing query controls/workflows while Task 5 cases still pass.

- [ ] **Step 3: Build the Query tab and result renderer**

Use a vertical `SplitContainer`: multiline editor/actions/status above and read-only dock-fill `DataGridView` below. Add concise help text: `Paging requires ORDER BY on stable, preferably unique columns.` On `Show-SqlUtilityPage`, bind `PageResult.Data`, update pager/status from metadata, then set every column to `AutoSizeMode = AllCells`, call `AutoResizeColumn`, cap `Width` at `200`, and return `AutoSizeMode` to `None` so the cap remains effective.

The editor `TextChanged` handler compares current text with `State.ExecutedQuery.OriginalEditorSql`. When unequal after a success, set stale and disable Previous, Next, and Export. Do not discard the displayed rows until the next Execute or Change Connection.

- [ ] **Step 4: Implement execution and paging state transitions**

On Execute: clear prior results first, validate, snapshot both original editor text and normalized SQL, then choose ordered or unordered service using active server/database and current settings. Ordered paging always calls the service again with the normalized snapshot and target page. Unordered paging always calls `GetLocalPage` on `CurrentResult.CachedData`; it never queries SQL again. Notify truncation once immediately after execution.

In every failure path set `CurrentResult = $null`, `CurrentPage = 0`, clear grid/status, and disable pager/export. Wrap synchronous work in `try/finally` around `Set-SqlUtilityBusy`; update status and call `$form.Refresh()` before invoking services.

- [ ] **Step 5: Compose ordered and cached export without crossing module boundaries**

Implement `Invoke-SqlUtilityExportWorkflow` with these two row sources:

```powershell
if ($State.ExecutedQuery.HasOrderBy) {
    $rowSource = {
        param($OnSchema, $OnRow, $ShouldContinue)
        Invoke-SqlUtilityOrderedRowStream -Server $State.ActiveServer `
            -Database $State.ActiveDatabase -Sql $State.ExecutedQuery.NormalizedSql `
            -CommandTimeoutSeconds $State.Config.queryExportTimeoutSeconds `
            -OnSchema $OnSchema -OnRow $OnRow -ShouldContinue $ShouldContinue
    }.GetNewClosure()
}
else {
    $rowSource = New-SqlUtilityDataTableRowSource -DataTable $State.CurrentResult.CachedData
}
Export-SqlUtilityXlsx -DestinationPath $DestinationPath -RowSource $rowSource `
    -TimeoutSeconds $State.Config.queryExportTimeoutSeconds
```

The default `ExportResult` service calls `Invoke-SqlUtilityExportWorkflow`. Enable Export only for a non-stale ordered result or a non-stale complete unordered result. `PromptSavePath` uses `SaveFileDialog` with filter `Excel Workbook (*.xlsx)|*.xlsx`, default extension `xlsx`, `AddExtension = true`, and overwrite confirmation. Never retain the selected path after the handler returns.

- [ ] **Step 6: Run the complete module/UI suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
```

Expected: all exit `0` with no failure output.

- [ ] **Step 7: Commit the query workspace**

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git commit -m "feat: add query paging and export workflows"
```

---

### Task 7: Production Launcher, POC Replacement, and End-to-End Verification

**Files:**

- Create: `StartSqlUtility.cmd`
- Create: `tests/Test-Launcher.ps1`
- Create: `tests/Test-All.ps1`
- Delete: `SqlConnectionPoc.ps1`
- Delete: `StartSqlPoc.cmd`
- Delete: `tests/Test-SqlConnectionPoc.ps1`

**Interfaces:**

- `StartSqlUtility.cmd` is the only user launch entry point.
- `tests/Test-All.ps1` is the single automated verification entry point.
- No runtime file outside the approved production structure is required.

- [ ] **Step 1: Write the failing launcher/distribution test**

Create `tests/Test-Launcher.ps1`. Assert the production runtime files exist, the three POC runtime/test files do not, and launcher text contains this exact command shape:

```cmd
@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0SqlUtility.ps1"
set "sqlUtilityExitCode=%ERRORLEVEL%"
endlocal & exit /b %sqlUtilityExitCode%
```

Dot-source `SqlUtility.ps1 -NoGui` to prove module paths resolve from a different current directory. Snapshot all files and process environment entries before/after constructing and disposing a form with an in-memory config and fake services; assert neither changes. Statically scan runtime scripts for registry/environment mutation commands (`Set-ItemProperty`, `New-ItemProperty`, `Remove-ItemProperty`, `SetEnvironmentVariable`, `setx`, or `reg.exe`) and for third-party module imports; none may appear.

- [ ] **Step 2: Run the launcher test and verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
```

Expected: FAIL because the production launcher is absent and POC files still exist.

- [ ] **Step 3: Add the production launcher and aggregate runner**

Create `StartSqlUtility.cmd` with the exact content asserted above. Create `tests/Test-All.ps1` that starts each test in a fresh `powershell.exe` process and fails immediately on a nonzero exit:

```powershell
$tests = @(
    'Test-Config.ps1', 'Test-QueryPolicy.ps1', 'Test-Database.ps1',
    'Test-Excel.ps1', 'Test-SqlUtilityUi.ps1', 'Test-Launcher.ps1'
)
foreach ($test in $tests) {
    $arguments = @('-NoLogo','-NoProfile')
    if ($test -in @('Test-SqlUtilityUi.ps1','Test-Launcher.ps1')) { $arguments += '-STA' }
    $arguments += @('-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot $test))
    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
Write-Host 'All SQL Utility tests passed.'
```

- [ ] **Step 4: Remove the superseded POC files**

Delete only `SqlConnectionPoc.ps1`, `StartSqlPoc.cmd`, and `tests/Test-SqlConnectionPoc.ps1`. Do not touch ignored user files. Confirm the connection-string and blank-field coverage from the POC exists in `Test-Database.ps1` and `Test-SqlUtilityUi.ps1` before deletion.

- [ ] **Step 5: Run automated verification from the repository root and another directory**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
Push-Location $env:TEMP
try {
    powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File 'C:\Users\bluesWalker\Documents\SQL Utility\tests\Test-Launcher.ps1'
}
finally { Pop-Location }
git diff --check
git status --short
```

Expected: `All SQL Utility tests passed.`, the outside-directory launcher test passes, `git diff --check` has no errors, and status lists only the intended Task 7 changes.

- [ ] **Step 6: Perform the Citrix acceptance pass when the environment is available**

Copy the folder to the approved cloud drive and perform this exact checklist without adding automated credentials:

1. Launch `StartSqlUtility.cmd`; confirm blank fields and connection-first view.
2. Test and save one real Windows-authenticated server/database pair; restart and select it from the saved list.
3. Connect, save both boundary and normal settings, restart, and verify persistence.
4. Run a complete unordered query below the configured limit and export all cached rows.
5. Run an unordered query beyond the limit; verify the bounded rows remain visible, the `ORDER BY` notice appears once, and export is disabled.
6. Run an ordered query with a stable unique ordering; verify 500-row Next/Previous navigation.
7. Export the ordered query and inspect bold/filterable/frozen headers and literal formula-looking strings in Excel.
8. Confirm the app folder contains only the approved runtime files, `SqlUtility.config.json`, and no logs/history/results; Excel exists only at the selected path.

Record environment-specific failures as evidence; do not weaken validation, authentication, or safe-write behavior to make the acceptance pass succeed.

- [ ] **Step 7: Commit the production replacement**

```powershell
git add StartSqlUtility.cmd tests/Test-Launcher.ps1 tests/Test-All.ps1
git add -u SqlConnectionPoc.ps1 StartSqlPoc.cmd tests/Test-SqlConnectionPoc.ps1
git commit -m "feat: complete portable SQL Utility v1"
```

- [ ] **Step 8: Run final verification after the commit**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git status --short
git log --oneline -8
```

Expected: aggregate suite exits `0`, the working tree is clean apart from already ignored user files, and the seven implementation commits appear after the approved design/plan commits.

## Implementation Review Gates

For subagent-driven execution, each task uses a fresh implementation subagent followed by two reviews before the next task starts:

1. Specification compliance review against this task, the global constraints, and `docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md`.
2. Code-quality review after compliance issues are resolved.

The primary agent independently inspects the diff and reruns the stated task/full-suite commands before accepting any subagent result or commit.
