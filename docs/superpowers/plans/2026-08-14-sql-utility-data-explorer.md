# SQL Utility Data Explorer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a constrained read-only Data Explorer tab that previews up to a configurable number of physical-table rows, builds typed `AND` filters, exports the displayed preview, and sends generated SQL to the Query tab.

**Architecture:** `SqlUtility.ps1` continues to own WinForms and workflow state. A new UI-neutral `SqlUtility.DataExplorer.ps1` owns filter/type rules plus parameterized preview and editor-SQL generation; `SqlUtility.Database.ps1` owns catalog access and typed command execution; Config and Excel retain their existing persistence/export boundaries.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework, `System.Windows.Forms`, `System.Drawing`, `System.Data.SqlClient`, existing dependency-free OOXML exporter, built-in JSON cmdlets, dependency-free PowerShell test scripts.

## Global Constraints

- Preserve `StartSqlUtility.cmd` and `powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass`.
- Use only Windows PowerShell 5.1 and .NET Framework components normally present in Citrix.
- Add no installer, executable compilation, administrator requirement, third-party module, NuGet package, Python, `sqlcmd`, Office automation, registry write, or environment-variable write.
- Preserve Windows integrated authentication; never request, log, serialize, or persist credentials.
- Keep the application read-only. Do not add any insert, update, delete, DDL, transaction, permission, or administrative workflow.
- Preserve QueryPolicy's approved named-source `SELECT`/JOIN grammar; Data Explorer editor SQL must pass the unchanged policy.
- Limit Data Explorer to physical user tables, `AND` filters, one preview snapshot, no paging/count/sorting, preview-only export, and synchronous execution.
- Keep `previewRowLimit` from `10` through `500`, default `100`; keep existing unordered limit, display-page size, timeouts, and Excel limits unchanged.
- Preview identifiers come only from returned catalog metadata and are bracket-quoted; preview values use typed `SqlParameter` descriptors and never enter command text as literals.
- Persist only schema version 2 config values and saved connections; do not persist Data Explorer tables, metadata, columns, filters, SQL, or preview rows.
- Preserve deterministic cleanup and existing safe destination/config replacement behavior.
- Treat Citrix, live SQL Server, cloud-drive, and desktop Excel verification as external acceptance until actually run there.
- Use regression-first TDD, stage only task-owned files, run focused tests followed by `Test-All.ps1`, and inspect diffs before every commit.

## Planned File Structure

```text
SqlUtility.ps1                                      WinForms and workflows
modules/SqlUtility.Config.ps1                       schema v2 and safe persistence
modules/SqlUtility.QueryPolicy.ps1                  unchanged editor SQL policy
modules/SqlUtility.DataExplorer.ps1                 new UI-neutral builder/query generation
modules/SqlUtility.Database.ps1                     catalog and preview execution
modules/SqlUtility.Excel.ps1                        unchanged workbook engine
tests/Test-Config.ps1                               config migration/settings
tests/Test-DataExplorer.ps1                         new builder/query unit suite
tests/Test-Database.ps1                             catalog/typed execution
tests/Test-SqlUtilityUi.ps1                         explorer workflow
tests/Test-Launcher.ps1                             seven-file distribution
tests/Test-All.ps1                                  aggregate runner
README.md                                           current user/architecture contract
AGENTS.md                                           repository-wide boundaries
docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md
docs/superpowers/specs/2026-08-05-project-documentation-design.md
```

---

### Task 1: Configuration Schema 2 and Preview Limit

**Files:**
- Modify: `modules/SqlUtility.Config.ps1:66-132,238-322`
- Modify: `tests/Test-Config.ps1:7-194`

**Interfaces:**
- Produces: `New-SqlUtilityDefaultConfig -> schemaVersion=2, previewRowLimit=100`
- Produces: `ConvertTo-SqlUtilityValidatedConfig -InputObject <v1|v2> -> validated schemaVersion=2 config`
- Contract: v1 input receives `previewRowLimit=100` in memory; reads never rewrite a file.
- Contract: Add/Remove saved-connection functions preserve `previewRowLimit`.

- [ ] **Step 1: Add failing defaults, migration, and bounds tests**

Update the opening assertions and add explicit v1/v2 cases:

```powershell
$defaults = New-SqlUtilityDefaultConfig
Assert-Equal 2 $defaults.schemaVersion 'Default schema version'
Assert-Equal 100 $defaults.previewRowLimit 'Default preview limit'

$version1 = [pscustomobject][ordered]@{
    schemaVersion = 1
    unorderedRowLimit = 1200
    queryExportTimeoutSeconds = 90
    connections = @([pscustomobject]@{ server = 'ServerA'; database = 'DbA' })
}
$migrated = ConvertTo-SqlUtilityValidatedConfig -InputObject $version1
Assert-Equal 2 $migrated.schemaVersion 'Version 1 migrates in memory'
Assert-Equal 100 $migrated.previewRowLimit 'Version 1 receives default preview limit'
Assert-Equal 1200 $migrated.unorderedRowLimit 'Migration preserves unordered limit'

foreach ($validLimit in @(10, 100, 500)) {
    $candidate = New-SqlUtilityDefaultConfig
    $candidate.previewRowLimit = $validLimit
    Assert-Equal $validLimit (ConvertTo-SqlUtilityValidatedConfig $candidate).previewRowLimit `
        "Accept preview limit $validLimit"
}
foreach ($invalidLimit in @(9, 501, 10.5, '100')) {
    $candidate = New-SqlUtilityDefaultConfig
    $candidate.previewRowLimit = $invalidLimit
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig $candidate } `
        'System.ArgumentException' "Reject preview limit $invalidLimit"
}
```

Add named assertions for each of these contracts: schema 2 missing/invalid `previewRowLimit` is corruption; schema 3 throws `NotSupportedException`; a schema 1 input ignores any extra `previewRowLimit` and receives `100`; `Read-SqlUtilityConfig` migrates v1 without modifying its bytes; write round-trip emits schema 2; Add/Remove connection outputs preserve a nondefault preview limit; pre-commit and post-commit cleanup tests retain all schema 2 fields. Re-run the existing parse-corruption reset and environmental-read-failure cases unchanged to prove migration did not widen when reset is offered or cause read-time writes.

- [ ] **Step 2: Run the config suite and verify the regression fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
```

Expected: FAIL because defaults still use schema version 1 and have no `previewRowLimit`.

- [ ] **Step 3: Implement explicit v1 migration and strict v2 validation**

Use the existing exact-integral helper and make the schema branch explicit:

```powershell
$schemaVersion = Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'schemaVersion'
if (-not (Test-SqlUtilityIntegralValue $schemaVersion)) {
    throw [System.ArgumentException]::new('schemaVersion must be an integer.')
}
if ($schemaVersion -eq 1) {
    $previewRowLimit = 100
}
elseif ($schemaVersion -eq 2) {
    $previewRowLimit = ConvertTo-SqlUtilityValidatedInteger `
        -Value (Get-SqlUtilityConfigPropertyValue $InputObject 'previewRowLimit') `
        -Minimum 10 -Maximum 500 -Name 'previewRowLimit'
}
else {
    throw [System.NotSupportedException]::new("Unsupported configuration schema version: $schemaVersion")
}
```

Return a new ordered object with `schemaVersion=2` and `previewRowLimit`. Add the property to `New-SqlUtilityDefaultConfig` and to candidate objects in `Add-SqlUtilitySavedConnection` and `Remove-SqlUtilitySavedConnection`. Do not change safe-write mechanics or write during `Read-SqlUtilityConfig`.

- [ ] **Step 4: Run focused and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: both suites exit `0`; config output says `All configuration tests passed.` and the aggregate says `All SQL Utility tests passed.`

- [ ] **Step 5: Commit schema 2**

```powershell
git add modules/SqlUtility.Config.ps1 tests/Test-Config.ps1
git diff --cached --check
git commit -m "feat: add data preview configuration"
```

---

### Task 2: UI-Neutral Data Explorer Query Builder

**Files:**
- Create: `modules/SqlUtility.DataExplorer.ps1`
- Create: `tests/Test-DataExplorer.ps1`
- Modify: `tests/Test-All.ps1:1-4`
- Modify: `tests/Test-Launcher.ps1:29-36`

**Interfaces:**
- Consumes table objects: `ObjectId`, `SchemaName`, `TableName`, `DisplayName`.
- Consumes column objects: `Name`, `Ordinal`, `SqlTypeName`, `MaxLength`, `Precision`, `Scale`, `IsNullable`, `IsUserDefined`.
- Consumes filters: `ColumnName`, `Operator`, `ValueText` where Operator is `Equals|NotEquals|Contains|StartsWith|GreaterThan|GreaterThanOrEqual|LessThan|LessThanOrEqual|IsNull|IsNotNull`.
- Produces: `Get-SqlUtilityDataExplorerOperators -Column <object> -> operator objects with Key, Label, RequiresValue`.
- Produces: `Get-SqlUtilityDataExplorerColumnDisplayText -Column <object> -> concise SQL type text such as PlantID (nvarchar(20))`.
- Produces: `New-SqlUtilityDataExplorerQuery -Table -Columns -SelectedColumnNames -Filters -PreviewRowLimit -> query descriptor`.
- Query descriptor properties: `PreviewSql`, `PreviewParameters`, `EditorSql`.
- Parameter descriptor properties: `Name`, `SqlDbType`, `Size`, `Precision`, `Scale`, `Value`.

- [ ] **Step 1: Write the failing builder matrix**

Create `tests/Test-DataExplorer.ps1`, dot-source helpers, QueryPolicy, and the new module. Define neutral fixtures:

```powershell
$table = [pscustomobject]@{ ObjectId=42; SchemaName='odd]schema'; TableName='Order Table'; DisplayName='[odd]]schema].[Order Table]' }
$columns = @(
    [pscustomobject]@{ Name='PlantID'; Ordinal=1; SqlTypeName='nvarchar'; MaxLength=40; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false },
    [pscustomobject]@{ Name='ProductID'; Ordinal=2; SqlTypeName='varchar'; MaxLength=20; Precision=0; Scale=0; IsNullable=$true; IsUserDefined=$false },
    [pscustomobject]@{ Name='OrderDate'; Ordinal=3; SqlTypeName='datetime2'; MaxLength=8; Precision=0; Scale=3; IsNullable=$false; IsUserDefined=$false },
    [pscustomobject]@{ Name='Payload'; Ordinal=4; SqlTypeName='varbinary'; MaxLength=-1; Precision=0; Scale=0; IsNullable=$true; IsUserDefined=$false }
)
$filters = @(
    [pscustomobject]@{ ColumnName='PlantID'; Operator='Equals'; ValueText="CN'01" },
    [pscustomobject]@{ ColumnName='ProductID'; Operator='Contains'; ValueText='A%_[B' },
    [pscustomobject]@{ ColumnName='OrderDate'; Operator='GreaterThanOrEqual'; ValueText='2026-01-01T00:00:00' },
    [pscustomobject]@{ ColumnName='ProductID'; Operator='IsNotNull'; ValueText='' }
)
$query = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns `
    -SelectedColumnNames @('PlantID','ProductID') -Filters $filters -PreviewRowLimit 100
```

Assert bracketed identifiers, `TOP (@PreviewLimit)`, four `AND` predicates, no literal filter values in `PreviewSql`, typed parameters in deterministic order, bracket-escaped LIKE value `%A[%][_][[]B%`, readable escaped literals in `EditorSql`, no `TOP`/`ORDER BY` in editor SQL, and `Test-SqlUtilityQuery($query.EditorSql).IsValid`.

Add one table-driven case for every approved type/operator family. Assert output-only rejection for binary, rowversion/timestamp, XML, spatial, hierarchy, `sql_variant`, CLR/user-defined, and legacy `text`/`ntext`/`image`. Add individually named cases for zero selected columns, unknown/duplicate selected names, unknown columns, invalid operators, invalid numeric/date/bit/GUID values, null operators ignoring `ValueText`, repeated columns, literal quotes/comments/semicolons, `%`, `_`, `[`, `]`, inclusive preview limits, and limits 9/501. Assert column display text for length, `(max)`, precision/scale, and temporal scale. Assert both SQL forms omit `ORDER BY`, and the preview form also omits paging/count/sentinel constructs.

- [ ] **Step 2: Run the new suite and verify it fails**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
```

Expected: FAIL because `modules\SqlUtility.DataExplorer.ps1` does not exist.

- [ ] **Step 3: Implement quoting, LIKE escaping, and operator definitions**

Create UI-neutral helpers and keep all operator keys stable:

```powershell
function ConvertTo-SqlUtilityBracketIdentifier([string] $Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { throw [System.ArgumentException]::new('Identifier cannot be blank.') }
    return '[' + $Name.Replace(']', ']]') + ']'
}

function ConvertTo-SqlUtilityLikePattern([string] $Value, [string] $Operator) {
    $escaped = $Value.Replace('[', '[[]').Replace('%', '[%]').Replace('_', '[_]')
    if ($Operator -eq 'Contains') { return '%' + $escaped + '%' }
    if ($Operator -eq 'StartsWith') { return $escaped + '%' }
    throw [System.ArgumentException]::new("Unsupported LIKE operator: $Operator")
}
```

Implement `Get-SqlUtilityDataExplorerOperators` with the exact spec matrix. Append `IsNull`/`IsNotNull` only when `IsNullable` is true. Return no operators for output-only types or any `IsUserDefined` column.

Implement `Get-SqlUtilityDataExplorerColumnDisplayText` with these formatting rules: character/binary lengths use `(n)` or `(max)`, decimal/numeric use `(precision,scale)`, and `time`/`datetime2`/`datetimeoffset` use `(scale)`; all other supported or output-only types show the base SQL type. For `nchar`/`nvarchar`, convert positive catalog `max_length` bytes to characters by dividing by two.

- [ ] **Step 4: Implement typed query and editor-SQL generation**

`New-SqlUtilityDataExplorerQuery` must:

1. Validate the table/column shape, unique known output names, at least one selected output, filter shape, and `10..500` limit.
2. Resolve columns case-insensitively while preserving catalog casing and ordinal order for output.
3. Convert value text according to this exact map and set parameter facets from metadata:

| SQL types | CLR type | `SqlDbType` |
| --- | --- | --- |
| `char`, `varchar`, `nchar`, `nvarchar` | `String` | matching enum member |
| `tinyint`, `smallint`, `int`, `bigint` | `Byte`, `Int16`, `Int32`, `Int64` | matching enum member |
| `decimal`, `numeric` | `Decimal` | `Decimal` |
| `smallmoney`, `money` | `Decimal` | `SmallMoney`, `Money` |
| `real`, `float` | `Single`, `Double` | `Real`, `Float` |
| `date`, `smalldatetime`, `datetime`, `datetime2` | `DateTime` | matching enum member |
| `time` | `TimeSpan` | `Time` |
| `datetimeoffset` | `DateTimeOffset` | `DateTimeOffset` |
| `bit` | `Boolean` | `Bit` |
| `uniqueidentifier` | `Guid` | `UniqueIdentifier` |

For `char`/`varchar`, parameter `Size` is catalog `max_length`; for `nchar`/`nvarchar`, positive `max_length` is divided by two; `-1` remains `-1`. Decimal/numeric parameters copy `Precision` and `Scale`; `time`/`datetime2`/`datetimeoffset` copy `Scale`. Parse ISO 8601 first, then the current Windows culture. Reject overflow, non-finite floating values, and lossy integer conversion.
4. Add `PreviewLimit` first, then value-bearing `Filter1`, `Filter2`, and so on; null predicates create no parameter.
5. Render preview SQL with parameters and editor SQL with invariant literals.

Use a complete descriptor shape even when optional facets are unused:

```powershell
[pscustomobject][ordered]@{
    Name = 'Filter1'
    SqlDbType = [System.Data.SqlDbType]::NVarChar
    Size = 20
    Precision = [byte] 0
    Scale = [byte] 0
    Value = [string] $typedValue
}
```

Render integers/decimals/floats with invariant culture; render `date` as `yyyy-MM-dd`, date-time values with a `T` separator and full applicable fractional precision, `time` as `HH:mm:ss.fffffff`, and offsets with `zzz`; render GUIDs in `D` form. Use `N'...'` for Unicode text, `'...'` for non-Unicode text, and double every single quote. End editor SQL with one semicolon.

- [ ] **Step 5: Register the seventh runtime file and focused suite**

Add `modules\SqlUtility.DataExplorer.ps1` to `$runtimeFiles` in `Test-Launcher.ps1`. Add `Test-DataExplorer.ps1` after QueryPolicy in `Test-All.ps1`; it does not require STA.

- [ ] **Step 6: Run focused, policy, launcher, and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: all exit `0`; the new suite says `All Data Explorer tests passed.`

- [ ] **Step 7: Commit the isolated builder**

```powershell
git add modules/SqlUtility.DataExplorer.ps1 tests/Test-DataExplorer.ps1 tests/Test-All.ps1 tests/Test-Launcher.ps1
git diff --cached --check
git commit -m "feat: add safe data explorer query builder"
```

---

### Task 3: Catalog Metadata and Typed Preview Execution

**Files:**
- Modify: `modules/SqlUtility.Database.ps1:93-172,243-332`
- Modify: `tests/Test-Database.ps1:1-286`

**Interfaces:**
- Produces: `Get-SqlUtilityPhysicalTables -Server -Database -CommandTimeoutSeconds [-Executor] -> table objects`.
- Produces: `Get-SqlUtilityTableColumns -Server -Database -TableObjectId -CommandTimeoutSeconds [-Executor] -> column objects`.
- Produces: `Invoke-SqlUtilityDataPreview -Server -Database -Query -PreviewRowLimit -CommandTimeoutSeconds [-Executor] -> DataTable`.
- Executor contract becomes `param($ConnectionString,$CommandText,[AllowEmptyCollection()][object[]]$ParameterDescriptors,$CommandTimeoutSeconds,$MaximumRows) -> DataTable`.
- Existing ordered/unordered/connection-test behavior remains unchanged after converting its parameters to descriptors.

- [ ] **Step 1: Add failing catalog and preview tests**

Use injected executors that return catalog-shaped `DataTable` objects and record calls:

```powershell
$tables = Get-SqlUtilityPhysicalTables -Server 's' -Database 'd' -CommandTimeoutSeconds 120 -Executor $tableExecutor
Assert-Equal '[sales].[Order]]Header]' $tables[0].DisplayName 'Table display is safely bracketed'
Assert-True ($script:tableCall.CommandText -match 'sys\.tables') 'Catalog reads sys.tables'
Assert-True ($script:tableCall.CommandText -match 'is_ms_shipped\s*=\s*0') 'Catalog excludes shipped tables'

$columns = Get-SqlUtilityTableColumns -Server 's' -Database 'd' -TableObjectId 42 `
    -CommandTimeoutSeconds 120 -Executor $columnExecutor
Assert-Equal 'PlantID' $columns[0].Name 'Column metadata retains name'
Assert-Equal 'nvarchar' $columns[0].SqlTypeName 'Column metadata exposes base SQL type'
Assert-Equal 42 ($script:columnCall.ParameterDescriptors | Where-Object Name -eq 'TableObjectId').Value `
    'Column metadata uses object-id parameter'
```

Construct a query descriptor from Task 2 and assert Data Preview forwards its SQL/descriptors, timeout, and exact maximum. Test `Add-SqlUtilityCommandParameters` with a fake command/parameter collection for Size, Precision, Scale, null-to-`DBNull`, and typed values; assert the production source contains no `AddWithValue`. Update existing paging assertions from hashtable access to descriptor lookup while preserving exact offset/fetch values. Do not add a production-only connection factory: the existing executor's `try/finally` disposal and `Cancel` paths remain structurally unchanged.

- [ ] **Step 2: Run the database suite and verify new cases fail**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
```

Expected: FAIL on the three missing public functions and typed-descriptor assertions.

- [ ] **Step 3: Generalize the table executor to typed descriptors**

Replace the hard-coded Offset/Fetch loop with one helper:

```powershell
function Add-SqlUtilityCommandParameters(
    $Command,
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]] $ParameterDescriptors
) {
    foreach ($descriptor in @($ParameterDescriptors)) {
        $parameter = $Command.Parameters.Add('@' + [string] $descriptor.Name, $descriptor.SqlDbType)
        if ([int] $descriptor.Size -ne 0) { $parameter.Size = [int] $descriptor.Size }
        if ([byte] $descriptor.Precision -ne 0) { $parameter.Precision = [byte] $descriptor.Precision }
        if ([byte] $descriptor.Scale -ne 0) { $parameter.Scale = [byte] $descriptor.Scale }
        $parameter.Value = if ($null -eq $descriptor.Value) { [DBNull]::Value } else { $descriptor.Value }
    }
}
```

Call it before `ExecuteReader`. Convert existing paging parameters into the same six-property descriptors and use `[object[]] @()` for parameterless connection/unordered calls. Preserve reader schema handling, row caps, cancellation, and every `finally` disposal.

- [ ] **Step 4: Implement fixed catalog queries and neutral mapping**

Use fixed SQL, never interpolated identifiers:

```sql
SELECT t.object_id AS ObjectId, s.name AS SchemaName, t.name AS TableName
FROM sys.tables AS t
INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.is_ms_shipped = 0
ORDER BY s.name, t.name;
```

Use this fixed column query so alias/CLR declarations are identified while ordinary system types expose their base SQL name:

```sql
SELECT t.object_id AS ObjectId,
       s.name AS SchemaName,
       t.name AS TableName,
       c.name AS Name,
       c.column_id AS Ordinal,
       base_type.name AS SqlTypeName,
       c.max_length AS MaxLength,
       c.precision AS [Precision],
       c.scale AS Scale,
       c.is_nullable AS IsNullable,
       CONVERT(bit, CASE WHEN declared_type.is_user_defined = 1
                              OR declared_type.is_assembly_type = 1
                         THEN 1 ELSE 0 END) AS IsUserDefined
FROM sys.tables AS t
INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
INNER JOIN sys.columns AS c ON c.object_id = t.object_id
INNER JOIN sys.types AS declared_type ON declared_type.user_type_id = c.user_type_id
INNER JOIN sys.types AS base_type
    ON base_type.system_type_id = c.system_type_id
   AND base_type.user_type_id = base_type.system_type_id
WHERE t.is_ms_shipped = 0
  AND t.object_id = @TableObjectId
ORDER BY c.column_id;
```

Map rows into the exact Task 2 table/column properties. Compute `DisplayName` with a Database-local neutral formatter that applies the same closing-bracket doubling rule; query construction still uses the Data Explorer module's identifier helper. Catalog and metadata calls pass `MaximumRows=[int]::MaxValue` so all visible tables/columns are returned, while the preview call alone uses the configured bound.

- [ ] **Step 5: Implement the bounded preview service**

Validate `PreviewRowLimit` in `10..500`, build the connection string, and call the injected/default table executor with `Query.PreviewSql`, `Query.PreviewParameters`, configured timeout, and `MaximumRows=PreviewRowLimit`. Return the `DataTable` directly; do not produce page or completeness metadata.

- [ ] **Step 6: Run focused and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: all exit `0`; existing paging/count/streaming assertions remain green.

- [ ] **Step 7: Commit catalog and preview execution**

```powershell
git add modules/SqlUtility.Database.ps1 tests/Test-Database.ps1
git diff --cached --check
git commit -m "feat: add physical table preview services"
```

---

### Task 4: Data Explorer Tab, Catalog, and First Preview

**Files:**
- Modify: `SqlUtility.ps1:13-16,46-153,219-262,368-405,514-539,740-1214`
- Modify: `tests/Test-SqlUtilityUi.ps1:14-273` and append workflow cases
- Modify: `tests/Test-Launcher.ps1` fake-service construction contract

**Interfaces:**
- Adds services: `ListPhysicalTables($Server,$Database,$TimeoutSeconds)`, `GetTableColumns($Server,$Database,$ObjectId,$TimeoutSeconds)`, `BuildDataExplorerQuery($Table,$Columns,$SelectedNames,$Filters,$Limit)`, `ExecuteDataPreview($Server,$Database,$Query,$Limit,$TimeoutSeconds)`.
- Adds controls: `DataExplorerTab`, `TableFilterTextBox`, `RefreshTablesButton`, `PhysicalTablesList`, `OutputColumnsList`, `SelectAllColumnsButton`, `SelectNoColumnsButton`, `PreviewButton`, `PreviewSourceLabel`, `PreviewStatusLabel`, `PreviewGrid`; adds global `PreviewLimitNumeric` to the existing Settings tab.
- Adds state: `DataExplorerTablesLoaded`, `DataExplorerTables`, `DataExplorerBuilder`, `DataExplorerPreview`.
- Produces: `Set-SqlUtilityGridData -Grid -DataTable`, shared by Query and preview grids.

- [ ] **Step 1: Extend fake services and required-control tests**

Add recorder collections/results/errors for the four new services. Use these signatures exactly and add returned fixtures from Tasks 2/3. Extend required names with all Task 4 controls. Assert form construction does not call catalog services.

Add tests that enter workspace, select Data Explorer, and assert one automatic table call with active connection/timeout. Case-insensitive substring filtering makes `order` show both `[dbo].[Orders]` and `[sales].[OrderHistory]`; inputs `%`, `_`, and `*` are treated as literal substring characters, not wildcards. Selecting a table makes no column/preview call.

Cover Refresh precisely: keep the selection when the selected object id still exists, reset its metadata/output/filter builder, clear selection when it disappeared, and leave the preview snapshot untouched in both cases. A failed initial load leaves `DataExplorerTablesLoaded=$false`, shows an error, and allows Refresh to retry. Add the four new service stubs to `Test-Launcher.ps1` so form construction remains side-effect free after `Assert-SqlUtilityServices` expands.

- [ ] **Step 2: Add failing first-preview and setting tests**

Drive controls with `PerformClick()` and assert:

```powershell
$previewButton.PerformClick()
Assert-Equal 1 $recorder.ColumnCalls.Count 'First Preview loads metadata once'
Assert-Equal 1 $recorder.BuildExplorerCalls.Count 'First Preview builds one query'
Assert-Equal 1 $recorder.PreviewCalls.Count 'First Preview executes one bounded query'
Assert-Equal 4 (Get-TestControl $form 'OutputColumnsList').CheckedItems.Count `
    'First Preview selects all columns'
Assert-Equal $recorder.PreviewResult (Get-TestControl $form 'PreviewGrid').DataSource `
    'First Preview binds returned DataTable'
```

Assert later Preview reuses cached metadata, None prevents execution with a warning, and All/None never changes an already displayed snapshot. Assert metadata failure preserves an existing snapshot; if metadata succeeds but row execution fails, the new metadata/all-selected builder remains available while the old snapshot stays displayed. Cover a successful zero-row preview with its schema intact.

Assert the header shows the snapshot's source table plus exact displayed row count and labels the rows unordered. Assert no `ORDER BY`/paging/count controls exist in the tab, `PreviewLimitNumeric` is on Settings with bounds `10..500` and the current config value, and Save Settings writes/preserves all four schema 2 properties. A failed settings write leaves `state.Config.previewRowLimit` unchanged and does not alter the displayed preview.

- [ ] **Step 3: Run the UI suite and verify new cases fail**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: FAIL on missing service keys, controls, state, and workflows while existing UI cases remain green.

- [ ] **Step 4: Add module/service/state boundaries**

Dot-source `modules\SqlUtility.DataExplorer.ps1` between QueryPolicy and Database. Add default wrappers and names to `Assert-SqlUtilityServices`. Initialize state exactly:

```powershell
DataExplorerTablesLoaded = $false
DataExplorerTables = @()
DataExplorerBuilder = [pscustomobject][ordered]@{
    Table = $null; Columns = @(); SelectedColumnNames = @(); Filters = @()
}
DataExplorerPreview = $null
```

Add `previewRowLimit` to the Save Settings candidate and synchronize `PreviewLimitNumeric` only after persistence succeeds.

- [ ] **Step 5: Build the approved tab and shared renderer**

Use the approved inline layout: left table filter/list/Refresh, upper-right checked columns, toolbar, and lower preview grid. Leave the documented filter area for Task 5's concrete controls. `OutputColumnsList` is a normal `CheckedListBox` with `CheckOnClick=$true`, vertical scrolling, and no custom keyboard handler. Wrap each metadata column for display with `Name`, `Column`, and `DisplayText=Get-SqlUtilityDataExplorerColumnDisplayText`; set `DisplayMember='DisplayText'` and preserve ordinal order.

Extract Query grid binding/sizing from `Show-SqlUtilityPage`:

```powershell
function Set-SqlUtilityGridData($Grid, [System.Data.DataTable] $DataTable) {
    $Grid.DataSource = $null
    $Grid.Columns.Clear()
    $Grid.AutoGenerateColumns = $true
    $Grid.DataSource = $DataTable
    foreach ($column in $Grid.Columns) {
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        $Grid.AutoResizeColumn($column.Index, [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
        $width = [Math]::Min(300, $column.Width)
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
        $column.Width = $width
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }
}
```

Keep both grids read-only, cell-selectable, non-editable, and non-orderable.

- [ ] **Step 6: Implement catalog/filter/first-preview workflows**

On the first Data Explorer tab activation, synchronously load tables once. Rebuild the list with ordinal-ignore-case substring filtering and preserve builder selection when filtering hides it; restore the visible selection when the filter is cleared. Table selection resets builder metadata/selections/filters but never the preview snapshot. Keep Preview enabled whenever a table is selected, including before metadata exists; with cached metadata and no checked output, its click handler warns without executing.

On Preview, load columns only when builder Columns is empty, select all, read checked output names, build with empty filters, execute, and replace `DataExplorerPreview` only after success:

```powershell
$candidatePreview = & $executePreview $state.ActiveServer $state.ActiveDatabase `
    $query $state.Config.previewRowLimit $state.Config.queryExportTimeoutSeconds
$state.DataExplorerPreview = [pscustomobject][ordered]@{
    SourceTable = $state.DataExplorerBuilder.Table.DisplayName
    Data = $candidatePreview
}
Set-SqlUtilityGridData -Grid $previewGrid -DataTable $candidatePreview
```

After success, set `PreviewSourceLabel` from the snapshot and `PreviewStatusLabel` to `"<n> rows displayed (unordered)"`, including zero. Failures retain the prior snapshot/grid/source/status labels. Wrap catalog/preview work in the existing busy `try/finally` model. Set `DataExplorerTablesLoaded=$true` only after a successful catalog call so Refresh remains a real retry path.

- [ ] **Step 7: Run focused and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: all exit `0`; existing Query UI behavior remains unchanged.

- [ ] **Step 8: Commit the first-preview workflow**

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1 tests/Test-Launcher.ps1
git diff --cached --check
git commit -m "feat: add on-demand data explorer preview"
```

---

### Task 5: Structured Filters and Query-Tab Handoff

**Files:**
- Modify: `SqlUtility.ps1` Data Explorer helpers, controls, and handlers
- Modify: `tests/Test-SqlUtilityUi.ps1` Data Explorer cases

**Interfaces:**
- Adds controls: `DataExplorerFiltersPanel`, `AddFilterButton`, `ClearFiltersButton`, dynamic `FilterColumnCombo<N>`, `FilterOperatorCombo<N>`, `FilterValueText<N>` or `FilterValueBitCombo<N>`, `RemoveFilterButton<N>`, `SendToQueryButton`.
- Produces: `Add-SqlUtilityDataExplorerFilterRow -Form [-Filter]` and `Get-SqlUtilityDataExplorerFilters -Form -> filter objects`.
- Uses Task 2 operator keys and query descriptor without adding raw SQL input.

- [ ] **Step 1: Write failing filter-row interaction tests**

After first Preview, click Add Filter twice. Assert column choices include all filterable columns including unchecked output columns, exclude Payload, operator choices follow type/nullability, duplicate `OrderDate` selections are permitted, and choosing `IsNull` hides/disables its value control. For a bit column, assert the text input is replaced by a drop-down constrained to `True` and `False` and that the neutral filter emits that selected label as `ValueText`.

Set rows to a date range and text Contains, uncheck a selected output column used by a filter, Preview, and assert Build receives exact ordered filters and independent selected names. Supply invalid date/numeric values and assert no preview call, one validation message, and unchanged grid/snapshot.

- [ ] **Step 2: Write failing Send-to-Query tests**

Assert Send is disabled before metadata/with no selected output. With valid current choices:

```powershell
$sqlEditor.Text = 'SELECT Existing FROM dbo.KeepMe'
$recorder.ConfirmResult = $false
$sendButton.PerformClick()
Assert-Equal 'SELECT Existing FROM dbo.KeepMe' $sqlEditor.Text 'Decline preserves editor'

$recorder.ConfirmResult = $true
$sendButton.PerformClick()
Assert-Equal $recorder.BuildExplorerResult.EditorSql $sqlEditor.Text 'Confirm writes current builder SQL'
Assert-Equal (Get-TestControl $form 'QueryTab') (Get-TestControl $form 'WorkspaceTabs').SelectedTab `
    'Send selects Query tab'
Assert-Equal 0 $recorder.OrderedCalls.Count 'Send does not execute Query tab SQL'
```

Assert blank editor needs no confirmation; builder-generation failure changes no editor/tab/result; current builder SQL is used even when displayed preview came from another table; existing Query result follows its normal TextChanged stale behavior.

- [ ] **Step 3: Run UI tests and verify the new cases fail**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: FAIL on missing filter controls/workflows and Send action.

- [ ] **Step 4: Implement dynamic typed filter rows**

Create each row as a small `TableLayoutPanel`; store its child controls and row number in `Tag`. Populate column objects using DisplayMember, repopulate operators through `Get-SqlUtilityDataExplorerOperators`, and show a value control only when `RequiresValue`. Use a `DropDownList` containing only `True`/`False` for bit columns and a `TextBox` for other filterable types. Remove deletes only that row; Clear removes all rows. Table selection/Refresh clears rows; output checkbox changes do not.

`Get-SqlUtilityDataExplorerFilters` returns only this neutral shape in visual order:

```powershell
[pscustomobject][ordered]@{
    ColumnName = [string] $columnCombo.SelectedItem.Name
    Operator = [string] $operatorCombo.SelectedItem.Key
    ValueText = if (-not $operatorCombo.SelectedItem.RequiresValue) {
        ''
    }
    elseif ($columnCombo.SelectedItem.Column.SqlTypeName -eq 'bit') {
        [string] $valueBitCombo.SelectedItem
    }
    else {
        [string] $valueText.Text
    }
}
```

Do not validate or generate SQL in UI code; Task 2 remains authoritative.

- [ ] **Step 5: Use current columns/filters for later Preview and Send**

Preview reads current checked names and filter rows, calls Build, then Execute; neither successful nor failed builder changes mutate the existing grid before successful execution. Send calls Build with the current preview limit, uses `EditorSql`, confirms only for non-whitespace editor text, assigns editor text, and selects Query without execution.

- [ ] **Step 6: Run focused and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: all exit `0`.

- [ ] **Step 7: Commit filters and handoff**

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git diff --cached --check
git commit -m "feat: add structured preview filters"
```

---

### Task 6: Preview Snapshot Export and Connection Reset

**Files:**
- Modify: `SqlUtility.ps1` services, export workflow, state reset, handlers
- Modify: `tests/Test-SqlUtilityUi.ps1` export/snapshot/reset cases
- Modify: `tests/Test-Launcher.ps1` fake-service construction contract

**Interfaces:**
- Adds service: `ExportPreview($DataTable,$DestinationPath,$TimeoutSeconds)`.
- Adds control: `ExportPreviewButton`.
- Contract: Export Preview exports the exact `DataExplorerPreview.Data` without SQL, regardless of current builder/table/settings.
- Contract: only a successful Preview replaces the snapshot; Change Connection clears it.

- [ ] **Step 1: Write failing snapshot/export tests**

Create Preview A, change columns/filters/table/preview limit, and assert grid/source label/Export remain unchanged. Export and assert:

```powershell
$previewCallsBeforeExport = $recorder.PreviewCalls.Count
$exportButton.PerformClick()
Assert-Equal 1 $recorder.ExportPreviewCalls.Count 'Export Preview calls exporter once'
Assert-Equal $previewA $recorder.ExportPreviewCalls[0].DataTable `
    'Export Preview passes exact displayed DataTable'
Assert-Equal $previewCallsBeforeExport $recorder.PreviewCalls.Count `
    'Export Preview performs no additional SQL'
```

Then execute successful Preview B and assert A is replaced. Make Preview B fail and assert A remains. Assert prompt cancellation does not call exporter; export failure preserves snapshot and existing action state; confirmed Change Connection clears builder/catalog/snapshot/grid/source label and disables export.

- [ ] **Step 2: Run UI tests and verify export/reset cases fail**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: FAIL on missing ExportPreview service/control and reset behavior.

- [ ] **Step 3: Compose preview export from the existing cached-table path**

Add the default service without changing `SqlUtility.Excel.ps1`:

```powershell
ExportPreview = {
    param($DataTable, $DestinationPath, $TimeoutSeconds)
    $rowSource = New-SqlUtilityDataTableRowSource -DataTable $DataTable
    Export-SqlUtilityXlsx -DestinationPath $DestinationPath -RowSource $rowSource `
        -TimeoutSeconds $TimeoutSeconds
}
```

The handler uses the existing Save dialog service, passes the exact snapshot Data, and never consults builder state. Label action/message as Export Preview and include the exported row count. Preserve existing destination/error behavior through Excel.

- [ ] **Step 4: Finalize action states and workspace reset**

Enable Export Preview whenever a snapshot exists and the app is not busy. Keep it enabled through builder/table/settings changes. Preview is enabled when a physical table is selected: if metadata is absent, the action loads it and defaults all output columns; if metadata exists with zero checked outputs, the action warns and does not execute. Send requires metadata plus at least one checked output. Both actions remain disabled while busy, and builder validation still occurs at action time. Extend `Reset-SqlUtilityWorkspaceState` and the confirmation text to clear all Data Explorer transient state and controls only after confirmation. Add the `ExportPreview` stub to `Test-Launcher.ps1` when it becomes a required service.

- [ ] **Step 5: Run UI, Excel, and aggregate verification**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
```

Expected: all exit `0`; Excel production files remain unchanged.

- [ ] **Step 6: Commit preview export and reset**

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1 tests/Test-Launcher.ps1
git diff --cached --check
git commit -m "feat: export displayed data previews"
```

---

### Task 7: Documentation, Distribution Audit, and Final Verification

**Files:**
- Modify: `README.md`
- Modify: `AGENTS.md`
- Modify: `docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md`
- Modify: `docs/superpowers/specs/2026-08-05-project-documentation-design.md`
- Verify: `tests/Test-Launcher.ps1`, `tests/Test-All.ps1`, all runtime/test files

**Interfaces:**
- README becomes the current seven-file runtime/user contract.
- AGENTS assigns `SqlUtility.DataExplorer.ps1` its approved UI-neutral responsibility.
- Existing design/docs link to `2026-08-14-sql-utility-data-explorer-design.md` as the owning extension.
- No merge, push, PR, worktree removal, or external-acceptance claim occurs in this task.

- [ ] **Step 1: Update current user and architecture documentation**

In README add Data Explorer to features/workflow, seven-file distribution, architecture diagram/table, schema version 2 JSON/ranges/migration, table/column/filter behavior, unordered `TOP` performance caveat, preview export distinction, security/data boundaries, test suite, limitations, and external acceptance. State that Query-tab complete export behavior is unchanged.

In AGENTS update the runtime-file count/list, add the DataExplorer module boundary, schema 2 preview setting invariant, preview snapshot/export rules, and focused `Test-DataExplorer.ps1` command. Preserve every unrelated portability, authentication, query-policy, paging, Excel, Git, and acceptance instruction.

- [ ] **Step 2: Update design-document cross-references**

Add a concise extension note/link to the v1 design instead of copying the full Data Explorer contract. Update the project-documentation design's exact distribution and architecture statements from six to seven runtime files and include the new focused test/spec owner. Do not rewrite historical implementation-plan snippets as current behavior.

- [ ] **Step 3: Run documentation consistency searches**

```powershell
rg -n "exactly six|six files|schema version 1|Query and Settings tabs|four required files" README.md AGENTS.md docs\superpowers\specs
rg -n "SqlUtility.DataExplorer.ps1|previewRowLimit|Test-DataExplorer.ps1" README.md AGENTS.md docs\superpowers\specs tests
```

Expected: the first command returns only explicitly historical text that is labeled as such; the second shows current ownership, configuration, and test references.

- [ ] **Step 4: Run complete automated verification**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

Expected: every suite exits `0`, aggregate output says `All SQL Utility tests passed.`, diff check has no output, and status lists only Task 7 documentation changes before staging.

- [ ] **Step 5: Verify launcher behavior from outside the application folder**

```powershell
Push-Location ([System.IO.Path]::GetTempPath())
try {
    powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass `
        -File 'C:\Users\bluesWalker\Documents\SQL Utility\.worktrees\data-explorer\tests\Test-Launcher.ps1'
}
finally { Pop-Location }
```

Expected: `Launcher and distribution tests passed.` with the seven runtime files.

- [ ] **Step 6: Commit synchronized documentation**

```powershell
git add README.md AGENTS.md `
    docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md `
    docs/superpowers/specs/2026-08-05-project-documentation-design.md
git diff --cached --check
git diff --cached --stat
git commit -m "docs: document data explorer workflow"
```

- [ ] **Step 7: Run post-commit verification and report external boundaries**

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
git log --oneline -10
```

Expected: aggregate exits `0`, diff check is silent, worktree is clean, and scoped task commits follow the approved spec/plan commits. Report Citrix, live SQL Server, cloud-drive config reload, and desktop Excel opening as pending external acceptance.

## Implementation Review Gates

For subagent-driven execution, each task uses a fresh implementation subagent followed by:

1. Specification-compliance review against the task, global constraints, and `docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-design.md`.
2. Code-quality review after compliance issues are resolved.
3. Primary-agent inspection of the diff, focused output, aggregate output, staged scope, and commit before starting the next task.

Do not merge, push, open a pull request, or remove the worktree without the user's explicit integration choice.
