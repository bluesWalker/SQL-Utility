$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.QueryPolicy.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.DataExplorer.ps1')

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

$query = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('PlantID','ProductID') -Filters $filters -PreviewRowLimit 100
Assert-Equal "SELECT TOP (@PreviewLimit) [PlantID], [ProductID] FROM [odd]]schema].[Order Table] WHERE [PlantID] = @Filter1 AND [ProductID] LIKE @Filter2 AND [OrderDate] >= @Filter3 AND [ProductID] IS NOT NULL;" $query.PreviewSql 'Preview SQL is parameterized with bracket-escaped identifiers and four predicates'
Assert-True ($query.PreviewSql -notmatch "CN'01|A%_\[B|2026-01-01") 'Preview SQL does not contain literal filter values'
Assert-Equal "SELECT [PlantID], [ProductID] FROM [odd]]schema].[Order Table] WHERE [PlantID] = N'CN''01' AND [ProductID] LIKE '%A[%][_][[]B%' AND [OrderDate] >= '2026-01-01T00:00:00.000' AND [ProductID] IS NOT NULL;" $query.EditorSql 'Editor SQL uses readable escaped invariant literals'
Assert-True ($query.EditorSql -notmatch '(?i)\bTOP\b|\bORDER\s+BY\b') 'Editor SQL has no preview or ordering clause'
Assert-True ($query.PreviewSql -notmatch '(?i)\bORDER\s+BY\b|OFFSET|FETCH|COUNT\s*\(|501') 'Preview SQL has no paging count or sentinel construct'
Assert-True (Test-SqlUtilityQuery -Sql $query.EditorSql).IsValid 'Editor SQL remains QueryPolicy-compatible'
Assert-Equal 4 $query.PreviewParameters.Count 'Preview parameters include limit and three value filters'
Assert-Equal 'PreviewLimit' $query.PreviewParameters[0].Name 'Preview limit is first parameter'
Assert-Equal ([System.Data.SqlDbType]::Int) $query.PreviewParameters[0].SqlDbType 'Preview limit is Int'
Assert-Equal 100 $query.PreviewParameters[0].Value 'Preview limit preserves requested value'
Assert-Equal 'Filter1' $query.PreviewParameters[1].Name 'First value filter is deterministic'
Assert-Equal ([System.Data.SqlDbType]::NVarChar) $query.PreviewParameters[1].SqlDbType 'Unicode filter has NVarChar type'
Assert-Equal 20 $query.PreviewParameters[1].Size 'Unicode size converts catalog bytes to characters'
Assert-Equal 'Filter2' $query.PreviewParameters[2].Name 'Second value filter is deterministic'
Assert-Equal '%A[%][_][[]B%' $query.PreviewParameters[2].Value 'LIKE metacharacters are escaped before wildcard wrapping'
Assert-Equal ([System.Data.SqlDbType]::DateTime2) $query.PreviewParameters[3].SqlDbType 'Temporal filter has DateTime2 type'
Assert-Equal 3 $query.PreviewParameters[3].Scale 'Temporal filter copies scale'

foreach ($case in @(
    @{ Name='string equals'; Type='varchar'; Value='abc'; Operator='Equals'; DbType=[System.Data.SqlDbType]::VarChar },
    @{ Name='integer comparison'; Type='int'; Value='42'; Operator='GreaterThan'; DbType=[System.Data.SqlDbType]::Int },
    @{ Name='decimal comparison'; Type='decimal'; Value='1.25'; Operator='LessThan'; DbType=[System.Data.SqlDbType]::Decimal },
    @{ Name='money comparison'; Type='money'; Value='1.25'; Operator='LessThanOrEqual'; DbType=[System.Data.SqlDbType]::Money },
    @{ Name='float comparison'; Type='float'; Value='1.25'; Operator='GreaterThanOrEqual'; DbType=[System.Data.SqlDbType]::Float },
    @{ Name='date comparison'; Type='date'; Value='2026-01-01'; Operator='Equals'; DbType=[System.Data.SqlDbType]::Date },
    @{ Name='time comparison'; Type='time'; Value='12:34:56'; Operator='Equals'; DbType=[System.Data.SqlDbType]::Time },
    @{ Name='offset comparison'; Type='datetimeoffset'; Value='2026-01-01T12:34:56+08:00'; Operator='Equals'; DbType=[System.Data.SqlDbType]::DateTimeOffset },
    @{ Name='bit equals'; Type='bit'; Value='true'; Operator='Equals'; DbType=[System.Data.SqlDbType]::Bit },
    @{ Name='guid equals'; Type='uniqueidentifier'; Value='01234567-89ab-cdef-0123-456789abcdef'; Operator='Equals'; DbType=[System.Data.SqlDbType]::UniqueIdentifier }
)) {
    $column = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$case.Type; MaxLength=20; Precision=12; Scale=3; IsNullable=$true; IsUserDefined=$false }
    $result = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($column) -SelectedColumnNames @('Value') -Filters @([pscustomobject]@{ ColumnName='Value'; Operator=$case.Operator; ValueText=$case.Value }) -PreviewRowLimit 10
    Assert-Equal $case.DbType $result.PreviewParameters[1].SqlDbType "$($case.Name) maps to the correct SqlDbType"
}

foreach ($type in @('varbinary','binary','rowversion','timestamp','xml','geography','geometry','hierarchyid','sql_variant','text','ntext','image')) {
    $column = [pscustomobject]@{ Name='Blocked'; Ordinal=1; SqlTypeName=$type; MaxLength=-1; Precision=0; Scale=0; IsNullable=$true; IsUserDefined=$false }
    Assert-Equal 0 @(Get-SqlUtilityDataExplorerOperators -Column $column).Count "$type has no filter operators"
}
$udt = [pscustomobject]@{ Name='Blocked'; Ordinal=1; SqlTypeName='int'; MaxLength=4; Precision=0; Scale=0; IsNullable=$true; IsUserDefined=$true }
Assert-Equal 0 @(Get-SqlUtilityDataExplorerOperators -Column $udt).Count 'User-defined type has no filter operators'
Assert-Equal 'PlantID (nvarchar(20))' (Get-SqlUtilityDataExplorerColumnDisplayText -Column $columns[0]) 'Unicode display converts bytes to character length'
Assert-Equal 'Payload (varbinary(max))' (Get-SqlUtilityDataExplorerColumnDisplayText -Column $columns[3]) 'Binary max display is concise'
Assert-Equal 'Amount (decimal(12,3))' (Get-SqlUtilityDataExplorerColumnDisplayText -Column ([pscustomobject]@{ Name='Amount'; SqlTypeName='decimal'; MaxLength=0; Precision=12; Scale=3 })) 'Decimal display has precision and scale'
Assert-Equal 'At (datetime2(3))' (Get-SqlUtilityDataExplorerColumnDisplayText -Column ([pscustomobject]@{ Name='At'; SqlTypeName='datetime2'; MaxLength=8; Precision=0; Scale=3 })) 'Temporal display has scale'

foreach ($bad in @(
    @{ Name='zero selected columns'; Selected=@(); Filters=@() },
    @{ Name='unknown selected column'; Selected=@('Missing'); Filters=@() },
    @{ Name='duplicate selected column'; Selected=@('PlantID','plantid'); Filters=@() },
    @{ Name='unknown filter column'; Selected=@('PlantID'); Filters=@([pscustomobject]@{ ColumnName='Missing'; Operator='Equals'; ValueText='x' }) },
    @{ Name='invalid operator'; Selected=@('PlantID'); Filters=@([pscustomobject]@{ ColumnName='PlantID'; Operator='Bogus'; ValueText='x' }) },
    @{ Name='invalid numeric value'; Selected=@('PlantID'); Filters=@([pscustomobject]@{ ColumnName='PlantID'; Operator='GreaterThan'; ValueText='1.2' }) }
)) {
    Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames $bad.Selected -Filters $bad.Filters -PreviewRowLimit 100 } 'System.ArgumentException' "$($bad.Name) is rejected"
}
foreach ($limit in @(9,501)) {
    Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('PlantID') -Filters @() -PreviewRowLimit $limit } 'System.ArgumentException' "Preview limit $limit is rejected"
}
foreach ($limit in @(10,500)) {
    $limitQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('PlantID') -Filters @() -PreviewRowLimit $limit
    Assert-Equal $limit $limitQuery.PreviewParameters[0].Value "Preview limit $limit is accepted inclusively"
}
$nullQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('ProductID') -Filters @([pscustomobject]@{ ColumnName='ProductID'; Operator='IsNull'; ValueText="ignored'; --" }) -PreviewRowLimit 100
Assert-Equal 1 $nullQuery.PreviewParameters.Count 'Null predicate ignores supplied value text'
Assert-True ($nullQuery.EditorSql -match '\[ProductID\] IS NULL') 'Null predicate renders without a literal'

$preciseColumns = @(
    [pscustomobject]@{ Name='Precise'; Ordinal=1; SqlTypeName='datetime2'; MaxLength=8; Precision=0; Scale=7; IsNullable=$false; IsUserDefined=$false },
    [pscustomobject]@{ Name='WholeSecond'; Ordinal=2; SqlTypeName='datetime2'; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
)
$preciseQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns $preciseColumns -SelectedColumnNames @('Precise','WholeSecond') -Filters @(
    [pscustomobject]@{ ColumnName='Precise'; Operator='Equals'; ValueText='2026-01-02T03:04:05.1234567' },
    [pscustomobject]@{ ColumnName='WholeSecond'; Operator='Equals'; ValueText='2026-01-02T03:04:05.7654321' }
) -PreviewRowLimit 100
Assert-True ($preciseQuery.EditorSql -match "\[Precise\] = '2026-01-02T03:04:05\.1234567'") 'Datetime2 scale 7 editor literal preserves all fractional digits'
Assert-True ($preciseQuery.EditorSql -match "\[WholeSecond\] = '2026-01-02T03:04:05'") 'Datetime2 scale 0 editor literal omits fractional digits'
Assert-Equal 7 $preciseQuery.PreviewParameters[1].Scale 'Datetime2 scale 7 parameter facet is preserved'
Assert-Equal 0 $preciseQuery.PreviewParameters[2].Scale 'Datetime2 scale 0 parameter facet is preserved'

foreach ($invalidLimit in @('100', 10.5)) {
    Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('PlantID') -Filters @() -PreviewRowLimit $invalidLimit } 'System.ArgumentException' "Non-integral CLR preview limit $invalidLimit is rejected"
}

$startsWithQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('ProductID') -Filters @(
    [pscustomobject]@{ ColumnName='ProductID'; Operator='StartsWith'; ValueText="A]_%['; --" },
    [pscustomobject]@{ ColumnName='ProductID'; Operator='NotEquals'; ValueText='finished' }
) -PreviewRowLimit 100
Assert-Equal "A][_][%][[]'; --%" $startsWithQuery.PreviewParameters[1].Value 'StartsWith escapes LIKE metacharacters while retaining a literal bracket and quote'
Assert-True ($startsWithQuery.PreviewSql -notmatch "finished|A\]_") 'Preview SQL keeps repeated filter values parameterized'
Assert-True ($startsWithQuery.EditorSql -match "LIKE 'A\]\[_\]\[%\]\[\[]''; --%' AND \[ProductID\] <> 'finished'") 'Editor SQL safely renders repeated string filters with quote and comment-like content'
Assert-Equal 3 $startsWithQuery.PreviewParameters.Count 'Repeated value filters receive separate deterministic parameters'

foreach ($invalid in @(
    @{ Name='invalid date'; Column='OrderDate'; Operator='Equals'; Value='not-a-date' },
    @{ Name='invalid time'; Column='TimeValue'; Operator='Equals'; Value='25:99:99'; Type='time' },
    @{ Name='invalid bit'; Column='BitValue'; Operator='Equals'; Value='maybe'; Type='bit' },
    @{ Name='invalid guid'; Column='GuidValue'; Operator='Equals'; Value='not-a-guid'; Type='uniqueidentifier' },
    @{ Name='non-finite float'; Column='FloatValue'; Operator='Equals'; Value='NaN'; Type='float' },
    @{ Name='integer overflow'; Column='IntValue'; Operator='Equals'; Value='2147483648'; Type='int' }
)) {
    if ($invalid.ContainsKey('Type')) { $invalidColumns = @([pscustomobject]@{ Name=$invalid.Column; Ordinal=1; SqlTypeName=$invalid.Type; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }) } else { $invalidColumns = $columns }
    Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns $invalidColumns -SelectedColumnNames @($invalid.Column) -Filters @([pscustomobject]@{ ColumnName=$invalid.Column; Operator=$invalid.Operator; ValueText=$invalid.Value }) -PreviewRowLimit 100 } 'System.ArgumentException' "$($invalid.Name) is rejected"
}

Complete-TestFile 'All Data Explorer tests passed.'
