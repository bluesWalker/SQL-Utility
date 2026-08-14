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
Assert-Equal "SELECT [PlantID], [ProductID]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [PlantID] = N'CN''01'`r`n  AND [ProductID] LIKE '%A[%][_][[]B%'`r`n  AND [OrderDate] >= '2026-01-01T00:00:00.000'`r`n  AND [ProductID] IS NOT NULL;" $query.EditorSql 'Editor SQL uses readable multi-line escaped invariant literals'
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
    [pscustomobject]@{ ColumnName='WholeSecond'; Operator='Equals'; ValueText='2026-01-02T03:04:05' }
) -PreviewRowLimit 100
Assert-True ($preciseQuery.EditorSql -match "\[Precise\] = '2026-01-02T03:04:05\.1234567'") 'Datetime2 scale 7 editor literal preserves all fractional digits'
Assert-True ($preciseQuery.EditorSql -match "\[WholeSecond\] = '2026-01-02T03:04:05'") 'Datetime2 scale 0 editor literal omits fractional digits'
Assert-Equal 7 $preciseQuery.PreviewParameters[1].Scale 'Datetime2 scale 7 parameter facet is preserved'
Assert-Equal 0 $preciseQuery.PreviewParameters[2].Scale 'Datetime2 scale 0 parameter facet is preserved'
Assert-Throws {
    New-SqlUtilityDataExplorerQuery -Table $table -Columns @($columns[2]) -SelectedColumnNames @('OrderDate') -Filters @(
        [pscustomobject]@{ ColumnName='OrderDate'; Operator='Equals'; ValueText='2026-01-02T03:04:05.1234' }
    ) -PreviewRowLimit 100
} 'System.ArgumentException' 'Datetime2 rejects fractional seconds beyond catalog scale'

$sqlDateTimeColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='datetime'; MaxLength=8; Precision=0; Scale=3; IsNullable=$false; IsUserDefined=$false }
$sqlDateTimeQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($sqlDateTimeColumn) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='2026-01-02T03:04:05.0019' }
) -PreviewRowLimit 100
$sqlDateTimeInput = [datetime]::new(2026, 1, 2, 3, 4, 5).AddTicks(19000)
$sqlDateTimeExpected = [System.Data.SqlTypes.SqlDateTime]::new($sqlDateTimeInput).Value
Assert-Equal $sqlDateTimeExpected $sqlDateTimeQuery.PreviewParameters[1].Value 'Datetime parameter uses the SQL datetime representable value'
Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [Value] = '2026-01-02T03:04:05.003';" $sqlDateTimeQuery.EditorSql 'Datetime editor literal uses the same SQL datetime value'

$smallDateTimeColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='smalldatetime'; MaxLength=4; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
$smallDateTimeBoundary = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($smallDateTimeColumn) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='2079-06-06T23:59:00' }
) -PreviewRowLimit 100
Assert-Equal ([datetime]::new(2079, 6, 6, 23, 59, 0)) $smallDateTimeBoundary.PreviewParameters[1].Value 'Smalldatetime accepts its exact upper minute boundary'
Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [Value] = '2079-06-06T23:59:00.000';" $smallDateTimeBoundary.EditorSql 'Smalldatetime boundary parameter and editor literal are identical'
foreach ($invalidSmallDateTime in @('2026-01-02T03:04:05', '2080-01-01T00:00:00')) {
    Assert-Throws {
        New-SqlUtilityDataExplorerQuery -Table $table -Columns @($smallDateTimeColumn) -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$invalidSmallDateTime }
        ) -PreviewRowLimit 100
    } 'System.ArgumentException' "Smalldatetime rejects non-representable value $invalidSmallDateTime"
}

$timeBoundaryColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='time'; MaxLength=5; Precision=0; Scale=7; IsNullable=$false; IsUserDefined=$false }
$timeBoundaryQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($timeBoundaryColumn) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='23:59:59.9999999' }
) -PreviewRowLimit 100
$timeBoundaryExpected = [timespan]::new(0, 23, 59, 59).Add([timespan]::FromTicks(9999999))
Assert-Equal $timeBoundaryExpected $timeBoundaryQuery.PreviewParameters[1].Value 'Time accepts the last tick before one day'
Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [Value] = '23:59:59.9999999';" $timeBoundaryQuery.EditorSql 'Time boundary parameter and editor literal are identical'
foreach ($invalidTime in @('-03:04:05', '1.03:04:05')) {
    Assert-Throws {
        New-SqlUtilityDataExplorerQuery -Table $table -Columns @($timeBoundaryColumn) -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$invalidTime }
        ) -PreviewRowLimit 100
    } 'System.ArgumentException' "Time rejects out-of-domain value $invalidTime"
}

foreach ($invalidLimit in @('100', 10.5)) {
    Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('PlantID') -Filters @() -PreviewRowLimit $invalidLimit } 'System.ArgumentException' "Non-integral CLR preview limit $invalidLimit is rejected"
}

$startsWithQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns $columns -SelectedColumnNames @('ProductID') -Filters @(
    [pscustomobject]@{ ColumnName='ProductID'; Operator='StartsWith'; ValueText="A]_%['; --" },
    [pscustomobject]@{ ColumnName='ProductID'; Operator='NotEquals'; ValueText='finished' }
) -PreviewRowLimit 100
Assert-Equal "A][_][%][[]'; --%" $startsWithQuery.PreviewParameters[1].Value 'StartsWith escapes LIKE metacharacters while retaining a literal bracket and quote'
Assert-True ($startsWithQuery.PreviewSql -notmatch "finished|A\]_") 'Preview SQL keeps repeated filter values parameterized'
Assert-True ($startsWithQuery.EditorSql -match "LIKE 'A\]\[_\]\[%\]\[\[]''; --%'\r?\n\s+AND \[ProductID\] <> 'finished'") 'Editor SQL safely renders repeated string filters with quote and comment-like content'
Assert-Equal 3 $startsWithQuery.PreviewParameters.Count 'Repeated value filters receive separate deterministic parameters'

# LIKE parameter facets must describe the complete escaped pattern so Preview and editor SQL compare the same value.
foreach ($likeCase in @(
    @{ Type='char'; MaxLength=4; Operator='Contains'; Expected='%[%][_][[]A%'; Prefix=''; DbType=[System.Data.SqlDbType]::VarChar },
    @{ Type='char'; MaxLength=4; Operator='StartsWith'; Expected='[%][_][[]A%'; Prefix=''; DbType=[System.Data.SqlDbType]::VarChar },
    @{ Type='varchar'; MaxLength=4; Operator='Contains'; Expected='%[%][_][[]A%'; Prefix=''; DbType=[System.Data.SqlDbType]::VarChar },
    @{ Type='varchar'; MaxLength=4; Operator='StartsWith'; Expected='[%][_][[]A%'; Prefix=''; DbType=[System.Data.SqlDbType]::VarChar },
    @{ Type='nchar'; MaxLength=8; Operator='Contains'; Expected='%[%][_][[]A%'; Prefix='N'; DbType=[System.Data.SqlDbType]::NVarChar },
    @{ Type='nchar'; MaxLength=8; Operator='StartsWith'; Expected='[%][_][[]A%'; Prefix='N'; DbType=[System.Data.SqlDbType]::NVarChar },
    @{ Type='nvarchar'; MaxLength=8; Operator='Contains'; Expected='%[%][_][[]A%'; Prefix='N'; DbType=[System.Data.SqlDbType]::NVarChar },
    @{ Type='nvarchar'; MaxLength=8; Operator='StartsWith'; Expected='[%][_][[]A%'; Prefix='N'; DbType=[System.Data.SqlDbType]::NVarChar }
)) {
    $likeColumn = [pscustomobject]@{
        Name='Value'; Ordinal=1; SqlTypeName=$likeCase.Type; MaxLength=$likeCase.MaxLength
        Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false
    }
    $likeQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($likeColumn) `
        -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator=$likeCase.Operator; ValueText='%_[A' }
        ) -PreviewRowLimit 100
    Assert-Equal $likeCase.Expected $likeQuery.PreviewParameters[1].Value `
        "$($likeCase.Type) $($likeCase.Operator) Preview preserves the complete escaped boundary pattern"
    Assert-Equal $likeCase.Expected.Length $likeQuery.PreviewParameters[1].Size `
        "$($likeCase.Type) $($likeCase.Operator) parameter size covers wildcards and LIKE escaping"
    Assert-Equal $likeCase.DbType $likeQuery.PreviewParameters[1].SqlDbType `
        "$($likeCase.Type) $($likeCase.Operator) uses an unpadded variable-width LIKE parameter"
    $expectedPredicate = "[Value] LIKE $($likeCase.Prefix)'$($likeCase.Expected)'"
    Assert-True ($likeQuery.EditorSql -match [regex]::Escape($expectedPredicate)) `
        "$($likeCase.Type) $($likeCase.Operator) editor literal matches the Preview pattern exactly"
}

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

# Each expected predicate, literal, and value below is hand-derived from the public SQL contract.
foreach ($case in @(
    @{ Name='char equals'; Type='char'; MaxLength=12; Precision=0; Scale=0; Operator='Equals'; Value='AB'; Preview='[Value] = @Filter1'; Editor="[Value] = 'AB'"; Clr=[string]'AB'; DbType=[System.Data.SqlDbType]::Char; Size=12 },
    @{ Name='varchar not equals'; Type='varchar'; MaxLength=12; Precision=0; Scale=0; Operator='NotEquals'; Value='AB'; Preview='[Value] <> @Filter1'; Editor="[Value] <> 'AB'"; Clr=[string]'AB'; DbType=[System.Data.SqlDbType]::VarChar; Size=12 },
    @{ Name='nchar contains'; Type='nchar'; MaxLength=12; Precision=0; Scale=0; Operator='Contains'; Value='AB'; Preview='[Value] LIKE @Filter1'; Editor="[Value] LIKE N'%AB%'"; Clr=[string]'%AB%'; DbType=[System.Data.SqlDbType]::NVarChar; Size=4 },
    @{ Name='nvarchar starts with'; Type='nvarchar'; MaxLength=12; Precision=0; Scale=0; Operator='StartsWith'; Value='AB'; Preview='[Value] LIKE @Filter1'; Editor="[Value] LIKE N'AB%'"; Clr=[string]'AB%'; DbType=[System.Data.SqlDbType]::NVarChar; Size=3 },
    @{ Name='tinyint greater than'; Type='tinyint'; MaxLength=1; Precision=0; Scale=0; Operator='GreaterThan'; Value='7'; Preview='[Value] > @Filter1'; Editor='[Value] > 7'; Clr=[byte]7; DbType=[System.Data.SqlDbType]::TinyInt; Size=0 },
    @{ Name='smallint greater than or equal'; Type='smallint'; MaxLength=2; Precision=0; Scale=0; Operator='GreaterThanOrEqual'; Value='7'; Preview='[Value] >= @Filter1'; Editor='[Value] >= 7'; Clr=[int16]7; DbType=[System.Data.SqlDbType]::SmallInt; Size=0 },
    @{ Name='int less than'; Type='int'; MaxLength=4; Precision=0; Scale=0; Operator='LessThan'; Value='7'; Preview='[Value] < @Filter1'; Editor='[Value] < 7'; Clr=[int]7; DbType=[System.Data.SqlDbType]::Int; Size=0 },
    @{ Name='bigint less than or equal'; Type='bigint'; MaxLength=8; Precision=0; Scale=0; Operator='LessThanOrEqual'; Value='7'; Preview='[Value] <= @Filter1'; Editor='[Value] <= 7'; Clr=[int64]7; DbType=[System.Data.SqlDbType]::BigInt; Size=0 },
    @{ Name='decimal equals'; Type='decimal'; MaxLength=9; Precision=9; Scale=2; Operator='Equals'; Value='1.25'; Preview='[Value] = @Filter1'; Editor='[Value] = 1.25'; Clr=[decimal]1.25; DbType=[System.Data.SqlDbType]::Decimal; Size=0 },
    @{ Name='numeric not equals'; Type='numeric'; MaxLength=9; Precision=9; Scale=2; Operator='NotEquals'; Value='1.25'; Preview='[Value] <> @Filter1'; Editor='[Value] <> 1.25'; Clr=[decimal]1.25; DbType=[System.Data.SqlDbType]::Decimal; Size=0 },
    @{ Name='smallmoney greater than'; Type='smallmoney'; MaxLength=4; Precision=0; Scale=0; Operator='GreaterThan'; Value='1.25'; Preview='[Value] > @Filter1'; Editor='[Value] > 1.25'; Clr=[decimal]1.25; DbType=[System.Data.SqlDbType]::SmallMoney; Size=0 },
    @{ Name='money less than'; Type='money'; MaxLength=8; Precision=0; Scale=0; Operator='LessThan'; Value='1.25'; Preview='[Value] < @Filter1'; Editor='[Value] < 1.25'; Clr=[decimal]1.25; DbType=[System.Data.SqlDbType]::Money; Size=0 },
    @{ Name='real greater than or equal'; Type='real'; MaxLength=4; Precision=0; Scale=0; Operator='GreaterThanOrEqual'; Value='1.5'; Preview='[Value] >= @Filter1'; Editor='[Value] >= 1.5'; Clr=[single]1.5; DbType=[System.Data.SqlDbType]::Real; Size=0 },
    @{ Name='float less than or equal'; Type='float'; MaxLength=8; Precision=0; Scale=0; Operator='LessThanOrEqual'; Value='1.5'; Preview='[Value] <= @Filter1'; Editor='[Value] <= 1.5'; Clr=[double]1.5; DbType=[System.Data.SqlDbType]::Float; Size=0 },
    @{ Name='date equals'; Type='date'; MaxLength=3; Precision=0; Scale=0; Operator='Equals'; Value='2026-01-02'; Preview='[Value] = @Filter1'; Editor="[Value] = '2026-01-02'"; ClrType=[datetime]; DbType=[System.Data.SqlDbType]::Date; Size=0 },
    @{ Name='smalldatetime not equals'; Type='smalldatetime'; MaxLength=4; Precision=0; Scale=0; Operator='NotEquals'; Value='2026-01-02T03:04:00'; Preview='[Value] <> @Filter1'; Editor="[Value] <> '2026-01-02T03:04:00.000'"; ClrType=[datetime]; DbType=[System.Data.SqlDbType]::SmallDateTime; Size=0 },
    @{ Name='datetime greater than'; Type='datetime'; MaxLength=8; Precision=0; Scale=0; Operator='GreaterThan'; Value='2026-01-02T03:04:05'; Preview='[Value] > @Filter1'; Editor="[Value] > '2026-01-02T03:04:05.000'"; ClrType=[datetime]; DbType=[System.Data.SqlDbType]::DateTime; Size=0 },
    @{ Name='datetime2 less than'; Type='datetime2'; MaxLength=8; Precision=0; Scale=3; Operator='LessThan'; Value='2026-01-02T03:04:05.123'; Preview='[Value] < @Filter1'; Editor="[Value] < '2026-01-02T03:04:05.123'"; ClrType=[datetime]; DbType=[System.Data.SqlDbType]::DateTime2; Size=0 },
    @{ Name='time greater than or equal'; Type='time'; MaxLength=5; Precision=0; Scale=3; Operator='GreaterThanOrEqual'; Value='03:04:05.123'; Preview='[Value] >= @Filter1'; Editor="[Value] >= '03:04:05.1230000'"; ClrType=[timespan]; DbType=[System.Data.SqlDbType]::Time; Size=0 },
    @{ Name='datetimeoffset less than or equal'; Type='datetimeoffset'; MaxLength=10; Precision=0; Scale=3; Operator='LessThanOrEqual'; Value='2026-01-02T03:04:05+08:00'; Preview='[Value] <= @Filter1'; Editor="[Value] <= '2026-01-02T03:04:05.0000000+08:00'"; ClrType=[datetimeoffset]; DbType=[System.Data.SqlDbType]::DateTimeOffset; Size=0 },
    @{ Name='bit equals'; Type='bit'; MaxLength=1; Precision=0; Scale=0; Operator='Equals'; Value='true'; Preview='[Value] = @Filter1'; Editor='[Value] = 1'; Clr=$true; DbType=[System.Data.SqlDbType]::Bit; Size=0 },
    @{ Name='uniqueidentifier not equals'; Type='uniqueidentifier'; MaxLength=16; Precision=0; Scale=0; Operator='NotEquals'; Value='01234567-89ab-cdef-0123-456789abcdef'; Preview='[Value] <> @Filter1'; Editor="[Value] <> '01234567-89ab-cdef-0123-456789abcdef'"; ClrType=[guid]; DbType=[System.Data.SqlDbType]::UniqueIdentifier; Size=0 }
)) {
    $semanticColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$case.Type; MaxLength=$case.MaxLength; Precision=$case.Precision; Scale=$case.Scale; IsNullable=$true; IsUserDefined=$false }
    $semanticQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($semanticColumn) -SelectedColumnNames @('Value') -Filters @([pscustomobject]@{ ColumnName='Value'; Operator=$case.Operator; ValueText=$case.Value }) -PreviewRowLimit 100
    Assert-Equal "SELECT TOP (@PreviewLimit) [Value] FROM [odd]]schema].[Order Table] WHERE $($case.Preview);" $semanticQuery.PreviewSql "$($case.Name) preview predicate is exact"
    Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE $($case.Editor);" $semanticQuery.EditorSql "$($case.Name) invariant editor literal is exact"
    Assert-Equal $case.DbType $semanticQuery.PreviewParameters[1].SqlDbType "$($case.Name) parameter type is exact"
    Assert-Equal $case.Size $semanticQuery.PreviewParameters[1].Size "$($case.Name) parameter size is exact"
    Assert-Equal $case.Precision $semanticQuery.PreviewParameters[1].Precision "$($case.Name) parameter precision is exact"
    Assert-Equal $case.Scale $semanticQuery.PreviewParameters[1].Scale "$($case.Name) parameter scale is exact"
    if ($case.ContainsKey('Clr')) { Assert-Equal $case.Clr $semanticQuery.PreviewParameters[1].Value "$($case.Name) CLR parameter value is exact" }
    else { Assert-True ($semanticQuery.PreviewParameters[1].Value -is $case.ClrType) "$($case.Name) CLR parameter value type is exact" }
}
$nullableValue = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='int'; MaxLength=4; Precision=0; Scale=0; IsNullable=$true; IsUserDefined=$false }
foreach ($nullOperator in @('IsNull','IsNotNull')) {
    $nullableQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($nullableValue) -SelectedColumnNames @('Value') -Filters @([pscustomobject]@{ ColumnName='Value'; Operator=$nullOperator; ValueText='ignored' }) -PreviewRowLimit 100
    $nullPredicate = if ($nullOperator -eq 'IsNull') { '[Value] IS NULL' } else { '[Value] IS NOT NULL' }
    Assert-True ($nullableQuery.PreviewSql -match [regex]::Escape($nullPredicate)) "$nullOperator is available for nullable concrete types"
    Assert-Equal 1 $nullableQuery.PreviewParameters.Count "$nullOperator creates no value parameter"
}
$integerColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='int'; MaxLength=4; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
Assert-Throws { New-SqlUtilityDataExplorerQuery -Table $table -Columns @($integerColumn) -SelectedColumnNames @('Value') -Filters @([pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.2' }) -PreviewRowLimit 100 } 'System.ArgumentException' 'Integer column rejects lossy decimal text'

$boundedDecimal = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='decimal'; MaxLength=5; Precision=5; Scale=2; IsNullable=$false; IsUserDefined=$false }
$decimalBoundaryQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($boundedDecimal) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='999.99' }
) -PreviewRowLimit 100
Assert-Equal ([decimal] 999.99) $decimalBoundaryQuery.PreviewParameters[1].Value 'Decimal accepts the exact metadata precision boundary'
$decimalTrailingZeroQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($boundedDecimal) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.230' }
) -PreviewRowLimit 100
Assert-Equal ([decimal] 1.23) $decimalTrailingZeroQuery.PreviewParameters[1].Value 'Decimal normalizes an exactly representable trailing zero'
Assert-True ($decimalTrailingZeroQuery.EditorSql -match '\[Value\] = 1\.23(?:\r?\n|;)') 'Decimal editor SQL uses the normalized parameter value'
Assert-Throws {
    New-SqlUtilityDataExplorerQuery -Table $table -Columns @($boundedDecimal) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.234' }
    ) -PreviewRowLimit 100
} 'System.ArgumentException' 'Decimal rejects fractional digits beyond catalog scale'
Assert-Throws {
    New-SqlUtilityDataExplorerQuery -Table $table -Columns @($boundedDecimal) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1000.00' }
    ) -PreviewRowLimit 100
} 'System.ArgumentException' 'Decimal rejects values beyond catalog precision'

foreach ($floatingCase in @(
    @{ Type='real'; Input='1.2345678'; Expected='1.23456776'; ClrType=[single] },
    @{ Type='float'; Input='1.2345678901234567'; Expected='1.2345678901234567'; ClrType=[double] }
)) {
    $floatingColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$floatingCase.Type; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
    $floatingQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($floatingColumn) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$floatingCase.Input }
    ) -PreviewRowLimit 100
    Assert-True ($floatingQuery.PreviewParameters[1].Value -is $floatingCase.ClrType) "$($floatingCase.Type) Preview keeps its exact CLR type"
    Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [Value] = $($floatingCase.Expected);" $floatingQuery.EditorSql `
        "$($floatingCase.Type) editor literal uses round-trip-safe invariant formatting"
}

foreach ($moneyBoundary in @(
    @{ Type='money'; Input='922337203685477.5807' },
    @{ Type='money'; Input='-922337203685477.5808' },
    @{ Type='smallmoney'; Input='214748.3647' },
    @{ Type='smallmoney'; Input='-214748.3648' }
)) {
    $moneyColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$moneyBoundary.Type; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
    $moneyQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($moneyColumn) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$moneyBoundary.Input }
    ) -PreviewRowLimit 100
    Assert-Equal $moneyBoundary.Input $moneyQuery.PreviewParameters[1].Value.ToString([System.Globalization.CultureInfo]::InvariantCulture) `
        "$($moneyBoundary.Type) accepts exact SQL Server boundary $($moneyBoundary.Input)"
    Assert-True ($moneyQuery.EditorSql -match [regex]::Escape("[Value] = $($moneyBoundary.Input);")) `
        "$($moneyBoundary.Type) boundary parameter and editor literal are identical"
}

foreach ($moneyType in @('money', 'smallmoney')) {
    $moneyColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$moneyType; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
    $trailingZeroQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($moneyColumn) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.23000' }
    ) -PreviewRowLimit 100
    Assert-Equal ([decimal]1.23) $trailingZeroQuery.PreviewParameters[1].Value "$moneyType accepts harmless trailing fractional zero"
    Assert-True ($trailingZeroQuery.EditorSql -match [regex]::Escape('[Value] = 1.23;')) "$moneyType renders its normalized Preview value"
    Assert-Throws {
        New-SqlUtilityDataExplorerQuery -Table $table -Columns @($moneyColumn) -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.23001' }
        ) -PreviewRowLimit 100
    } 'System.ArgumentException' "$moneyType rejects excess nonzero fractional precision"
}

foreach ($invalidMoney in @(
    @{ Type='money'; Input='922337203685477.5808' },
    @{ Type='money'; Input='-922337203685477.5809' },
    @{ Type='smallmoney'; Input='214748.3648' },
    @{ Type='smallmoney'; Input='-214748.3649' }
)) {
    $moneyColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName=$invalidMoney.Type; MaxLength=8; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
    Assert-Throws {
        New-SqlUtilityDataExplorerQuery -Table $table -Columns @($moneyColumn) -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$invalidMoney.Input }
        ) -PreviewRowLimit 100
    } 'System.ArgumentException' "$($invalidMoney.Type) rejects out-of-range value $($invalidMoney.Input)"
}

$priorCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
try {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-GB')
    $cultureDateColumn = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='date'; MaxLength=3; Precision=0; Scale=0; IsNullable=$false; IsUserDefined=$false }
    $cultureDateQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($cultureDateColumn) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='01/02/2026' }
    ) -PreviewRowLimit 100
    Assert-Equal ([datetime]::new(2026, 2, 1)) $cultureDateQuery.PreviewParameters[1].Value 'Ambiguous slash date follows the active en-GB culture'
    Assert-True ($cultureDateQuery.EditorSql -match "\[Value\] = '2026-02-01'") 'Culture-parsed date renders the same invariant editor value'
}
finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $priorCulture
}

$precision38Integer = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='decimal'; MaxLength=17; Precision=38; Scale=0; IsNullable=$false; IsUserDefined=$false }
foreach ($precision38Value in @(
    '99999999999999999999999999999999999999',
    '-99999999999999999999999999999999999999'
)) {
    $precision38Query = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Integer) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$precision38Value }
    ) -PreviewRowLimit 100
    Assert-True ($precision38Query.PreviewParameters[1].Value -is [System.Data.SqlTypes.SqlDecimal]) "Precision 38 value $precision38Value uses SqlDecimal"
    Assert-Equal $precision38Value $precision38Query.PreviewParameters[1].Value.ToString() "Precision 38 value $precision38Value is parameterized exactly"
    Assert-Equal "SELECT [Value]`r`nFROM [odd]]schema].[Order Table]`r`nWHERE [Value] = $precision38Value;" $precision38Query.EditorSql "Precision 38 value $precision38Value renders identically"
    Assert-True (Test-SqlUtilityQuery -Sql $precision38Query.EditorSql).IsValid "Precision 38 value $precision38Value remains QueryPolicy-compatible"
}

$precision38Scaled = [pscustomobject]@{ Name='Value'; Ordinal=1; SqlTypeName='numeric'; MaxLength=17; Precision=38; Scale=2; IsNullable=$false; IsUserDefined=$false }
$precision38BoundaryText = '999999999999999999999999999999999999.99'
$precision38BoundaryQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Scaled) -SelectedColumnNames @('Value') -Filters @(
    [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$precision38BoundaryText }
) -PreviewRowLimit 100
Assert-Equal $precision38BoundaryText $precision38BoundaryQuery.PreviewParameters[1].Value.ToString() 'Precision 38 scaled boundary is parameterized exactly'
Assert-Equal 38 $precision38BoundaryQuery.PreviewParameters[1].Precision 'Precision 38 scaled boundary keeps the parameter precision facet'
Assert-Equal 2 $precision38BoundaryQuery.PreviewParameters[1].Scale 'Precision 38 scaled boundary keeps the parameter scale facet'
Assert-True ($precision38BoundaryQuery.EditorSql -match ([regex]::Escape("[Value] = $precision38BoundaryText;"))) 'Precision 38 scaled boundary renders exactly'
foreach ($redundantZeroCase in @(
    @{ Column=$precision38Scaled; Input='999999999999999999999999999999999999.990'; Expected='999999999999999999999999999999999999.99' },
    @{ Column=$precision38Scaled; Input='-999999999999999999999999999999999999.990'; Expected='-999999999999999999999999999999999999.99' },
    @{ Column=$precision38Integer; Input='99999999999999999999999999999999999999.0'; Expected='99999999999999999999999999999999999999' },
    @{ Column=$precision38Integer; Input='-99999999999999999999999999999999999999.0'; Expected='-99999999999999999999999999999999999999' }
)) {
    $redundantZeroQuery = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($redundantZeroCase.Column) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$redundantZeroCase.Input }
    ) -PreviewRowLimit 100
    Assert-Equal $redundantZeroCase.Expected $redundantZeroQuery.PreviewParameters[1].Value.ToString() "Redundant fractional zeros preserve exact SQL value $($redundantZeroCase.Input)"
    Assert-True ($redundantZeroQuery.EditorSql -match ([regex]::Escape("[Value] = $($redundantZeroCase.Expected);"))) "Redundant fractional zeros render exact SQL value $($redundantZeroCase.Input)"
}
Assert-Throws {
    New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Integer) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='100000000000000000000000000000000000000' }
    ) -PreviewRowLimit 100
} 'System.ArgumentException' 'Precision 38 rejects a 39-digit value'

$priorCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
try {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    $culturePrecision38Text = '123456789012345678901234567890123456,78'
    $culturePrecision38Query = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Scaled) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$culturePrecision38Text }
    ) -PreviewRowLimit 100
    $culturePrecision38Invariant = '123456789012345678901234567890123456.78'
    Assert-Equal $culturePrecision38Invariant $culturePrecision38Query.PreviewParameters[1].Value.ToString() 'Precision 38 accepts current-culture decimal input exactly'
    Assert-True ($culturePrecision38Query.EditorSql -match ([regex]::Escape("[Value] = $culturePrecision38Invariant;"))) 'Current-culture precision 38 parameter and editor literal are identical'
}
finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $priorCulture
}

$priorCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
try {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
    $groupedPrecision38Text = '999,999,999,999,999,999,999,999,999,999,999,999.99'
    $groupedPrecision38Query = New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Scaled) -SelectedColumnNames @('Value') -Filters @(
        [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText=$groupedPrecision38Text }
    ) -PreviewRowLimit 100
    Assert-Equal $precision38BoundaryText $groupedPrecision38Query.PreviewParameters[1].Value.ToString() 'Valid current-culture grouping preserves the precision 38 value'
    Assert-True ($groupedPrecision38Query.EditorSql -match ([regex]::Escape("[Value] = $precision38BoundaryText;"))) 'Valid grouped precision 38 input renders invariantly'
    Assert-Throws {
        New-SqlUtilityDataExplorerQuery -Table $table -Columns @($precision38Scaled) -SelectedColumnNames @('Value') -Filters @(
            [pscustomobject]@{ ColumnName='Value'; Operator='Equals'; ValueText='1.2,3' }
        ) -PreviewRowLimit 100
    } 'System.ArgumentException' 'Malformed current-culture grouping after the decimal separator is rejected'
}
finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $priorCulture
}

Complete-TestFile 'All Data Explorer tests passed.'
