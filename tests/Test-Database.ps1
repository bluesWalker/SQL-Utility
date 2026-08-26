$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.Database.ps1')

function New-NumberedTable([int] $Count) {
    $table = [System.Data.DataTable]::new('Results')
    [void] $table.Columns.Add('Id', [int])
    for ($number = 1; $number -le $Count; $number++) {
        [void] $table.Rows.Add($number)
    }
    return (, $table)
}

function Assert-PageResultShape($Result, [string] $Message) {
    $expected = 'Data,CachedData,PageNumber,DisplayedRowCount,HasPrevious,HasNext,IsComplete,IsTruncated'
    $actual = @($Result.PSObject.Properties.Name) -join ','
    Assert-Equal $expected $actual "$Message exposes the neutral page-result contract"
}

function Find-ParameterDescriptor($Descriptors, [string] $Name) {
    return @($Descriptors | Where-Object { $_.Name -eq $Name })[0]
}

$duplicateSchema = New-SqlUtilityResultTable -Columns @(
    [pscustomobject]@{ Name = 'Column 5'; DataType = [bool]; Ordinal = 6 },
    [pscustomobject]@{ Name = 'ID'; DataType = [string]; Ordinal = 2 },
    [pscustomobject]@{ Name = 'Id'; DataType = [int]; Ordinal = 0 },
    [pscustomobject]@{ Name = '   '; DataType = [guid]; Ordinal = 5 },
    [pscustomobject]@{ Name = 'Id (2)'; DataType = [decimal]; Ordinal = 3 },
    [pscustomobject]@{ Name = 'id'; DataType = [long]; Ordinal = 1 },
    [pscustomobject]@{ Name = ''; DataType = [datetime]; Ordinal = 4 }
)
Assert-Equal 'Id,Id (2),Id (3),Id (2) (2),Column 5,Column 6,Column 5 (2)' `
    (@($duplicateSchema.Columns.ColumnName) -join ',') `
    'Result schema makes duplicate, pre-suffixed, and blank names unique in reader ordinal order'
$expectedDuplicateTypes = @([int], [long], [string], [decimal], [datetime], [guid], [bool])
for ($ordinal = 0; $ordinal -lt $expectedDuplicateTypes.Count; $ordinal++) {
    Assert-Equal $ordinal $duplicateSchema.Columns[$ordinal].Ordinal "Result schema preserves column ordinal $ordinal"
    Assert-Equal $expectedDuplicateTypes[$ordinal] $duplicateSchema.Columns[$ordinal].DataType `
        "Result schema preserves CLR type at ordinal $ordinal"
}
Assert-Equal 0 $duplicateSchema.Rows.Count 'Result schema parser does not add data rows'

$boundedReader = [pscustomobject]@{
    FieldCount = 2
    Index = -1
    Rows = @(
        [object[]] @('abc', [byte[]] @(1, 2, 3)),
        [object[]] @('1234567890', [byte[]] @(4, 5, 6))
    )
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetSchemaTable -Value { return $null }
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetName -Value {
    param($Ordinal)
    return @('Name', 'Payload')[$Ordinal]
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetFieldType -Value {
    param($Ordinal)
    return @([string], [byte[]])[$Ordinal]
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name Read -Value {
    $this.Index++
    return $this.Index -lt $this.Rows.Count
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name IsDBNull -Value {
    param($Ordinal)
    return [DBNull]::Value.Equals($this.Rows[$this.Index][$Ordinal])
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetChars -Value {
    param($Ordinal, $DataIndex, $Buffer, $BufferIndex, $Length)
    $value = [string] $this.Rows[$this.Index][$Ordinal]
    if ($null -eq $Buffer) { return [long] $value.Length }
    $copyCount = [Math]::Min([int] $Length, $value.Length - [int] $DataIndex)
    $value.CopyTo([int] $DataIndex, $Buffer, [int] $BufferIndex, $copyCount)
    return [long] $copyCount
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetBytes -Value {
    param($Ordinal, $DataIndex, $Buffer, $BufferIndex, $Length)
    $value = [byte[]] $this.Rows[$this.Index][$Ordinal]
    if ($null -eq $Buffer) { return [long] $value.Length }
    $copyCount = [Math]::Min([int] $Length, $value.Length - [int] $DataIndex)
    [Array]::Copy($value, [int] $DataIndex, $Buffer, [int] $BufferIndex, $copyCount)
    return [long] $copyCount
}
$boundedReader | Add-Member -MemberType ScriptMethod -Name GetValue -Value {
    param($Ordinal)
    return $this.Rows[$this.Index][$Ordinal]
}
$boundedReader.Index = -1
$withinLimitResult = Read-SqlUtilityResultTable -Reader $boundedReader -MaximumRows 1 -MaximumResultDataBytes 32
Assert-Equal 1 $withinLimitResult.Rows.Count 'Result data limit retains a complete row that fits the budget'
Assert-Equal 'abc' $withinLimitResult.Rows[0].Name 'Sequential text read preserves the exact value'
Assert-Equal '1,2,3' (($withinLimitResult.Rows[0].Payload | ForEach-Object { [string] $_ }) -join ',') `
    'Sequential binary read preserves the exact value'

$boundedReader.Index = -1
$boundedResult = $null
$boundedError = $null
try {
    $boundedResult = Read-SqlUtilityResultTable -Reader $boundedReader -MaximumRows 10 -MaximumResultDataBytes 32
}
catch {
    $boundedError = $_.Exception
}
Assert-Equal 'System.InvalidOperationException' $boundedError.GetType().FullName `
    'Result data limit rejects the value that would exceed the allocation budget'
Assert-True ($boundedError.Message -match "column 'Name'") `
    'Result data limit identifies the column that exceeded the budget'
Assert-True ($boundedError.Message -match 'review the selected columns and filters') `
    'Result data limit tells the user how to reduce the result'
Assert-Equal $null $boundedResult 'Result data limit never returns a partial table'

$originalBoundedRows = $boundedReader.Rows
$binaryLimitRow = [object[]]::new(2)
$binaryLimitRow[0] = ''
$binaryLimitRow[1] = [byte[]]::new(33)
$boundedReader.Rows = @(, $binaryLimitRow)
$boundedReader.Index = -1
$binaryLimitError = $null
try {
    $null = Read-SqlUtilityResultTable -Reader $boundedReader -MaximumRows 1 -MaximumResultDataBytes 32
}
catch {
    $binaryLimitError = $_.Exception
}
Assert-Equal 'System.InvalidOperationException' $binaryLimitError.GetType().FullName `
    'Result data limit rejects a binary value that exceeds the budget'
Assert-True ($binaryLimitError.Message -match "column 'Payload'") `
    'Binary result data limit identifies the column that exceeded the budget'
$boundedReader.Rows = $originalBoundedRows

$boundedReader.Index = -1
$boundedCancellation = [pscustomobject]@{ Count = 0 }
Assert-Throws {
    Invoke-SqlUtilityBoundedResultRead -Reader $boundedReader -MaximumRows 10 -MaximumResultDataBytes 32 `
        -Cancel { $boundedCancellation.Count++ }
} 'System.InvalidOperationException' 'Bounded reader propagates the result data limit error'
Assert-Equal 1 $boundedCancellation.Count 'Result data limit cancels remaining database work once'

$connectionString = New-SqlUtilityConnectionString -Server '  server.example  ' -Database '  UtilityDb  '
$connectionBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($connectionString)
Assert-Equal 'server.example' $connectionBuilder.DataSource 'Connection string trims the server'
Assert-Equal 'UtilityDb' $connectionBuilder.InitialCatalog 'Connection string trims the database'
Assert-True $connectionBuilder.IntegratedSecurity 'Connection string uses integrated security'
Assert-Equal '' $connectionBuilder.UserID 'Connection string has no user name'
Assert-Equal '' $connectionBuilder.Password 'Connection string has no password'
Assert-Equal 'SQL Utility' $connectionBuilder.ApplicationName 'Connection string has the application name'
Assert-Equal 10 $connectionBuilder.ConnectTimeout 'Connection string has a ten-second connect timeout'

$script:addedParameters = New-Object System.Collections.Generic.List[object]
$fakeParameters = [pscustomobject]@{}
$fakeParameters | Add-Member -MemberType ScriptMethod -Name Add -Value {
    param($Name, $SqlDbType)
    $parameter = [pscustomobject]@{ Name = $Name; SqlDbType = $SqlDbType; Size = 0; Precision = 0; Scale = 0; Value = $null }
    [void] $script:addedParameters.Add($parameter)
    return $parameter
}
$fakeCommand = [pscustomobject]@{ Parameters = $fakeParameters }
Add-SqlUtilityCommandParameters -Command $fakeCommand -ParameterDescriptors @(
    [pscustomobject]@{ Name = 'Text'; SqlDbType = [System.Data.SqlDbType]::NVarChar; Size = 40; Precision = 0; Scale = 0; Value = 'typed' },
    [pscustomobject]@{ Name = 'Amount'; SqlDbType = [System.Data.SqlDbType]::Decimal; Size = 0; Precision = 12; Scale = 3; Value = [decimal] 1.25 },
    [pscustomobject]@{ Name = 'Missing'; SqlDbType = [System.Data.SqlDbType]::Int; Size = 0; Precision = 0; Scale = 0; Value = $null }
)
Assert-Equal '@Text' $script:addedParameters[0].Name 'Typed parameters add the SQL marker'
Assert-Equal 40 $script:addedParameters[0].Size 'Typed parameters preserve Size'
Assert-Equal 12 $script:addedParameters[1].Precision 'Typed parameters preserve Precision'
Assert-Equal 3 $script:addedParameters[1].Scale 'Typed parameters preserve Scale'
Assert-Equal ([decimal] 1.25) $script:addedParameters[1].Value 'Typed parameters preserve typed values'
Assert-True ([DBNull]::Value.Equals($script:addedParameters[2].Value)) 'Typed parameters convert null to DBNull'
$databaseSource = Get-Content -Raw (Join-Path $projectRoot 'modules\SqlUtility.Database.ps1')
Assert-True ($databaseSource -notmatch 'AddWithValue') 'Database parameters never use AddWithValue'

$script:tableCall = $null
$tableExecutor = {
    param($ConnectionString, $CommandText, $ParameterDescriptors, $CommandTimeoutSeconds, $MaximumRows)
    $script:tableCall = [pscustomobject]@{ CommandText = $CommandText; ParameterDescriptors = $ParameterDescriptors; CommandTimeoutSeconds = $CommandTimeoutSeconds; MaximumRows = $MaximumRows }
    $table = [System.Data.DataTable]::new()
    [void] $table.Columns.Add('ObjectId', [int]); [void] $table.Columns.Add('SchemaName', [string]); [void] $table.Columns.Add('TableName', [string])
    [void] $table.Rows.Add(42, 'sales', 'Order]Header')
    return (, $table)
}
$tables = Get-SqlUtilityPhysicalTables -Server 's' -Database 'd' -CommandTimeoutSeconds 120 -Executor $tableExecutor
Assert-Equal '[sales].[Order]]Header]' $tables[0].DisplayName 'Table display is safely bracketed'
Assert-True ($script:tableCall.CommandText -match 'sys\.tables') 'Catalog reads sys.tables'
Assert-True ($script:tableCall.CommandText -match 'is_ms_shipped\s*=\s*0') 'Catalog excludes shipped tables'
Assert-Equal ([int]::MaxValue) $script:tableCall.MaximumRows 'Catalog returns all visible physical tables'

$script:columnCall = $null
$columnExecutor = {
    param($ConnectionString, $CommandText, $ParameterDescriptors, $CommandTimeoutSeconds, $MaximumRows)
    $script:columnCall = [pscustomobject]@{ CommandText = $CommandText; ParameterDescriptors = $ParameterDescriptors; CommandTimeoutSeconds = $CommandTimeoutSeconds; MaximumRows = $MaximumRows }
    $table = [System.Data.DataTable]::new()
    foreach ($column in @(@('ObjectId',[int]),@('SchemaName',[string]),@('TableName',[string]),@('Name',[string]),@('Ordinal',[int]),@('SqlTypeName',[string]),@('MaxLength',[int]),@('Precision',[byte]),@('Scale',[byte]),@('IsNullable',[bool]),@('IsUserDefined',[bool]))) { [void] $table.Columns.Add($column[0], $column[1]) }
    [void] $table.Rows.Add(42, 'dbo', 'Plants', 'PlantID', 1, 'nvarchar', 100, 0, 0, $false, $false)
    [void] $table.Rows.Add(42, 'dbo', 'Plants', 'Shape', 2, 'geometry', -1, 0, 0, $true, $true)
    return (, $table)
}
$columns = Get-SqlUtilityTableColumns -Server 's' -Database 'd' -TableObjectId 42 -CommandTimeoutSeconds 120 -Executor $columnExecutor
Assert-Equal 'PlantID' $columns[0].Name 'Column metadata retains name'
Assert-Equal 'nvarchar' $columns[0].SqlTypeName 'Column metadata exposes base SQL type'
Assert-True ($script:columnCall.CommandText -match '(?s)COALESCE\s*\(\s*base_type\.name\s*,\s*declared_type\.name\s*\)\s+AS\s+SqlTypeName') 'Column metadata falls back to the declared type name'
Assert-True ($script:columnCall.CommandText -match '(?s)LEFT\s+JOIN\s+sys\.types\s+AS\s+base_type') 'Column metadata retains declarations without a canonical base-type row'
Assert-True ($script:columnCall.CommandText -match 'declared_type\.is_assembly_type\s*=\s*1') 'Column metadata keeps assembly types marked as output-only'
Assert-Equal 2 $columns.Count 'Column metadata retains ordinary and assembly-backed columns'
Assert-Equal 'geometry' $columns[1].SqlTypeName 'Assembly-backed metadata exposes the declared type fallback'
Assert-Equal $true $columns[1].IsUserDefined 'Assembly-backed metadata remains output-only'
Assert-Equal 42 (Find-ParameterDescriptor $script:columnCall.ParameterDescriptors 'TableObjectId').Value 'Column metadata uses object-id parameter'

$script:previewCall = $null
$previewTable = New-NumberedTable 2
$previewQuery = [pscustomobject]@{
    PreviewSql = 'SELECT TOP (@PreviewRowLimit) [PlantID] FROM [dbo].[Plants];'
    PreviewParameters = [object[]] @([pscustomobject]@{ Name = 'PreviewRowLimit'; SqlDbType = [System.Data.SqlDbType]::Int; Size = 0; Precision = 0; Scale = 0; Value = 25 })
    EditorSql = 'SELECT [PlantID] FROM [dbo].[Plants];'
}
$preview = Invoke-SqlUtilityDataPreview -Server 's' -Database 'd' -Query $previewQuery -PreviewRowLimit 25 `
    -CommandTimeoutSeconds 77 -ResultDataLimitMiB 512 -Executor {
    param($ConnectionString, $CommandText, $ParameterDescriptors, $CommandTimeoutSeconds, $MaximumRows, $MaximumResultDataBytes)
    $script:previewCall = [pscustomobject]@{ CommandText = $CommandText; ParameterDescriptors = $ParameterDescriptors; CommandTimeoutSeconds = $CommandTimeoutSeconds; MaximumRows = $MaximumRows; MaximumResultDataBytes = $MaximumResultDataBytes }
    return (, $previewTable)
}
Assert-True ([object]::ReferenceEquals($previewTable, $preview)) 'Preview returns executor DataTable directly'
Assert-Equal $previewQuery.PreviewSql $script:previewCall.CommandText 'Preview forwards descriptor SQL'
Assert-True ([object]::ReferenceEquals($previewQuery.PreviewParameters, $script:previewCall.ParameterDescriptors)) 'Preview forwards typed descriptors'
Assert-Equal 77 $script:previewCall.CommandTimeoutSeconds 'Preview forwards timeout'
Assert-Equal 25 $script:previewCall.MaximumRows 'Preview uses exact configured maximum'
Assert-Equal 536870912 $script:previewCall.MaximumResultDataBytes 'Preview forwards the configured result data limit in bytes'
foreach ($invalidLimit in @(9, 501)) {
    Assert-Throws { Invoke-SqlUtilityDataPreview -Server 's' -Database 'd' -Query $previewQuery -PreviewRowLimit $invalidLimit -CommandTimeoutSeconds 5 -Executor { throw 'must not run' } } `
        'System.ArgumentOutOfRangeException' 'Preview rejects limits outside 10 through 500'
}
foreach ($nonIntegralLimit in @('25', 25.4)) {
    Assert-Throws { Invoke-SqlUtilityDataPreview -Server 's' -Database 'd' -Query $previewQuery -PreviewRowLimit $nonIntegralLimit -CommandTimeoutSeconds 5 -Executor { throw 'must not run' } } `
        'System.ArgumentException' 'Preview rejects values that are not integral CLR numeric types'
}

$script:connectionTestCall = $null
$connectionTestExecutor = {
    param($ConnectionString, $CommandText, $Parameters, $CommandTimeoutSeconds, $MaximumRows)
    $script:connectionTestCall = [pscustomobject]@{
        ConnectionString = $ConnectionString
        CommandText = $CommandText
        Parameters = $Parameters
        CommandTimeoutSeconds = $CommandTimeoutSeconds
        MaximumRows = $MaximumRows
    }
    return (, (New-NumberedTable 1))
}
$connectionTestResult = Invoke-SqlUtilityConnectionTest -Server ' server ' -Database ' db ' -Executor $connectionTestExecutor
Assert-Equal $null $connectionTestResult 'Connection test does not expose executor data'
$testedBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($script:connectionTestCall.ConnectionString)
Assert-Equal 'server' $testedBuilder.DataSource 'Connection test forwards the built connection string'
Assert-Equal 'db' $testedBuilder.InitialCatalog 'Connection test forwards the built database catalog'
Assert-Equal 'SELECT 1;' $script:connectionTestCall.CommandText 'Connection test uses the fixed probe'
Assert-Equal 0 $script:connectionTestCall.Parameters.Count 'Connection test has no SQL parameters'
Assert-Equal 10 $script:connectionTestCall.CommandTimeoutSeconds 'Connection test uses the fixed command timeout'
Assert-Equal 1 $script:connectionTestCall.MaximumRows 'Connection test bounds the injected probe'

foreach ($orderedCount in @(0, 500, 501)) {
    $script:orderedCall = $null
    $orderedExecutor = {
        param($ConnectionString, $CommandText, $Parameters, $CommandTimeoutSeconds, $MaximumRows, $MaximumResultDataBytes)
        $script:orderedCall = [pscustomobject]@{
            ConnectionString = $ConnectionString
            CommandText = $CommandText
            Parameters = $Parameters
            CommandTimeoutSeconds = $CommandTimeoutSeconds
            MaximumRows = $MaximumRows
            MaximumResultDataBytes = $MaximumResultDataBytes
        }
        return (, (New-NumberedTable $orderedCount))
    }

    $ordered = Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' `
        -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' -PageNumber 2 `
        -CommandTimeoutSeconds 120 -ResultDataLimitMiB 384 -Executor $orderedExecutor

    $expectedDisplayed = [Math]::Min($orderedCount, 500)
    Assert-Equal $expectedDisplayed $ordered.Data.Rows.Count "Ordered $orderedCount-row result hides only a sentinel"
    Assert-Equal $expectedDisplayed $ordered.DisplayedRowCount "Ordered $orderedCount-row result reports its displayed count"
    Assert-True $ordered.HasPrevious "Ordered $orderedCount-row page two has Previous"
    Assert-Equal ($orderedCount -eq 501) $ordered.HasNext "Ordered $orderedCount-row result derives Next from the sentinel"
    Assert-Equal $null $ordered.CachedData "Ordered $orderedCount-row result is not cached"
    Assert-Equal 2 $ordered.PageNumber "Ordered $orderedCount-row result keeps the page number"
    Assert-Equal $false $ordered.IsComplete "Ordered $orderedCount-row result is not a complete client cache"
    Assert-Equal $false $ordered.IsTruncated "Ordered $orderedCount-row result is not a truncated unordered cache"
    Assert-Equal 1 $ordered.Data.Columns.Count "Ordered $orderedCount-row result preserves schema"
    Assert-Equal 'Id' $ordered.Data.Columns[0].ColumnName "Ordered $orderedCount-row result preserves the column name"
    Assert-Equal ([int]) $ordered.Data.Columns[0].DataType "Ordered $orderedCount-row result preserves the column type"
    Assert-PageResultShape $ordered "Ordered $orderedCount-row result"
}

Assert-Equal "SELECT Id FROM dbo.Items ORDER BY Id`r`nOFFSET @Offset ROWS FETCH NEXT @FetchCount ROWS ONLY" $script:orderedCall.CommandText 'Ordered paging appends OFFSET/FETCH exactly'
Assert-Equal 500 (Find-ParameterDescriptor $script:orderedCall.Parameters 'Offset').Value 'Ordered page two sends the correct offset'
Assert-Equal 501 (Find-ParameterDescriptor $script:orderedCall.Parameters 'FetchCount').Value 'Ordered paging probes one sentinel row'
Assert-Equal 120 $script:orderedCall.CommandTimeoutSeconds 'Ordered paging forwards the command timeout'
Assert-Equal 501 $script:orderedCall.MaximumRows 'Ordered paging bounds the executor read'
Assert-Equal 402653184 $script:orderedCall.MaximumResultDataBytes 'Ordered paging forwards the configured result data limit in bytes'

$orderedFirst = Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' `
    -PageNumber 1 -CommandTimeoutSeconds 5 -Executor { param($a, $b, $c, $d, $e) return (, (New-NumberedTable 0)) }
Assert-Equal $false $orderedFirst.HasPrevious 'Ordered page one has no Previous'
Assert-Throws {
    Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' `
        -PageNumber 0 -CommandTimeoutSeconds 5 -Executor { throw 'must not run' }
} 'System.ArgumentOutOfRangeException' 'Ordered paging rejects page zero'

$executorFailure = [System.InvalidOperationException]::new('executor failed')
Assert-Throws {
    Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' `
        -PageNumber 1 -CommandTimeoutSeconds 5 -Executor { throw $executorFailure }
} 'System.InvalidOperationException' 'Ordered paging propagates executor exceptions'

foreach ($unorderedCount in @(999, 1000, 1001)) {
    $script:unorderedCall = $null
    $sourceTable = New-NumberedTable $unorderedCount
    $unorderedExecutor = {
        param($ConnectionString, $CommandText, $Parameters, $CommandTimeoutSeconds, $MaximumRows, $MaximumResultDataBytes)
        $script:unorderedCall = [pscustomobject]@{
            ConnectionString = $ConnectionString
            CommandText = $CommandText
            Parameters = $Parameters
            CommandTimeoutSeconds = $CommandTimeoutSeconds
            MaximumRows = $MaximumRows
            MaximumResultDataBytes = $MaximumResultDataBytes
        }
        return (, $sourceTable)
    }

    $unordered = Invoke-SqlUtilityUnorderedQuery -Server 's' -Database 'd' `
        -Sql 'SELECT Id FROM dbo.Items' -RowLimit 1000 -CommandTimeoutSeconds 321 `
        -ResultDataLimitMiB 640 -Executor $unorderedExecutor

    $expectedCached = [Math]::Min($unorderedCount, 1000)
    Assert-Equal 500 $unordered.Data.Rows.Count "Unordered $unorderedCount-row result displays the first local page"
    Assert-Equal $expectedCached $unordered.CachedData.Rows.Count "Unordered $unorderedCount-row result retains a bounded cache"
    Assert-Equal $unorderedCount $sourceTable.Rows.Count "Unordered $unorderedCount-row execution does not mutate executor-owned data"
    Assert-Equal ($unorderedCount -le 1000) $unordered.IsComplete "Unordered $unorderedCount-row result reports completeness"
    Assert-Equal ($unorderedCount -eq 1001) $unordered.IsTruncated "Unordered $unorderedCount-row result reports sentinel truncation"
    Assert-Equal ($expectedCached -gt 500) $unordered.HasNext "Unordered $unorderedCount-row result exposes local page two"
    Assert-Equal $false $unordered.HasPrevious "Unordered $unorderedCount-row result starts on page one"
    Assert-PageResultShape $unordered "Unordered $unorderedCount-row result"
}

Assert-Equal 'SELECT Id FROM dbo.Items' $script:unorderedCall.CommandText 'Unordered execution keeps normalized SQL unpaged'
Assert-Equal 0 $script:unorderedCall.Parameters.Count 'Unordered execution sends no parameters'
Assert-Equal 321 $script:unorderedCall.CommandTimeoutSeconds 'Unordered execution forwards the command timeout'
Assert-Equal 1001 $script:unorderedCall.MaximumRows 'Unordered execution requests the limit plus one'
Assert-Equal 671088640 $script:unorderedCall.MaximumResultDataBytes 'Unordered execution forwards the configured result data limit in bytes'
Assert-Equal 1000 $unordered.CachedData.Rows[999].Id 'Unordered execution removes only the sentinel row'

$cached = New-NumberedTable 750
$localFirst = Get-SqlUtilityLocalPage -CachedData $cached -PageNumber 1 -IsComplete $true -IsTruncated $false
Assert-Equal 500 $localFirst.Data.Rows.Count 'Local page one has 500 rows'
Assert-Equal 1 $localFirst.Data.Rows[0].Id 'Local page one starts at the first cached row'
Assert-Equal 500 $localFirst.Data.Rows[499].Id 'Local page one ends at row 500'
Assert-Equal $false $localFirst.HasPrevious 'Local page one has no Previous'
Assert-Equal $true $localFirst.HasNext 'Local page one has Next when cached rows remain'
Assert-Equal $true $localFirst.IsComplete 'Local page preserves complete state'
Assert-Equal $false $localFirst.IsTruncated 'Local page preserves truncated state'
Assert-True ([object]::ReferenceEquals($cached, $localFirst.CachedData)) 'Local page retains the cache object'

$localSecond = Get-SqlUtilityLocalPage -CachedData $cached -PageNumber 2 -IsComplete $false -IsTruncated $true
Assert-Equal 250 $localSecond.Data.Rows.Count 'Local page two has the remaining rows'
Assert-Equal 501 $localSecond.Data.Rows[0].Id 'Local page two starts at row 501'
Assert-Equal 750 $localSecond.Data.Rows[249].Id 'Local page two ends at the final cached row'
Assert-Equal $true $localSecond.HasPrevious 'Local page two has Previous'
Assert-Equal $false $localSecond.HasNext 'Local final page has no Next'
Assert-Equal $false $localSecond.IsComplete 'Local page preserves incomplete state'
Assert-Equal $true $localSecond.IsTruncated 'Local page preserves truncated state'
Assert-PageResultShape $localSecond 'Local page result'

$emptyCache = New-NumberedTable 0
$emptyLocal = Get-SqlUtilityLocalPage -CachedData $emptyCache -PageNumber 1 -IsComplete $true -IsTruncated $false
Assert-Equal 0 $emptyLocal.Data.Rows.Count 'Empty local page has no rows'
Assert-Equal 1 $emptyLocal.Data.Columns.Count 'Empty local page preserves schema'
Assert-Equal ([int]) $emptyLocal.Data.Columns[0].DataType 'Empty local page preserves column type'
Assert-Throws {
    Get-SqlUtilityLocalPage -CachedData $cached -PageNumber 0 -IsComplete $true -IsTruncated $false
} 'System.ArgumentOutOfRangeException' 'Local paging rejects page zero'

$boundedReader.Index = -1
[void] $boundedReader.Read()
$streamCellError = $null
try {
    $null = Read-SqlUtilityStreamRow -Reader $boundedReader -MaximumTextCharacters 2 -MaximumBinaryBytes 10
}
catch {
    $streamCellError = $_.Exception
}
Assert-Equal 'System.InvalidOperationException' $streamCellError.GetType().FullName `
    'Streaming export rejects oversized text before returning the row'
Assert-True ($streamCellError.Message -match '32,767 characters') `
    'Streaming export reports the existing Excel text-cell limit'

$streamBinaryRow = [object[]]::new(2)
$streamBinaryRow[0] = ''
$streamBinaryRow[1] = [byte[]]::new(11)
$boundedReader.Rows = @(, $streamBinaryRow)
$boundedReader.Index = -1
[void] $boundedReader.Read()
$streamBinaryError = $null
try {
    $null = Read-SqlUtilityStreamRow -Reader $boundedReader -MaximumTextCharacters 10 -MaximumBinaryBytes 10
}
catch {
    $streamBinaryError = $_.Exception
}
Assert-Equal 'System.InvalidOperationException' $streamBinaryError.GetType().FullName `
    'Streaming export rejects oversized binary before returning the row'
Assert-True ($streamBinaryError.Message -match 'hexadecimal text') `
    'Streaming export reports the Excel binary-cell limit'
$boundedReader.Rows = $originalBoundedRows

$script:streamCall = $null
$streamEvents = New-Object System.Collections.Generic.List[string]
$streamSchema = $null
$streamRows = New-Object System.Collections.Generic.List[object]
$onSchema = {
    param($Columns)
    $script:streamSchema = @($Columns)
    [void] $streamEvents.Add('schema')
}
$onRow = {
    param($Values)
    [void] $streamRows.Add($Values)
    [void] $streamEvents.Add("row:$($Values[0])")
}
$shouldContinue = { return $true }
$streamExecutor = {
    param($ConnectionString, $CommandText, $CommandTimeoutSeconds, $OnSchema, $OnRow, $ShouldContinue)
    $script:streamCall = [pscustomobject]@{
        ConnectionString = $ConnectionString
        CommandText = $CommandText
        CommandTimeoutSeconds = $CommandTimeoutSeconds
    }
    $schema = @(
        [pscustomobject]@{ Name = 'Id'; DataType = [int]; Ordinal = 0 },
        [pscustomobject]@{ Name = 'Note'; DataType = [string]; Ordinal = 1 }
    )
    & $OnSchema $schema
    if (& $ShouldContinue) { & $OnRow ([object[]]@(1, [DBNull]::Value)) }
    if (& $ShouldContinue) { & $OnRow ([object[]]@(2, 'second')) }
}

$streamResult = Invoke-SqlUtilityOrderedRowStream -Server ' s ' -Database ' d ' `
    -Sql 'SELECT Id, Note FROM dbo.Items ORDER BY Id' -CommandTimeoutSeconds 456 `
    -OnSchema $onSchema -OnRow $onRow -ShouldContinue $shouldContinue -StreamExecutor $streamExecutor
Assert-Equal $null $streamResult 'Ordered streaming does not expose database resources'
Assert-Equal 'schema,row:1,row:2' ($streamEvents -join ',') 'Ordered streaming invokes schema once before rows'
Assert-Equal 2 $script:streamSchema.Count 'Ordered streaming forwards complete schema metadata'
Assert-Equal 'Id' $script:streamSchema[0].Name 'Ordered streaming schema exposes column names'
Assert-Equal ([int]) $script:streamSchema[0].DataType 'Ordered streaming schema exposes column types'
Assert-Equal 0 $script:streamSchema[0].Ordinal 'Ordered streaming schema exposes ordinals'
Assert-True ([DBNull]::Value.Equals($streamRows[0][1])) 'Ordered streaming preserves SQL null as DBNull'
Assert-Equal 'SELECT Id, Note FROM dbo.Items ORDER BY Id' $script:streamCall.CommandText 'Ordered streaming uses normalized unpaged SQL'
Assert-Equal 456 $script:streamCall.CommandTimeoutSeconds 'Ordered streaming forwards the command timeout'
$streamConnectionBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($script:streamCall.ConnectionString)
Assert-Equal 's' $streamConnectionBuilder.DataSource 'Ordered streaming trims the server in its connection string'
Assert-Equal 'd' $streamConnectionBuilder.InitialCatalog 'Ordered streaming trims the database in its connection string'

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
$countBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($script:countCall.ConnectionString)
Assert-Equal 'server' $countBuilder.DataSource 'Exact count trims and forwards server'
Assert-Equal 'db' $countBuilder.InitialCatalog 'Exact count trims and forwards database'
Assert-True $countBuilder.IntegratedSecurity 'Exact count uses integrated authentication'
Assert-Equal 'SELECT COUNT_BIG(*) FROM (...) AS q ([c1]);' $script:countCall.CommandText `
    'Exact count forwards only policy-generated SQL'
Assert-Equal 321 $script:countCall.CommandTimeoutSeconds 'Exact count forwards query timeout'

foreach ($invalidScalar in @($null, [DBNull]::Value, 'not a number', -1)) {
    $script:countExecutorCalls = 0
    Assert-Throws {
        Invoke-SqlUtilityExactCount -Server 'server' -Database 'db' -CountSql 'SELECT COUNT_BIG(*) FROM q;' `
            -CommandTimeoutSeconds 5 -Executor {
                param($ConnectionString, $CommandText, $CommandTimeoutSeconds)
                $script:countExecutorCalls++
                return $invalidScalar
            }
    } 'System.Data.DataException' 'Exact count rejects an invalid scalar result'
    Assert-Equal 1 $script:countExecutorCalls 'Exact count does not retry or fall back after an invalid scalar result'
}

$executorFailure = [System.InvalidOperationException]::new('count executor failed')
$script:countExecutorCalls = 0
Assert-Throws {
    Invoke-SqlUtilityExactCount -Server 'server' -Database 'db' -CountSql 'SELECT COUNT_BIG(*) FROM q;' `
        -CommandTimeoutSeconds 5 -Executor {
            $script:countExecutorCalls++
            throw $executorFailure
        }
} 'System.InvalidOperationException' 'Exact count propagates executor exceptions'
Assert-Equal 1 $script:countExecutorCalls 'Exact count calls a failing executor only once'

Complete-TestFile 'All database tests passed.'
