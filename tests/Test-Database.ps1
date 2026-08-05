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

$connectionString = New-SqlUtilityConnectionString -Server '  server.example  ' -Database '  UtilityDb  '
$connectionBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($connectionString)
Assert-Equal 'server.example' $connectionBuilder.DataSource 'Connection string trims the server'
Assert-Equal 'UtilityDb' $connectionBuilder.InitialCatalog 'Connection string trims the database'
Assert-True $connectionBuilder.IntegratedSecurity 'Connection string uses integrated security'
Assert-Equal '' $connectionBuilder.UserID 'Connection string has no user name'
Assert-Equal '' $connectionBuilder.Password 'Connection string has no password'
Assert-Equal 'SQL Utility' $connectionBuilder.ApplicationName 'Connection string has the application name'
Assert-Equal 10 $connectionBuilder.ConnectTimeout 'Connection string has a ten-second connect timeout'

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
        param($ConnectionString, $CommandText, $Parameters, $CommandTimeoutSeconds, $MaximumRows)
        $script:orderedCall = [pscustomobject]@{
            ConnectionString = $ConnectionString
            CommandText = $CommandText
            Parameters = $Parameters
            CommandTimeoutSeconds = $CommandTimeoutSeconds
            MaximumRows = $MaximumRows
        }
        return (, (New-NumberedTable $orderedCount))
    }

    $ordered = Invoke-SqlUtilityOrderedPage -Server 's' -Database 'd' `
        -Sql 'SELECT Id FROM dbo.Items ORDER BY Id' -PageNumber 2 `
        -CommandTimeoutSeconds 120 -Executor $orderedExecutor

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
Assert-Equal 500 $script:orderedCall.Parameters.Offset 'Ordered page two sends the correct offset'
Assert-Equal 501 $script:orderedCall.Parameters.FetchCount 'Ordered paging probes one sentinel row'
Assert-Equal 120 $script:orderedCall.CommandTimeoutSeconds 'Ordered paging forwards the command timeout'
Assert-Equal 501 $script:orderedCall.MaximumRows 'Ordered paging bounds the executor read'

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
        param($ConnectionString, $CommandText, $Parameters, $CommandTimeoutSeconds, $MaximumRows)
        $script:unorderedCall = [pscustomobject]@{
            ConnectionString = $ConnectionString
            CommandText = $CommandText
            Parameters = $Parameters
            CommandTimeoutSeconds = $CommandTimeoutSeconds
            MaximumRows = $MaximumRows
        }
        return (, $sourceTable)
    }

    $unordered = Invoke-SqlUtilityUnorderedQuery -Server 's' -Database 'd' `
        -Sql 'SELECT Id FROM dbo.Items' -RowLimit 1000 -CommandTimeoutSeconds 321 `
        -Executor $unorderedExecutor

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

Complete-TestFile 'All database tests passed.'
