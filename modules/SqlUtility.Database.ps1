function New-SqlUtilityConnectionString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database
    )

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder.psbase.DataSource = $Server.Trim()
    $builder.psbase.InitialCatalog = $Database.Trim()
    $builder.psbase.IntegratedSecurity = $true
    $builder.psbase.ApplicationName = 'SQL Utility'
    $builder.psbase.ConnectTimeout = 10
    return $builder.ConnectionString
}

function New-SqlUtilityResultTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]] $Columns
    )

    $table = [System.Data.DataTable]::new('Results')
    $usedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $canonicalNames = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($column in @($Columns | Sort-Object -Property Ordinal)) {
        $baseName = [string] $column.Name
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            $baseName = 'Column ' + ([int] $column.Ordinal + 1)
        }

        if ($canonicalNames.ContainsKey($baseName)) {
            $baseName = $canonicalNames[$baseName]
        }
        else {
            $canonicalNames.Add($baseName, $baseName)
        }

        $columnName = $baseName
        $suffix = 2
        while (-not $usedNames.Add($columnName)) {
            $columnName = "$baseName ($suffix)"
            $suffix++
        }

        [void] $table.Columns.Add($columnName, [type] $column.DataType)
    }

    return (, $table)
}

function New-SqlUtilityPageResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.DataTable] $Data,
        [AllowNull()][System.Data.DataTable] $CachedData,
        [Parameter(Mandatory = $true)][int] $PageNumber,
        [Parameter(Mandatory = $true)][bool] $HasPrevious,
        [Parameter(Mandatory = $true)][bool] $HasNext,
        [Parameter(Mandatory = $true)][bool] $IsComplete,
        [Parameter(Mandatory = $true)][bool] $IsTruncated
    )

    return [pscustomobject][ordered]@{
        Data = $Data
        CachedData = $CachedData
        PageNumber = $PageNumber
        DisplayedRowCount = $Data.Rows.Count
        HasPrevious = $HasPrevious
        HasNext = $HasNext
        IsComplete = $IsComplete
        IsTruncated = $IsTruncated
    }
}

function Copy-SqlUtilityDataRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.DataTable] $Source,
        [Parameter(Mandatory = $true)][int] $StartIndex,
        [Parameter(Mandatory = $true)][int] $Count
    )

    $copy = $Source.Clone()
    $endExclusive = [Math]::Min($Source.Rows.Count, $StartIndex + $Count)
    for ($index = $StartIndex; $index -lt $endExclusive; $index++) {
        $copy.ImportRow($Source.Rows[$index])
    }
    return (, $copy)
}

function Invoke-SqlUtilityTableExecutor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ConnectionString,
        [Parameter(Mandatory = $true)][string] $CommandText,
        [Parameter(Mandatory = $true)][hashtable] $Parameters,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [Parameter(Mandatory = $true)][int] $MaximumRows
    )

    $connection = [System.Data.SqlClient.SqlConnection]::new($ConnectionString)
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        try {
            $command.CommandText = $CommandText
            $command.CommandTimeout = $CommandTimeoutSeconds

            foreach ($parameterName in @('Offset', 'FetchCount')) {
                if ($Parameters.ContainsKey($parameterName)) {
                    $parameter = $command.Parameters.Add("@$parameterName", [System.Data.SqlDbType]::Int)
                    $parameter.Value = [int] $Parameters[$parameterName]
                }
            }

            $reader = $null
            try {
                $reader = $command.ExecuteReader()
                $schemaTable = $reader.GetSchemaTable()
                $columns = @(
                    for ($ordinal = 0; $ordinal -lt $reader.FieldCount; $ordinal++) {
                        $columnName = $reader.GetName($ordinal)
                        $columnType = $reader.GetFieldType($ordinal)
                        if ($null -ne $schemaTable -and $ordinal -lt $schemaTable.Rows.Count) {
                            if ($null -ne $schemaTable.Rows[$ordinal].ColumnName) {
                                $columnName = [string] $schemaTable.Rows[$ordinal].ColumnName
                            }
                            if ($schemaTable.Rows[$ordinal].DataType -is [type]) {
                                $columnType = [type] $schemaTable.Rows[$ordinal].DataType
                            }
                        }

                        [pscustomobject]@{
                            Name = $columnName
                            DataType = $columnType
                            Ordinal = $ordinal
                        }
                    }
                )
                $table = New-SqlUtilityResultTable -Columns $columns

                $rowCount = 0
                while ($rowCount -lt $MaximumRows -and $reader.Read()) {
                    $row = $table.NewRow()
                    for ($ordinal = 0; $ordinal -lt $reader.FieldCount; $ordinal++) {
                        $row[$ordinal] = $reader.GetValue($ordinal)
                    }
                    [void] $table.Rows.Add($row)
                    $rowCount++
                }

                if ($rowCount -ge $MaximumRows) {
                    try { $command.Cancel() } catch { }
                }
                return (, $table)
            }
            finally {
                if ($null -ne $reader) {
                    $reader.Dispose()
                }
            }
        }
        finally {
            $command.Dispose()
        }
    }
    finally {
        $connection.Dispose()
    }
}

function Invoke-SqlUtilityConnectionTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [scriptblock] $Executor
    )

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    if ($null -ne $Executor) {
        $null = & $Executor $connectionString 'SELECT 1;' @{} 10 1
        return
    }

    $connection = [System.Data.SqlClient.SqlConnection]::new($connectionString)
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        try {
            $command.CommandText = 'SELECT 1;'
            $command.CommandTimeout = 10
            $null = $command.ExecuteScalar()
        }
        finally {
            $command.Dispose()
        }
    }
    finally {
        $connection.Dispose()
    }
}

function Invoke-SqlUtilityOrderedPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][string] $Sql,
        [Parameter(Mandatory = $true)][int] $PageNumber,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    if ($PageNumber -lt 1) {
        throw [System.ArgumentOutOfRangeException]::new('PageNumber', 'Page number must be at least 1.')
    }

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    $commandText = $Sql + "`r`nOFFSET @Offset ROWS FETCH NEXT @FetchCount ROWS ONLY"
    $parameters = @{ Offset = (($PageNumber - 1) * 500); FetchCount = 501 }
    if ($null -eq $Executor) {
        $data = Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $commandText `
            -Parameters $parameters -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows 501
    }
    else {
        $data = & $Executor $connectionString $commandText $parameters $CommandTimeoutSeconds 501
    }

    $hasNext = $data.Rows.Count -gt 500
    $displayData = Copy-SqlUtilityDataRows -Source $data -StartIndex 0 -Count 500
    return New-SqlUtilityPageResult -Data $displayData -CachedData $null -PageNumber $PageNumber `
        -HasPrevious ($PageNumber -gt 1) -HasNext $hasNext -IsComplete $false -IsTruncated $false
}

function Invoke-SqlUtilityUnorderedQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][string] $Sql,
        [Parameter(Mandatory = $true)][int] $RowLimit,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    $maximumRows = $RowLimit + 1
    if ($null -eq $Executor) {
        $data = Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $Sql `
            -Parameters @{} -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows $maximumRows
    }
    else {
        $data = & $Executor $connectionString $Sql @{} $CommandTimeoutSeconds $maximumRows
    }

    $isTruncated = $data.Rows.Count -gt $RowLimit
    $cachedData = Copy-SqlUtilityDataRows -Source $data -StartIndex 0 -Count $RowLimit
    return Get-SqlUtilityLocalPage -CachedData $cachedData -PageNumber 1 `
        -IsComplete (-not $isTruncated) -IsTruncated $isTruncated
}

function Get-SqlUtilityLocalPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.DataTable] $CachedData,
        [Parameter(Mandatory = $true)][int] $PageNumber,
        [Parameter(Mandatory = $true)][bool] $IsComplete,
        [Parameter(Mandatory = $true)][bool] $IsTruncated
    )

    if ($PageNumber -lt 1) {
        throw [System.ArgumentOutOfRangeException]::new('PageNumber', 'Page number must be at least 1.')
    }

    $startIndex = ($PageNumber - 1) * 500
    $pageData = Copy-SqlUtilityDataRows -Source $CachedData -StartIndex $startIndex -Count 500
    return New-SqlUtilityPageResult -Data $pageData -CachedData $CachedData -PageNumber $PageNumber `
        -HasPrevious ($PageNumber -gt 1) -HasNext (($startIndex + 500) -lt $CachedData.Rows.Count) `
        -IsComplete $IsComplete -IsTruncated $IsTruncated
}

function Invoke-SqlUtilityStreamExecutor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ConnectionString,
        [Parameter(Mandatory = $true)][string] $CommandText,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [Parameter(Mandatory = $true)][scriptblock] $OnSchema,
        [Parameter(Mandatory = $true)][scriptblock] $OnRow,
        [Parameter(Mandatory = $true)][scriptblock] $ShouldContinue
    )

    $connection = [System.Data.SqlClient.SqlConnection]::new($ConnectionString)
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        try {
            $command.CommandText = $CommandText
            $command.CommandTimeout = $CommandTimeoutSeconds
            $reader = $null
            try {
                $reader = $command.ExecuteReader()
                $schema = @(
                    for ($ordinal = 0; $ordinal -lt $reader.FieldCount; $ordinal++) {
                        [pscustomobject][ordered]@{
                            Name = $reader.GetName($ordinal)
                            DataType = $reader.GetFieldType($ordinal)
                            Ordinal = $ordinal
                        }
                    }
                )
                $null = & $OnSchema $schema

                while ($reader.Read()) {
                    if (-not (& $ShouldContinue)) {
                        try { $command.Cancel() } catch { }
                        break
                    }

                    $values = [object[]]::new($reader.FieldCount)
                    [void] $reader.GetValues($values)
                    $null = & $OnRow $values
                }
            }
            finally {
                if ($null -ne $reader) {
                    $reader.Dispose()
                }
            }
        }
        finally {
            $command.Dispose()
        }
    }
    finally {
        $connection.Dispose()
    }
}

function Invoke-SqlUtilityOrderedRowStream {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][string] $Sql,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [Parameter(Mandatory = $true)][scriptblock] $OnSchema,
        [Parameter(Mandatory = $true)][scriptblock] $OnRow,
        [Parameter(Mandatory = $true)][scriptblock] $ShouldContinue,
        [scriptblock] $StreamExecutor
    )

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    if ($null -eq $StreamExecutor) {
        $null = Invoke-SqlUtilityStreamExecutor -ConnectionString $connectionString -CommandText $Sql `
            -CommandTimeoutSeconds $CommandTimeoutSeconds -OnSchema $OnSchema -OnRow $OnRow `
            -ShouldContinue $ShouldContinue
    }
    else {
        $null = & $StreamExecutor $connectionString $Sql $CommandTimeoutSeconds $OnSchema $OnRow $ShouldContinue
    }
}
