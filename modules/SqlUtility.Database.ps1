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

function Invoke-SqlUtilityTableExecutor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ConnectionString,
        [Parameter(Mandatory = $true)][string] $CommandText,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]] $ParameterDescriptors,
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

            Add-SqlUtilityCommandParameters -Command $command -ParameterDescriptors $ParameterDescriptors

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

function Invoke-SqlUtilityExactCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][string] $CountSql,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    if ([string]::IsNullOrWhiteSpace($CountSql)) {
        throw [System.ArgumentException]::new('Count SQL must not be blank.', 'CountSql')
    }
    if ($CommandTimeoutSeconds -lt 1) {
        throw [System.ArgumentOutOfRangeException]::new('CommandTimeoutSeconds', 'Command timeout must be at least 1 second.')
    }

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    if ($null -eq $Executor) {
        $value = Invoke-SqlUtilityScalarExecutor -ConnectionString $connectionString -CommandText $CountSql `
            -CommandTimeoutSeconds $CommandTimeoutSeconds
    }
    else {
        $value = & $Executor $connectionString $CountSql $CommandTimeoutSeconds
    }

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
        $null = & $Executor $connectionString 'SELECT 1;' ([object[]] @()) 10 1
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

function Format-SqlUtilityDatabaseIdentifier([string] $Name) {
    return '[' + $Name.Replace(']', ']]') + ']'
}

function Get-SqlUtilityPhysicalTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    $commandText = @'
SELECT t.object_id AS ObjectId, s.name AS SchemaName, t.name AS TableName
FROM sys.tables AS t
INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.is_ms_shipped = 0
ORDER BY s.name, t.name;
'@
    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    $parameters = [object[]] @()
    if ($null -eq $Executor) {
        $data = Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $commandText `
            -ParameterDescriptors $parameters -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows ([int]::MaxValue)
    }
    else {
        $data = & $Executor $connectionString $commandText $parameters $CommandTimeoutSeconds ([int]::MaxValue)
    }

    return @($data.Rows | ForEach-Object {
        [pscustomobject][ordered]@{
            ObjectId = [int] $_.ObjectId
            SchemaName = [string] $_.SchemaName
            TableName = [string] $_.TableName
            DisplayName = (Format-SqlUtilityDatabaseIdentifier ([string] $_.SchemaName)) + '.' + `
                (Format-SqlUtilityDatabaseIdentifier ([string] $_.TableName))
        }
    })
}

function Get-SqlUtilityTableColumns {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)][int] $TableObjectId,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    $commandText = @'
SELECT t.object_id AS ObjectId,
       s.name AS SchemaName,
       t.name AS TableName,
       c.name AS Name,
       c.column_id AS Ordinal,
       COALESCE(base_type.name, declared_type.name) AS SqlTypeName,
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
LEFT JOIN sys.types AS base_type
    ON base_type.system_type_id = c.system_type_id
   AND base_type.user_type_id = base_type.system_type_id
WHERE t.is_ms_shipped = 0
  AND t.object_id = @TableObjectId
ORDER BY c.column_id;
'@
    $parameters = [object[]] @(
        [pscustomobject]@{ Name = 'TableObjectId'; SqlDbType = [System.Data.SqlDbType]::Int; Size = 0; Precision = 0; Scale = 0; Value = $TableObjectId }
    )
    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    if ($null -eq $Executor) {
        $data = Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $commandText `
            -ParameterDescriptors $parameters -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows ([int]::MaxValue)
    }
    else {
        $data = & $Executor $connectionString $commandText $parameters $CommandTimeoutSeconds ([int]::MaxValue)
    }

    return @($data.Rows | ForEach-Object {
        [pscustomobject][ordered]@{
            Name = [string] $_.Name
            Ordinal = [int] $_.Ordinal
            SqlTypeName = [string] $_.SqlTypeName
            MaxLength = [int] $_.MaxLength
            Precision = [byte] $_.Precision
            Scale = [byte] $_.Scale
            IsNullable = [bool] $_.IsNullable
            IsUserDefined = [bool] $_.IsUserDefined
        }
    })
}

function Invoke-SqlUtilityDataPreview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database,
        [Parameter(Mandatory = $true)] $Query,
        [Parameter(Mandatory = $true)][object] $PreviewRowLimit,
        [Parameter(Mandatory = $true)][int] $CommandTimeoutSeconds,
        [scriptblock] $Executor
    )

    if ($PreviewRowLimit -isnot [sbyte] -and $PreviewRowLimit -isnot [byte] -and `
        $PreviewRowLimit -isnot [int16] -and $PreviewRowLimit -isnot [uint16] -and `
        $PreviewRowLimit -isnot [int] -and $PreviewRowLimit -isnot [uint32] -and `
        $PreviewRowLimit -isnot [int64] -and $PreviewRowLimit -isnot [uint64]) {
        throw [System.ArgumentException]::new('Preview row limit must be an integral CLR value.', 'PreviewRowLimit')
    }
    $previewLimitValue = [decimal] $PreviewRowLimit
    if ($previewLimitValue -lt 10 -or $previewLimitValue -gt 500) {
        throw [System.ArgumentOutOfRangeException]::new('PreviewRowLimit', 'Preview row limit must be from 10 through 500.')
    }

    $connectionString = New-SqlUtilityConnectionString -Server $Server -Database $Database
    if ($null -eq $Executor) {
        return Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $Query.PreviewSql `
            -ParameterDescriptors $Query.PreviewParameters -CommandTimeoutSeconds $CommandTimeoutSeconds `
            -MaximumRows ([int] $previewLimitValue)
    }
    return & $Executor $connectionString $Query.PreviewSql $Query.PreviewParameters $CommandTimeoutSeconds ([int] $previewLimitValue)
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
    $parameters = [object[]] @(
        [pscustomobject]@{ Name = 'Offset'; SqlDbType = [System.Data.SqlDbType]::Int; Size = 0; Precision = 0; Scale = 0; Value = (($PageNumber - 1) * 500) },
        [pscustomobject]@{ Name = 'FetchCount'; SqlDbType = [System.Data.SqlDbType]::Int; Size = 0; Precision = 0; Scale = 0; Value = 501 }
    )
    if ($null -eq $Executor) {
        $data = Invoke-SqlUtilityTableExecutor -ConnectionString $connectionString -CommandText $commandText `
            -ParameterDescriptors $parameters -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows 501
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
            -ParameterDescriptors ([object[]] @()) -CommandTimeoutSeconds $CommandTimeoutSeconds -MaximumRows $maximumRows
    }
    else {
        $data = & $Executor $connectionString $Sql ([object[]] @()) $CommandTimeoutSeconds $maximumRows
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
