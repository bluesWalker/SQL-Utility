$script:SqlUtilityDataExplorerOutputOnlyTypes = @(
    'binary', 'varbinary', 'rowversion', 'timestamp', 'xml', 'geography', 'geometry',
    'hierarchyid', 'sql_variant', 'text', 'ntext', 'image'
)
$script:SqlUtilityDataExplorerStringTypes = @('char', 'varchar', 'nchar', 'nvarchar')
$script:SqlUtilityDataExplorerNumericTypes = @('tinyint', 'smallint', 'int', 'bigint', 'decimal', 'numeric', 'smallmoney', 'money', 'real', 'float')
$script:SqlUtilityDataExplorerTemporalTypes = @('date', 'smalldatetime', 'datetime', 'datetime2', 'time', 'datetimeoffset')

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

function Get-SqlUtilityDataExplorerColumnType([object] $Column) {
    if ($null -eq $Column -or [string]::IsNullOrWhiteSpace([string] $Column.SqlTypeName)) {
        throw [System.ArgumentException]::new('Column type metadata is required.')
    }
    return ([string] $Column.SqlTypeName).ToLowerInvariant()
}

function Test-SqlUtilityDataExplorerOutputOnlyColumn([object] $Column) {
    return [bool] $Column.IsUserDefined -or ((Get-SqlUtilityDataExplorerColumnType $Column) -in $script:SqlUtilityDataExplorerOutputOnlyTypes)
}

function Get-SqlUtilityDataExplorerOperators {
    param([Parameter(Mandatory = $true)][object] $Column)
    if (Test-SqlUtilityDataExplorerOutputOnlyColumn $Column) { return @() }
    $type = Get-SqlUtilityDataExplorerColumnType $Column
    $operators = @()
    if ($type -in $script:SqlUtilityDataExplorerStringTypes) {
        $operators += @(
            [pscustomobject][ordered]@{ Key='Equals'; Label='Equals'; RequiresValue=$true },
            [pscustomobject][ordered]@{ Key='NotEquals'; Label='Does not equal'; RequiresValue=$true },
            [pscustomobject][ordered]@{ Key='Contains'; Label='Contains'; RequiresValue=$true },
            [pscustomobject][ordered]@{ Key='StartsWith'; Label='Starts with'; RequiresValue=$true }
        )
    }
    elseif (($type -in $script:SqlUtilityDataExplorerNumericTypes) -or ($type -in $script:SqlUtilityDataExplorerTemporalTypes) -or $type -in @('bit', 'uniqueidentifier')) {
        $operators += @(
            [pscustomobject][ordered]@{ Key='Equals'; Label='Equals'; RequiresValue=$true },
            [pscustomobject][ordered]@{ Key='NotEquals'; Label='Does not equal'; RequiresValue=$true }
        )
        if ($type -notin @('bit', 'uniqueidentifier')) {
            $operators += @(
                [pscustomobject][ordered]@{ Key='GreaterThan'; Label='Greater than'; RequiresValue=$true },
                [pscustomobject][ordered]@{ Key='GreaterThanOrEqual'; Label='Greater than or equal to'; RequiresValue=$true },
                [pscustomobject][ordered]@{ Key='LessThan'; Label='Less than'; RequiresValue=$true },
                [pscustomobject][ordered]@{ Key='LessThanOrEqual'; Label='Less than or equal to'; RequiresValue=$true }
            )
        }
    }
    else { return @() }
    if ([bool] $Column.IsNullable) {
        $operators += @(
            [pscustomobject][ordered]@{ Key='IsNull'; Label='Is null'; RequiresValue=$false },
            [pscustomobject][ordered]@{ Key='IsNotNull'; Label='Is not null'; RequiresValue=$false }
        )
    }
    return $operators
}

function Get-SqlUtilityDataExplorerColumnDisplayText {
    param([Parameter(Mandatory = $true)][object] $Column)
    if ([string]::IsNullOrWhiteSpace([string] $Column.Name)) { throw [System.ArgumentException]::new('Column name is required.') }
    $type = Get-SqlUtilityDataExplorerColumnType $Column
    $displayType = $type
    if ($type -in @('char', 'varchar', 'binary', 'varbinary')) {
        $displayType += if ([int] $Column.MaxLength -eq -1) { '(max)' } else { "($([int] $Column.MaxLength))" }
    }
    elseif ($type -in @('nchar', 'nvarchar')) {
        $displayType += if ([int] $Column.MaxLength -eq -1) { '(max)' } else { "($([int] $Column.MaxLength / 2))" }
    }
    elseif ($type -in @('decimal', 'numeric')) { $displayType += "($([int] $Column.Precision),$([int] $Column.Scale))" }
    elseif ($type -in @('time', 'datetime2', 'datetimeoffset')) { $displayType += "($([int] $Column.Scale))" }
    return "$($Column.Name) ($displayType)"
}

function ConvertTo-SqlUtilityDataExplorerTypedValue {
    param([Parameter(Mandatory = $true)][object] $Column, [Parameter(Mandatory = $true)][string] $ValueText)
    $type = Get-SqlUtilityDataExplorerColumnType $Column
    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    try {
        switch ($type) {
            { $_ -in $script:SqlUtilityDataExplorerStringTypes } { return [string] $ValueText }
            'tinyint' { return [byte]::Parse($ValueText, [System.Globalization.NumberStyles]::Integer, $culture) }
            'smallint' { return [int16]::Parse($ValueText, [System.Globalization.NumberStyles]::Integer, $culture) }
            'int' { return [int]::Parse($ValueText, [System.Globalization.NumberStyles]::Integer, $culture) }
            'bigint' { return [int64]::Parse($ValueText, [System.Globalization.NumberStyles]::Integer, $culture) }
            { $_ -in @('decimal', 'numeric', 'smallmoney', 'money') } { return [decimal]::Parse($ValueText, [System.Globalization.NumberStyles]::Number, $culture) }
            'real' { $value = [single]::Parse($ValueText, [System.Globalization.NumberStyles]::Float, $culture); if ([single]::IsNaN($value) -or [single]::IsInfinity($value)) { throw 'Non-finite value.' }; return $value }
            'float' { $value = [double]::Parse($ValueText, [System.Globalization.NumberStyles]::Float, $culture); if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { throw 'Non-finite value.' }; return $value }
            { $_ -in @('date', 'smalldatetime', 'datetime', 'datetime2') } {
                $value = [datetime]::MinValue
                $isoFormats = [string[]] @(
                    'yyyy-MM-dd',
                    "yyyy-MM-dd'T'HH:mm:ss",
                    "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
                    "yyyy-MM-dd'T'HH:mm:ssK",
                    "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"
                )
                if (-not [datetime]::TryParseExact($ValueText, $isoFormats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref] $value)) {
                    $value = [datetime]::Parse($ValueText, $culture)
                }
                return $value
            }
            'time' {
                $value = [timespan]::Zero
                if (-not [timespan]::TryParseExact($ValueText, 'c', [System.Globalization.CultureInfo]::InvariantCulture, [ref] $value)) {
                    $value = [timespan]::Parse($ValueText, $culture)
                }
                return $value
            }
            'datetimeoffset' {
                $value = [datetimeoffset]::MinValue
                $isoFormats = [string[]] @(
                    "yyyy-MM-dd'T'HH:mm:sszzz",
                    "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz",
                    "yyyy-MM-dd'T'HH:mm:ss'Z'",
                    "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'"
                )
                if (-not [datetimeoffset]::TryParseExact($ValueText, $isoFormats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref] $value)) {
                    $value = [datetimeoffset]::Parse($ValueText, $culture)
                }
                return $value
            }
            'bit' { return [System.Convert]::ToBoolean($ValueText, $culture) }
            'uniqueidentifier' { return [guid]::Parse($ValueText) }
            default { throw "Unsupported filter type: $type" }
        }
    }
    catch { throw [System.ArgumentException]::new("Invalid $type value.", $_.Exception) }
}

function ConvertTo-SqlUtilityDataExplorerMetadataValue {
    param(
        [Parameter(Mandatory = $true)][object] $Value,
        [Parameter(Mandatory = $true)][object] $Column
    )

    $type = Get-SqlUtilityDataExplorerColumnType $Column
    if ($type -in @('decimal', 'numeric')) {
        $precision = [int] $Column.Precision
        $scale = [int] $Column.Scale
        if ($precision -lt 1 -or $precision -gt 38 -or $scale -lt 0 -or $scale -gt $precision) {
            throw [System.ArgumentException]::new('Decimal column precision and scale metadata is invalid.')
        }

        $normalized = [decimal]::Round([decimal] $Value, $scale, [System.MidpointRounding]::ToEven)
        if ($normalized -ne [decimal] $Value) {
            throw [System.ArgumentException]::new("Invalid $type value: fractional digits exceed scale $scale.")
        }

        $integralDigits = $precision - $scale
        if ($integralDigits -lt 29) {
            $exclusiveLimit = [decimal] 1
            for ($index = 0; $index -lt $integralDigits; $index++) {
                $exclusiveLimit *= [decimal] 10
            }
            if ([Math]::Abs([decimal] $normalized) -ge $exclusiveLimit) {
                throw [System.ArgumentException]::new("Invalid $type value: digits exceed precision $precision.")
            }
        }
        return $normalized
    }

    if ($type -in @('time', 'datetime2', 'datetimeoffset')) {
        $scale = [int] $Column.Scale
        if ($scale -lt 0 -or $scale -gt 7) {
            throw [System.ArgumentException]::new('Temporal column scale metadata is invalid.')
        }
        $tickQuantum = [long] [Math]::Pow(10, 7 - $scale)
        if (([long] $Value.Ticks % $tickQuantum) -ne 0) {
            throw [System.ArgumentException]::new("Invalid $type value: fractional seconds exceed scale $scale.")
        }
    }
    return $Value
}

function Get-SqlUtilityDataExplorerDbType([string] $Type) {
    switch ($Type) {
        'char' { return [System.Data.SqlDbType]::Char }; 'varchar' { return [System.Data.SqlDbType]::VarChar }
        'nchar' { return [System.Data.SqlDbType]::NChar }; 'nvarchar' { return [System.Data.SqlDbType]::NVarChar }
        'tinyint' { return [System.Data.SqlDbType]::TinyInt }; 'smallint' { return [System.Data.SqlDbType]::SmallInt }
        'int' { return [System.Data.SqlDbType]::Int }; 'bigint' { return [System.Data.SqlDbType]::BigInt }
        'decimal' { return [System.Data.SqlDbType]::Decimal }; 'numeric' { return [System.Data.SqlDbType]::Decimal }
        'smallmoney' { return [System.Data.SqlDbType]::SmallMoney }; 'money' { return [System.Data.SqlDbType]::Money }
        'real' { return [System.Data.SqlDbType]::Real }; 'float' { return [System.Data.SqlDbType]::Float }
        'date' { return [System.Data.SqlDbType]::Date }; 'smalldatetime' { return [System.Data.SqlDbType]::SmallDateTime }
        'datetime' { return [System.Data.SqlDbType]::DateTime }; 'datetime2' { return [System.Data.SqlDbType]::DateTime2 }
        'time' { return [System.Data.SqlDbType]::Time }; 'datetimeoffset' { return [System.Data.SqlDbType]::DateTimeOffset }
        'bit' { return [System.Data.SqlDbType]::Bit }; 'uniqueidentifier' { return [System.Data.SqlDbType]::UniqueIdentifier }
        default { throw [System.ArgumentException]::new("Unsupported filter type: $Type") }
    }
}

function ConvertTo-SqlUtilityDataExplorerLiteral([object] $Value, [object] $Column) {
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $type = Get-SqlUtilityDataExplorerColumnType $Column
    if ($Type -in @('char', 'varchar', 'nchar', 'nvarchar')) { $prefix = if ($Type -in @('nchar', 'nvarchar')) { 'N' } else { '' }; return $prefix + "'" + ([string] $Value).Replace("'", "''") + "'" }
    if ($Type -in @('tinyint','smallint','int','bigint','decimal','numeric','smallmoney','money','real','float')) { return [System.Convert]::ToString($Value, $invariant) }
    if ($Type -eq 'date') { return "'$($Value.ToString('yyyy-MM-dd', $invariant))'" }
    if ($type -eq 'datetime2') {
        $scale = [int] $Column.Scale
        $format = "yyyy-MM-dd'T'HH:mm:ss"
        if ($scale -gt 0) { $format += '.' + ('f' * $scale) }
        return "'$($Value.ToString($format, $invariant))'"
    }
    if ($type -in @('smalldatetime','datetime')) { return "'$($Value.ToString("yyyy-MM-dd'T'HH:mm:ss.fff", $invariant))'" }
    if ($Type -eq 'time') { return "'$($Value.ToString('hh\:mm\:ss\.fffffff', $invariant))'" }
    if ($Type -eq 'datetimeoffset') { return "'$($Value.ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz', $invariant))'" }
    if ($Type -eq 'bit') { if ($Value) { return '1' }; return '0' }
    if ($Type -eq 'uniqueidentifier') { return "'$($Value.ToString('D'))'" }
    throw [System.ArgumentException]::new("Unsupported literal type: $Type")
}

function New-SqlUtilityDataExplorerQuery {
    param(
        [Parameter(Mandatory = $true)][object] $Table, [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]] $Columns,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]] $SelectedColumnNames, [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]] $Filters,
        [Parameter(Mandatory = $true)][object] $PreviewRowLimit
    )
    if ($null -eq $Table -or [string]::IsNullOrWhiteSpace([string] $Table.SchemaName) -or [string]::IsNullOrWhiteSpace([string] $Table.TableName)) { throw [System.ArgumentException]::new('Valid table metadata is required.') }
    if ($PreviewRowLimit -isnot [sbyte] -and $PreviewRowLimit -isnot [byte] -and $PreviewRowLimit -isnot [int16] -and $PreviewRowLimit -isnot [uint16] -and $PreviewRowLimit -isnot [int] -and $PreviewRowLimit -isnot [uint32] -and $PreviewRowLimit -isnot [int64] -and $PreviewRowLimit -isnot [uint64]) { throw [System.ArgumentException]::new('Preview row limit must be an integral CLR value.') }
    $previewLimitValue = [decimal] $PreviewRowLimit
    if ($previewLimitValue -lt 10 -or $previewLimitValue -gt 500) { throw [System.ArgumentException]::new('Preview row limit must be from 10 through 500.') }
    $previewLimit = [int] $previewLimitValue
    if ($SelectedColumnNames.Count -eq 0) { throw [System.ArgumentException]::new('Select at least one column.') }
    $lookup = @{}; foreach ($column in $Columns) { if ($null -eq $column -or [string]::IsNullOrWhiteSpace([string] $column.Name) -or $lookup.ContainsKey($column.Name.ToLowerInvariant())) { throw [System.ArgumentException]::new('Column metadata names must be unique.') }; $lookup[$column.Name.ToLowerInvariant()] = $column }
    $selected = @(); $seen = @{}; foreach ($name in $SelectedColumnNames) { $key = $name.ToLowerInvariant(); if (-not $lookup.ContainsKey($key) -or $seen.ContainsKey($key)) { throw [System.ArgumentException]::new('Selected columns must be known and unique.') }; $seen[$key] = $true; $selected += $lookup[$key] }
    $selected = @($selected | Sort-Object Ordinal)
    $parameters = @([pscustomobject][ordered]@{ Name='PreviewLimit'; SqlDbType=[System.Data.SqlDbType]::Int; Size=0; Precision=[byte] 0; Scale=[byte] 0; Value=$previewLimit })
    $previewPredicates = @(); $editorPredicates = @(); $filterIndex = 0
    foreach ($filter in $Filters) {
        if ($null -eq $filter -or [string]::IsNullOrWhiteSpace([string] $filter.ColumnName) -or [string]::IsNullOrWhiteSpace([string] $filter.Operator)) { throw [System.ArgumentException]::new('Filter metadata is required.') }
        $key = $filter.ColumnName.ToLowerInvariant(); if (-not $lookup.ContainsKey($key)) { throw [System.ArgumentException]::new('Filter column must be known.') }; $column = $lookup[$key]
        $operator = @((Get-SqlUtilityDataExplorerOperators -Column $column) | Where-Object { $_.Key -eq $filter.Operator })
        if ($operator.Count -ne 1) { throw [System.ArgumentException]::new('Filter operator is not valid for this column.') }
        $identifier = ConvertTo-SqlUtilityBracketIdentifier $column.Name
        if ($filter.Operator -eq 'IsNull' -or $filter.Operator -eq 'IsNotNull') { $word = if ($filter.Operator -eq 'IsNull') { 'IS NULL' } else { 'IS NOT NULL' }; $previewPredicates += "$identifier $word"; $editorPredicates += "$identifier $word"; continue }
        $filterIndex++; $name = "Filter$filterIndex"; $type = Get-SqlUtilityDataExplorerColumnType $column; $value = ConvertTo-SqlUtilityDataExplorerTypedValue -Column $column -ValueText ([string] $filter.ValueText)
        $value = ConvertTo-SqlUtilityDataExplorerMetadataValue -Value $value -Column $column
        $sqlOperator = @{ Equals='='; NotEquals='<>'; GreaterThan='>'; GreaterThanOrEqual='>='; LessThan='<'; LessThanOrEqual='<=' }[$filter.Operator]
        if ($filter.Operator -in @('Contains','StartsWith')) { $value = ConvertTo-SqlUtilityLikePattern -Value ([string] $value) -Operator $filter.Operator; $sqlOperator = 'LIKE' }
        $size = 0; if ($type -in @('char','varchar','nchar','nvarchar')) { $size = [int] $column.MaxLength; if ($type -in @('nchar','nvarchar') -and $size -gt 0) { $size = [int] ($size / 2) } }
        $precision = [byte] 0; $scale = [byte] 0; if ($type -in @('decimal','numeric')) { $precision = [byte] $column.Precision; $scale = [byte] $column.Scale } elseif ($type -in @('time','datetime2','datetimeoffset')) { $scale = [byte] $column.Scale }
        $parameters += [pscustomobject][ordered]@{ Name=$name; SqlDbType=(Get-SqlUtilityDataExplorerDbType $type); Size=$size; Precision=$precision; Scale=$scale; Value=$value }
        $previewPredicates += "$identifier $sqlOperator @$name"; $editorPredicates += "$identifier $sqlOperator $(ConvertTo-SqlUtilityDataExplorerLiteral -Value $value -Column $column)"
    }
    $source = "$(ConvertTo-SqlUtilityBracketIdentifier $Table.SchemaName).$(ConvertTo-SqlUtilityBracketIdentifier $Table.TableName)"
    $output = (@($selected | ForEach-Object { ConvertTo-SqlUtilityBracketIdentifier $_.Name }) -join ', ')
    $where = if ($previewPredicates.Count -gt 0) { ' WHERE ' + ($previewPredicates -join ' AND ') } else { '' }
    $editorLines = @("SELECT $output", "FROM $source")
    if ($editorPredicates.Count -gt 0) {
        $editorLines += "WHERE $($editorPredicates[0])"
        for ($index = 1; $index -lt $editorPredicates.Count; $index++) {
            $editorLines += "  AND $($editorPredicates[$index])"
        }
    }
    $editorSql = ($editorLines -join "`r`n") + ';'
    return [pscustomobject][ordered]@{ PreviewSql="SELECT TOP (@PreviewLimit) $output FROM $source$where;"; PreviewParameters=$parameters; EditorSql=$editorSql }
}
