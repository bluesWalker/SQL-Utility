function Get-SqlUtilityConfigPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $InputObject,
        [Parameter(Mandatory = $true)][string] $Name
    )

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if ([string]::Equals([string] $key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-Output -NoEnumerate $InputObject[$key]
                return
            }
        }
    }
    else {
        $property = $InputObject.PSObject.Properties | Where-Object {
            [string]::Equals($_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1
        if ($null -ne $property) {
            Write-Output -NoEnumerate $property.Value
            return
        }
    }

    throw [System.ArgumentException]::new("Configuration is missing required property '$Name'.")
}

function Test-SqlUtilityIntegralValue {
    [CmdletBinding()]
    param($Value)

    if ($null -eq $Value) {
        return $false
    }

    return $Value -is [sbyte] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64]
}

function ConvertTo-SqlUtilityValidatedInteger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Value,
        [Parameter(Mandatory = $true)][int] $Minimum,
        [Parameter(Mandatory = $true)][int] $Maximum,
        [Parameter(Mandatory = $true)][string] $Name
    )

    if (-not (Test-SqlUtilityIntegralValue -Value $Value)) {
        throw [System.ArgumentException]::new("$Name must be an integer.")
    }
    if ($Value -lt $Minimum -or $Value -gt $Maximum) {
        throw [System.ArgumentException]::new("$Name must be between $Minimum and $Maximum.")
    }

    return [int] $Value
}

function New-SqlUtilityDefaultConfig {
    [CmdletBinding()]
    param()

    return [pscustomobject][ordered]@{
        schemaVersion = 1
        unorderedRowLimit = 1000
        queryExportTimeoutSeconds = 120
        connections = @()
    }
}

function ConvertTo-SqlUtilityValidatedConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $InputObject
    )

    $schemaVersion = Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'schemaVersion'
    if (-not (Test-SqlUtilityIntegralValue -Value $schemaVersion)) {
        throw [System.ArgumentException]::new('schemaVersion must be an integer.')
    }
    if ($schemaVersion -ne 1) {
        throw [System.NotSupportedException]::new("Unsupported configuration schema version: $schemaVersion")
    }

    $unorderedRowLimit = ConvertTo-SqlUtilityValidatedInteger -Value (Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'unorderedRowLimit') -Minimum 100 -Maximum 2000 -Name 'unorderedRowLimit'
    $queryExportTimeoutSeconds = ConvertTo-SqlUtilityValidatedInteger -Value (Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'queryExportTimeoutSeconds') -Minimum 5 -Maximum 3600 -Name 'queryExportTimeoutSeconds'
    $rawConnections = Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'connections'
    if ($null -eq $rawConnections) {
        throw [System.ArgumentException]::new('connections must be an array.')
    }

    $connections = New-Object System.Collections.Generic.List[object]
    $seenKeys = @{}
    foreach ($connection in @($rawConnections)) {
        if ($null -eq $connection) {
            throw [System.ArgumentException]::new('connections cannot contain null.')
        }

        $serverValue = Get-SqlUtilityConfigPropertyValue -InputObject $connection -Name 'server'
        $databaseValue = Get-SqlUtilityConfigPropertyValue -InputObject $connection -Name 'database'
        if ($serverValue -isnot [string] -or $databaseValue -isnot [string]) {
            throw [System.ArgumentException]::new('Connection server and database must be strings.')
        }

        $server = $serverValue.Trim()
        $database = $databaseValue.Trim()
        if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($database)) {
            throw [System.ArgumentException]::new('Connection server and database cannot be blank.')
        }

        $connectionKey = $server.ToUpperInvariant() + "`0" + $database.ToUpperInvariant()
        if (-not $seenKeys.ContainsKey($connectionKey)) {
            $seenKeys[$connectionKey] = $true
            [void] $connections.Add([pscustomobject][ordered]@{
                server = $server
                database = $database
            })
        }
    }

    return [pscustomobject][ordered]@{
        schemaVersion = 1
        unorderedRowLimit = $unorderedRowLimit
        queryExportTimeoutSeconds = $queryExportTimeoutSeconds
        connections = @($connections.ToArray())
    }
}

function Read-SqlUtilityConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path
    )

    if (-not [System.IO.File]::Exists($Path)) {
        return New-SqlUtilityDefaultConfig
    }

    $inputObject = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    return ConvertTo-SqlUtilityValidatedConfig -InputObject $inputObject
}

function Write-SqlUtilityConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)] $Config
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
        if ([System.IO.File]::Exists($Path)) {
            [System.IO.File]::Replace($temporaryPath, $Path, $null)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
        return $validated
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) {
            [System.IO.File]::Delete($temporaryPath)
        }
    }
}

function Add-SqlUtilitySavedConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Config,
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database
    )

    $validated = ConvertTo-SqlUtilityValidatedConfig -InputObject $Config
    $trimmedServer = $Server.Trim()
    $trimmedDatabase = $Database.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedServer) -or [string]::IsNullOrWhiteSpace($trimmedDatabase)) {
        throw [System.ArgumentException]::new('Connection server and database cannot be blank.')
    }

    $targetKey = $trimmedServer.ToUpperInvariant() + "`0" + $trimmedDatabase.ToUpperInvariant()
    $connections = New-Object System.Collections.Generic.List[object]
    $replaced = $false
    foreach ($connection in $validated.connections) {
        $connectionKey = $connection.server.ToUpperInvariant() + "`0" + $connection.database.ToUpperInvariant()
        if ($connectionKey -eq $targetKey) {
            if (-not $replaced) {
                [void] $connections.Add([pscustomobject][ordered]@{
                    server = $trimmedServer
                    database = $trimmedDatabase
                })
                $replaced = $true
            }
        }
        else {
            [void] $connections.Add([pscustomobject][ordered]@{
                server = $connection.server
                database = $connection.database
            })
        }
    }
    if (-not $replaced) {
        [void] $connections.Add([pscustomobject][ordered]@{
            server = $trimmedServer
            database = $trimmedDatabase
        })
    }

    return ConvertTo-SqlUtilityValidatedConfig -InputObject ([pscustomobject][ordered]@{
        schemaVersion = $validated.schemaVersion
        unorderedRowLimit = $validated.unorderedRowLimit
        queryExportTimeoutSeconds = $validated.queryExportTimeoutSeconds
        connections = @($connections.ToArray())
    })
}

function Remove-SqlUtilitySavedConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Config,
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database
    )

    $validated = ConvertTo-SqlUtilityValidatedConfig -InputObject $Config
    $trimmedServer = $Server.Trim()
    $trimmedDatabase = $Database.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedServer) -or [string]::IsNullOrWhiteSpace($trimmedDatabase)) {
        throw [System.ArgumentException]::new('Connection server and database cannot be blank.')
    }

    $targetKey = $trimmedServer.ToUpperInvariant() + "`0" + $trimmedDatabase.ToUpperInvariant()
    $connections = New-Object System.Collections.Generic.List[object]
    foreach ($connection in $validated.connections) {
        $connectionKey = $connection.server.ToUpperInvariant() + "`0" + $connection.database.ToUpperInvariant()
        if ($connectionKey -ne $targetKey) {
            [void] $connections.Add([pscustomobject][ordered]@{
                server = $connection.server
                database = $connection.database
            })
        }
    }

    return ConvertTo-SqlUtilityValidatedConfig -InputObject ([pscustomobject][ordered]@{
        schemaVersion = $validated.schemaVersion
        unorderedRowLimit = $validated.unorderedRowLimit
        queryExportTimeoutSeconds = $validated.queryExportTimeoutSeconds
        connections = @($connections.ToArray())
    })
}
