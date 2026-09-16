function Initialize-SqlUtilityTemplateDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string] $Path)

    [void] [System.IO.Directory]::CreateDirectory($Path)
}

function Get-SqlUtilityTemplatePath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string] $Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ([System.IO.Path]::GetExtension($fullPath) -ine '.sql') {
        throw [System.ArgumentException]::new('Select a template file with the .sql extension.')
    }
    return $fullPath
}

function Read-SqlUtilityTemplate {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string] $Path)

    $fullPath = Get-SqlUtilityTemplatePath -Path $Path
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $encoding = [System.Text.UTF8Encoding]::new($false, $true)
    $offset = 0
    # Decode explicitly so BOM detection also retains strict error handling.
    if ($bytes.Length -ge 4 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe -and $bytes[2] -eq 0 -and $bytes[3] -eq 0) {
        $encoding = [System.Text.UTF32Encoding]::new($false, $false, $true); $offset = 4
    }
    elseif ($bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 0xfe -and $bytes[3] -eq 0xff) {
        $encoding = [System.Text.UTF32Encoding]::new($true, $false, $true); $offset = 4
    }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) {
        $offset = 3
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) {
        $encoding = [System.Text.UnicodeEncoding]::new($false, $false, $true); $offset = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff) {
        $encoding = [System.Text.UnicodeEncoding]::new($true, $false, $true); $offset = 2
    }
    $text = $encoding.GetString($bytes, $offset, $bytes.Length - $offset)
    if ($text.IndexOf([char] 0) -ge 0) {
        throw [System.IO.InvalidDataException]::new('The template contains a NUL character and cannot be loaded into the editor.')
    }
    return $text
}

function Remove-SqlUtilityTemplateTransientFile {
    param([Parameter(Mandatory = $true)][string] $Path)
    [System.IO.File]::Delete($Path)
}

function Write-SqlUtilityTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $Text,
        [bool] $AllowOverwrite = $false
    )

    $fullPath = Get-SqlUtilityTemplatePath -Path $Path
    if ([string]::IsNullOrWhiteSpace($Text) -or $Text.IndexOf([char] 0) -ge 0) {
        throw [System.ArgumentException]::new('Enter template text without NUL characters before saving.')
    }
    $directory = [System.IO.Path]::GetDirectoryName($fullPath)
    if (-not [System.IO.Directory]::Exists($directory)) {
        throw [System.IO.DirectoryNotFoundException]::new("Template directory does not exist: $directory")
    }
    $destinationExists = [System.IO.File]::Exists($fullPath)
    if ($destinationExists -and -not $AllowOverwrite) {
        throw [System.IO.IOException]::new('The template already exists. Save again and confirm replacement.')
    }
    $temporaryPath = Join-Path $directory ('.SqlUtility.template.{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    $backupPath = Join-Path $directory ('.SqlUtility.template.{0}.bak' -f [guid]::NewGuid().ToString('N'))
    $committed = $false
    $primaryError = $null
    $cleanupWarning = ''
    try {
        [System.IO.File]::WriteAllText($temporaryPath, $Text, [System.Text.UTF8Encoding]::new($true, $true))
        if ($destinationExists) {
            [System.IO.File]::Replace($temporaryPath, $fullPath, $backupPath)
        }
        else {
            # Move does not overwrite a file that appeared after the existence check.
            [System.IO.File]::Move($temporaryPath, $fullPath)
        }
        $committed = $true
    }
    catch { $primaryError = $_ }
    finally {
        $cleanupPaths = @($temporaryPath)
        if ($committed) { $cleanupPaths += $backupPath }
        foreach ($cleanupPath in $cleanupPaths) {
            if ([System.IO.File]::Exists($cleanupPath)) {
                try { Remove-SqlUtilityTemplateTransientFile -Path $cleanupPath }
                catch {
                    $cleanupWarning = "The template was saved, but a temporary or backup file could not be removed: $cleanupPath`r`n$($_.Exception.Message)"
                }
            }
        }
    }
    if ($null -ne $primaryError) { $PSCmdlet.ThrowTerminatingError($primaryError) }
    return [pscustomobject]@{ Path = $fullPath; CleanupWarning = $cleanupWarning }
}

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
        schemaVersion = 3
        previewRowLimit = 100
        resultDataLimitMiB = 256
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
    if ($schemaVersion -eq 1) {
        $previewRowLimit = 100
        $resultDataLimitMiB = 256
    }
    elseif ($schemaVersion -eq 2) {
        $previewRowLimit = ConvertTo-SqlUtilityValidatedInteger `
            -Value (Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'previewRowLimit') `
            -Minimum 10 -Maximum 500 -Name 'previewRowLimit'
        $resultDataLimitMiB = 256
    }
    elseif ($schemaVersion -eq 3) {
        $previewRowLimit = ConvertTo-SqlUtilityValidatedInteger `
            -Value (Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'previewRowLimit') `
            -Minimum 10 -Maximum 500 -Name 'previewRowLimit'
        $resultDataLimitMiB = ConvertTo-SqlUtilityValidatedInteger `
            -Value (Get-SqlUtilityConfigPropertyValue -InputObject $InputObject -Name 'resultDataLimitMiB') `
            -Minimum 128 -Maximum 1024 -Name 'resultDataLimitMiB'
    }
    else {
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
        schemaVersion = 3
        previewRowLimit = $previewRowLimit
        resultDataLimitMiB = $resultDataLimitMiB
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

    try {
        $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    }
    catch [System.Management.Automation.ItemNotFoundException] {
        return New-SqlUtilityDefaultConfig
    }

    try {
        $inputObject = $json | ConvertFrom-Json
        return ConvertTo-SqlUtilityValidatedConfig -InputObject $inputObject
    }
    catch [System.NotSupportedException] {
        throw
    }
    catch {
        throw [System.IO.InvalidDataException]::new(
            ("The configuration file is malformed or does not match the required schema. {0}" -f $_.Exception.Message),
            $_.Exception
        )
    }
}

function Remove-SqlUtilityConfigTransientFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $TransientPath
    )

    [System.IO.File]::Delete($TransientPath)
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
    $backupPath = Join-Path $directory ('.SqlUtility.config.{0}.bak' -f [guid]::NewGuid().ToString('N'))
    $primaryError = $null
    $cleanupError = $null
    $destinationCommitted = $false
    try {
        $json = $validated | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($temporaryPath, $json, [System.Text.UTF8Encoding]::new($false))
        if ([System.IO.File]::Exists($Path)) {
            [System.IO.File]::Replace($temporaryPath, $Path, $backupPath)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
        $destinationCommitted = $true
    }
    catch {
        $primaryError = $_
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) {
            try {
                Remove-SqlUtilityConfigTransientFile -TransientPath $temporaryPath
            }
            catch {
                if ($null -eq $cleanupError) { $cleanupError = $_ }
            }
        }
        if ($destinationCommitted -and [System.IO.File]::Exists($backupPath)) {
            try {
                Remove-SqlUtilityConfigTransientFile -TransientPath $backupPath
            }
            catch {
                if ($null -eq $cleanupError) { $cleanupError = $_ }
            }
        }
    }

    if ($null -ne $primaryError) {
        $PSCmdlet.ThrowTerminatingError($primaryError)
    }
    if ($null -ne $cleanupError) {
        if ($destinationCommitted) {
            Write-Warning ("The configuration was saved, but a transient file could not be removed: {0}" -f $cleanupError.Exception.Message)
            return $validated
        }
        $PSCmdlet.ThrowTerminatingError($cleanupError)
    }

    return $validated
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
        previewRowLimit = $validated.previewRowLimit
        resultDataLimitMiB = $validated.resultDataLimitMiB
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
        previewRowLimit = $validated.previewRowLimit
        resultDataLimitMiB = $validated.resultDataLimitMiB
        unorderedRowLimit = $validated.unorderedRowLimit
        queryExportTimeoutSeconds = $validated.queryExportTimeoutSeconds
        connections = @($connections.ToArray())
    })
}
