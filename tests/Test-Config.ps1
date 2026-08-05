$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.Config.ps1')

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SqlUtilityConfigTests-' + [guid]::NewGuid().ToString('N'))
[void] [System.IO.Directory]::CreateDirectory($testRoot)

try {
    $defaults = New-SqlUtilityDefaultConfig
    Assert-Equal 1 $defaults.schemaVersion 'Default schema version'
    Assert-Equal 1000 $defaults.unorderedRowLimit 'Default unordered limit'
    Assert-Equal 120 $defaults.queryExportTimeoutSeconds 'Default timeout'
    Assert-Equal 0 @($defaults.connections).Count 'Default saved connections'

    $missingPath = Join-Path $testRoot 'missing.json'
    $missing = Read-SqlUtilityConfig -Path $missingPath
    Assert-Equal 1000 $missing.unorderedRowLimit 'Missing config uses defaults'
    Assert-True (-not (Test-Path -LiteralPath $missingPath)) 'Read does not create config'

    $invalidPath = Join-Path $testRoot 'invalid|config.json'
    Assert-Throws { Read-SqlUtilityConfig -Path $invalidPath } 'System.ArgumentException' 'Invalid config path is not treated as missing'

    $withPair = Add-SqlUtilitySavedConnection -Config $defaults -Server ' ServerA ' -Database ' DbA '
    $deduplicated = Add-SqlUtilitySavedConnection -Config $withPair -Server 'servera' -Database 'dba'
    Assert-Equal 1 @($deduplicated.connections).Count 'Connection pair is unique case-insensitively'
    Assert-Equal 'servera' $deduplicated.connections[0].server 'Latest server casing is kept'
    Assert-Equal 'dba' $deduplicated.connections[0].database 'Latest database casing is kept'
    Assert-True (-not [object]::ReferenceEquals($defaults, $withPair)) 'Add returns a new config object'
    Assert-Equal 0 @($defaults.connections).Count 'Add does not mutate the input config'

    $roundTripPath = Join-Path $testRoot 'SqlUtility.config.json'
    $saved = Write-SqlUtilityConfig -Path $roundTripPath -Config $deduplicated
    $loaded = Read-SqlUtilityConfig -Path $roundTripPath
    Assert-Equal ($saved | ConvertTo-Json -Depth 4) ($loaded | ConvertTo-Json -Depth 4) 'Config round-trips'

    $replacementConfig = ConvertTo-SqlUtilityValidatedConfig -InputObject $saved
    $replacementConfig.unorderedRowLimit = 1500
    $replaced = Write-SqlUtilityConfig -Path $roundTripPath -Config $replacementConfig
    $reloadedReplacement = Read-SqlUtilityConfig -Path $roundTripPath
    Assert-Equal 1500 $replaced.unorderedRowLimit 'Existing config replacement returns updated settings'
    Assert-Equal 1500 $reloadedReplacement.unorderedRowLimit 'Existing config replacement persists updated settings'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.tmp' -File).Count 'Successful replacement leaves no temporary sibling'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.bak' -File).Count 'Successful replacement leaves no backup sibling'

    $readLock = [System.IO.File]::Open($roundTripPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        Assert-Throws { Read-SqlUtilityConfig -Path $roundTripPath } $null 'Locked configuration read propagates its I/O error'
    }
    finally {
        $readLock.Dispose()
    }

    foreach ($validLimit in @(100, 2000)) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.unorderedRowLimit = $validLimit
        $validated = ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate
        Assert-Equal $validLimit $validated.unorderedRowLimit "Accept limit $validLimit"
    }
    foreach ($invalidLimit in @(99, 2001, 100.5, '100')) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.unorderedRowLimit = $invalidLimit
        Assert-Throws { ConvertTo-SqlUtilityValidatedConfig $candidate } 'System.ArgumentException' "Reject limit $invalidLimit"
    }

    foreach ($validTimeout in @(5, 3600)) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.queryExportTimeoutSeconds = $validTimeout
        $validated = ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate
        Assert-Equal $validTimeout $validated.queryExportTimeoutSeconds "Accept timeout $validTimeout"
    }
    foreach ($invalidTimeout in @(4, 3601, 5.5, '120')) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.queryExportTimeoutSeconds = $invalidTimeout
        Assert-Throws { ConvertTo-SqlUtilityValidatedConfig $candidate } 'System.ArgumentException' "Reject timeout $invalidTimeout"
    }

    Assert-Throws { Add-SqlUtilitySavedConnection -Config $defaults -Server '  ' -Database 'db' } 'System.ArgumentException' 'Reject blank saved-connection server'
    Assert-Throws { Add-SqlUtilitySavedConnection -Config $defaults -Server 'server' -Database '  ' } 'System.ArgumentException' 'Reject blank saved-connection database'

    $blankConnectionConfig = New-SqlUtilityDefaultConfig
    $blankConnectionConfig.connections = @([pscustomobject]@{ server = 'server'; database = ' ' })
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $blankConnectionConfig } 'System.ArgumentException' 'Reject blank persisted connection database'

    $missingPropertyConfig = [pscustomobject]@{
        schemaVersion = 1
        unorderedRowLimit = 1000
        queryExportTimeoutSeconds = 120
    }
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $missingPropertyConfig } 'System.ArgumentException' 'Reject config missing connections'

    $missingSchemaVersionConfig = [pscustomobject]@{
        unorderedRowLimit = 1000
        queryExportTimeoutSeconds = 120
        connections = @()
    }
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $missingSchemaVersionConfig } 'System.ArgumentException' 'Reject config missing schemaVersion'

    $missingLimitConfig = [pscustomobject]@{
        schemaVersion = 1
        queryExportTimeoutSeconds = 120
        connections = @()
    }
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $missingLimitConfig } 'System.ArgumentException' 'Reject config missing unorderedRowLimit'

    $missingTimeoutConfig = [pscustomobject]@{
        schemaVersion = 1
        unorderedRowLimit = 1000
        connections = @()
    }
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $missingTimeoutConfig } 'System.ArgumentException' 'Reject config missing queryExportTimeoutSeconds'

    $futureSchemaConfig = New-SqlUtilityDefaultConfig
    $futureSchemaConfig.schemaVersion = 2
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $futureSchemaConfig } 'System.NotSupportedException' 'Reject unsupported schema version'

    $malformedPath = Join-Path $testRoot 'malformed.json'
    [System.IO.File]::WriteAllText($malformedPath, '{ malformed json')
    Assert-Throws { Read-SqlUtilityConfig -Path $malformedPath } 'System.ArgumentException' 'Malformed JSON is rejected'

    $withSecondPair = Add-SqlUtilitySavedConnection -Config $deduplicated -Server 'ServerB' -Database 'DbB'
    $removed = Remove-SqlUtilitySavedConnection -Config $withSecondPair -Server 'SERVERA' -Database 'DBA'
    Assert-Equal 1 @($removed.connections).Count 'Remove deletes the matching connection pair'
    Assert-Equal 'ServerB' $removed.connections[0].server 'Remove keeps other saved connections'
    Assert-True (-not [object]::ReferenceEquals($withSecondPair, $removed)) 'Remove returns a new config object'
    Assert-Equal 2 @($withSecondPair.connections).Count 'Remove does not mutate the input config'

    $beforeFailureBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($roundTripPath))
    $updatedConfig = Add-SqlUtilitySavedConnection -Config $saved -Server 'ServerC' -Database 'DbC'
    $destinationLock = [System.IO.File]::Open($roundTripPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        Assert-Throws { Write-SqlUtilityConfig -Path $roundTripPath -Config $updatedConfig } $null 'Locked destination prevents replacement'
    }
    finally {
        $destinationLock.Dispose()
    }
    $afterFailureBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($roundTripPath))
    Assert-Equal $beforeFailureBytes $afterFailureBytes 'Failed replacement preserves existing bytes'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.tmp' -File).Count 'Failed replacement cleans temporary files'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

Complete-TestFile 'All configuration tests passed.'
