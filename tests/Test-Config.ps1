$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.Config.ps1')

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SqlUtilityConfigTests-' + [guid]::NewGuid().ToString('N'))
[void] [System.IO.Directory]::CreateDirectory($testRoot)

try {
    $defaults = New-SqlUtilityDefaultConfig
    Assert-Equal 3 $defaults.schemaVersion 'Default schema version'
    Assert-Equal 100 $defaults.previewRowLimit 'Default preview limit'
    Assert-Equal 256 $defaults.resultDataLimitMiB 'Default result data limit'
    Assert-Equal 1000 $defaults.unorderedRowLimit 'Default unordered limit'
    Assert-Equal 120 $defaults.queryExportTimeoutSeconds 'Default timeout'
    Assert-Equal 0 @($defaults.connections).Count 'Default saved connections'

    $version1 = [pscustomobject][ordered]@{
        schemaVersion = 1
        previewRowLimit = 500
        unorderedRowLimit = 1200
        queryExportTimeoutSeconds = 90
        connections = @([pscustomobject]@{ server = 'ServerA'; database = 'DbA' })
    }
    $migrated = ConvertTo-SqlUtilityValidatedConfig -InputObject $version1
    Assert-Equal 3 $migrated.schemaVersion 'Version 1 migrates in memory'
    Assert-Equal 100 $migrated.previewRowLimit 'Version 1 receives default preview limit'
    Assert-Equal 256 $migrated.resultDataLimitMiB 'Version 1 receives default result data limit'
    Assert-Equal 1200 $migrated.unorderedRowLimit 'Migration preserves unordered limit'
    Assert-Equal 100 (ConvertTo-SqlUtilityValidatedConfig -InputObject $version1).previewRowLimit 'Version 1 ignores extra preview limit'
    $version1Path = Join-Path $testRoot 'version1.json'
    $version1Json = '{"schemaVersion":1,"unorderedRowLimit":1200,"queryExportTimeoutSeconds":90,"connections":[{"server":"ServerA","database":"DbA"}]}'
    [System.IO.File]::WriteAllText($version1Path, $version1Json, [System.Text.UTF8Encoding]::new($false))
    $version1Bytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($version1Path))
    $readVersion1 = Read-SqlUtilityConfig -Path $version1Path
    Assert-Equal 3 $readVersion1.schemaVersion 'Read migrates version 1 in memory'
    Assert-Equal $version1Bytes ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($version1Path))) 'Version 1 read does not rewrite bytes'

    $version2 = [pscustomobject][ordered]@{
        schemaVersion = 2
        previewRowLimit = 250
        unorderedRowLimit = 1400
        queryExportTimeoutSeconds = 180
        connections = @()
    }
    $migratedVersion2 = ConvertTo-SqlUtilityValidatedConfig -InputObject $version2
    Assert-Equal 3 $migratedVersion2.schemaVersion 'Version 2 migrates in memory'
    Assert-Equal 250 $migratedVersion2.previewRowLimit 'Version 2 migration preserves preview limit'
    Assert-Equal 256 $migratedVersion2.resultDataLimitMiB 'Version 2 receives default result data limit'

    foreach ($validResultDataLimit in @(128, 256, 1024)) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.resultDataLimitMiB = $validResultDataLimit
        Assert-Equal $validResultDataLimit (ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate).resultDataLimitMiB `
            "Accept result data limit $validResultDataLimit"
    }
    foreach ($invalidResultDataLimit in @(127, 1025, 128.5, '256')) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.resultDataLimitMiB = $invalidResultDataLimit
        Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate } `
            'System.ArgumentException' "Reject result data limit $invalidResultDataLimit"
    }

    foreach ($validPreviewLimit in @(10, 100, 500)) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.previewRowLimit = $validPreviewLimit
        Assert-Equal $validPreviewLimit (ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate).previewRowLimit `
            "Accept preview limit $validPreviewLimit"
    }
    foreach ($invalidPreviewLimit in @(9, 501, 10.5, '100')) {
        $candidate = New-SqlUtilityDefaultConfig
        $candidate.previewRowLimit = $invalidPreviewLimit
        Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $candidate } `
            'System.ArgumentException' "Reject preview limit $invalidPreviewLimit"
    }

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
    Assert-Equal 100 $reloadedReplacement.previewRowLimit 'Write round-trip emits schema 3 preview limit'
    Assert-Equal 256 $reloadedReplacement.resultDataLimitMiB 'Write round-trip emits schema 3 result data limit'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.tmp' -File).Count 'Successful replacement leaves no temporary sibling'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.bak' -File).Count 'Successful replacement leaves no backup sibling'

    $cleanupFailureConfig = ConvertTo-SqlUtilityValidatedConfig -InputObject $replaced
    $cleanupFailureConfig.unorderedRowLimit = 1600
    $script:ConfigCleanupPaths = [System.Collections.Generic.List[string]]::new()
    $originalCleanupCommand = Get-Command -Name Remove-SqlUtilityConfigTransientFile -CommandType Function -ErrorAction SilentlyContinue
    $originalCleanupScript = if ($null -ne $originalCleanupCommand) { $originalCleanupCommand.ScriptBlock } else { $null }
    try {
        Set-Item -Path Function:\Remove-SqlUtilityConfigTransientFile -Value {
            param([string] $TransientPath)
            [void] $script:ConfigCleanupPaths.Add($TransientPath)
            if ([System.IO.Path]::GetExtension($TransientPath) -eq '.bak') {
                throw [System.IO.IOException]::new('simulated backup cleanup failure')
            }
            if ([System.IO.File]::Exists($TransientPath)) {
                [System.IO.File]::Delete($TransientPath)
            }
        }

        $cleanupWarnings = @()
        $cleanupCommitError = $null
        $committedAfterCleanupFailure = $null
        try {
            $committedAfterCleanupFailure = Write-SqlUtilityConfig -Path $roundTripPath `
                -Config $cleanupFailureConfig -WarningVariable cleanupWarnings -WarningAction SilentlyContinue
        }
        catch {
            $cleanupCommitError = $_
        }

        Assert-True ($null -eq $cleanupCommitError) 'Backup cleanup failure does not report a committed save as failed'
        if ($null -ne $committedAfterCleanupFailure) {
            Assert-Equal 1600 $committedAfterCleanupFailure.unorderedRowLimit `
                'Backup cleanup failure still returns the committed configuration'
        }
        Assert-Equal 1600 (Read-SqlUtilityConfig -Path $roundTripPath).unorderedRowLimit `
            'Backup cleanup failure leaves the new configuration committed'
        Assert-Equal 100 (Read-SqlUtilityConfig -Path $roundTripPath).previewRowLimit `
            'Backup cleanup failure retains preview limit'
        Assert-Equal 1 @($cleanupWarnings).Count 'Backup cleanup failure is reported separately as one warning'
        Assert-True (@($cleanupWarnings)[0].Message -match 'saved') `
            'Backup cleanup warning identifies the save as committed'
        $remainingBackups = @(Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.bak' -File)
        Assert-Equal 1 $remainingBackups.Count 'Interrupted backup cleanup leaves one recoverable sibling backup'
        if ($remainingBackups.Count -eq 1) {
            Assert-Equal $testRoot $remainingBackups[0].DirectoryName 'Transient backup is a sibling of the configuration file'
        }
        Assert-Equal 1 @($script:ConfigCleanupPaths | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.bak' }).Count `
            'Configuration replacement attempts backup cleanup exactly once'
    }
    finally {
        if ($null -ne $originalCleanupScript) {
            Set-Item -Path Function:\Remove-SqlUtilityConfigTransientFile -Value $originalCleanupScript
        }
        else {
            Remove-Item -Path Function:\Remove-SqlUtilityConfigTransientFile -ErrorAction SilentlyContinue
        }
        Get-ChildItem -LiteralPath $testRoot -Filter '.SqlUtility.config.*.bak' -File | Remove-Item -Force
        Remove-Variable -Name ConfigCleanupPaths -Scope Script -ErrorAction SilentlyContinue
    }

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

    $missingPreviewConfig = New-SqlUtilityDefaultConfig
    $missingPreviewConfig.PSObject.Properties.Remove('previewRowLimit')
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $missingPreviewConfig } 'System.ArgumentException' 'Schema 2 missing preview limit is corruption'

    $futureSchemaConfig = New-SqlUtilityDefaultConfig
    $futureSchemaConfig.schemaVersion = 4
    Assert-Throws { ConvertTo-SqlUtilityValidatedConfig -InputObject $futureSchemaConfig } 'System.NotSupportedException' 'Reject unsupported schema version'

    $previewConfig = New-SqlUtilityDefaultConfig
    $previewConfig.previewRowLimit = 250
    $previewConfig.resultDataLimitMiB = 512
    $previewWithPair = Add-SqlUtilitySavedConnection -Config $previewConfig -Server 'ServerA' -Database 'DbA'
    Assert-Equal 250 $previewWithPair.previewRowLimit 'Add preserves preview limit'
    Assert-Equal 512 $previewWithPair.resultDataLimitMiB 'Add preserves result data limit'
    $previewRemoved = Remove-SqlUtilitySavedConnection -Config $previewWithPair -Server 'ServerA' -Database 'DbA'
    Assert-Equal 250 $previewRemoved.previewRowLimit 'Remove preserves preview limit'
    Assert-Equal 512 $previewRemoved.resultDataLimitMiB 'Remove preserves result data limit'

    $malformedPath = Join-Path $testRoot 'malformed.json'
    [System.IO.File]::WriteAllText($malformedPath, '{ malformed json')
    Assert-Throws { Read-SqlUtilityConfig -Path $malformedPath } 'System.IO.InvalidDataException' `
        'Malformed JSON is classified as recoverable configuration corruption'

    $invalidSchemaPath = Join-Path $testRoot 'invalid-schema.json'
    [System.IO.File]::WriteAllText(
        $invalidSchemaPath,
        '{"schemaVersion":1,"unorderedRowLimit":1000,"queryExportTimeoutSeconds":120}',
        [System.Text.UTF8Encoding]::new($false)
    )
    Assert-Throws { Read-SqlUtilityConfig -Path $invalidSchemaPath } 'System.IO.InvalidDataException' `
        'Invalid persisted schema is classified as recoverable configuration corruption'

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
