$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')

function Get-TestFileSnapshot {
    param([Parameter(Mandatory = $true)][string] $Root)

    return @(
        Get-ChildItem -LiteralPath $Root -File -Recurse -Force |
            Sort-Object FullName |
            ForEach-Object {
                '{0}|{1}|{2}' -f $_.FullName, $_.Length, (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
    )
}

function Get-TestEnvironmentSnapshot {
    $environment = [System.Environment]::GetEnvironmentVariables(
        [System.EnvironmentVariableTarget]::Process
    )
    return @(
        $environment.Keys |
            Sort-Object { [string] $_ } |
            ForEach-Object { '{0}={1}' -f $_, $environment[$_] }
    )
}

$protectedRuntimeFiles = @(
    'StartSqlUtility.cmd',
    'SqlUtility.ps1',
    'modules\SqlUtility.Config.ps1',
    'modules\SqlUtility.QueryPolicy.ps1',
    'modules\SqlUtility.DataExplorer.ps1',
    'modules\SqlUtility.Database.ps1',
    'modules\SqlUtility.Excel.ps1'
)
$runtimeFiles = @($protectedRuntimeFiles + 'SqlUtility.cat')
foreach ($relativePath in $runtimeFiles) {
    Assert-True (Test-Path -LiteralPath (Join-Path $projectRoot $relativePath) -PathType Leaf) `
        "Production runtime file exists: $relativePath"
}

foreach ($relativePath in @(
    'SqlConnectionPoc.ps1',
    'StartSqlPoc.cmd',
    'tests\Test-SqlConnectionPoc.ps1'
)) {
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $projectRoot $relativePath))) `
        "Superseded POC file is absent: $relativePath"
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('SqlUtility.LauncherTest.{0}' -f [guid]::NewGuid().ToString('N'))
$launcherFolder = Join-Path $temporaryRoot 'Application Folder'
$alternateWorkingDirectory = Join-Path $temporaryRoot 'Caller Folder'
[void] [System.IO.Directory]::CreateDirectory($launcherFolder)
[void] [System.IO.Directory]::CreateDirectory($alternateWorkingDirectory)

try {
    Push-Location $alternateWorkingDirectory
    try {
        . (Join-Path $projectRoot 'SqlUtility.ps1') -NoGui
    }
    finally {
        Pop-Location
    }
    Assert-True ($null -ne (Get-Command New-SqlUtilityMainForm -ErrorAction SilentlyContinue)) `
        'SqlUtility.ps1 resolves its production modules outside the application directory'

    $entryPointValidation = Test-SqlUtilityQuery -Sql 'SELECT a, b FROM dbo.Items WHERE a >= 1'
    Assert-True $entryPointValidation.IsValid `
        'SqlUtility.ps1 validates realistic symbol tokens under its strict-mode runtime'

    $catalogCommand = Get-Command Test-SqlUtilityRuntimeCatalog -ErrorAction SilentlyContinue
    Assert-True ($null -ne $catalogCommand) `
        'SqlUtility.ps1 exposes runtime catalog validation before application startup'
    if ($null -ne $catalogCommand) {
        $sourceCatalogResult = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $projectRoot
        Assert-True $sourceCatalogResult.IsValid `
            'Checked-in runtime files match the checked-in catalog'

        $catalogFolder = Join-Path $temporaryRoot 'Catalog Folder'
        [void] [System.IO.Directory]::CreateDirectory($catalogFolder)
        foreach ($relativePath in $protectedRuntimeFiles) {
            $sourcePath = Join-Path $projectRoot $relativePath
            $destinationPath = Join-Path $catalogFolder $relativePath
            $destinationParent = Split-Path -Parent $destinationPath
            [void] [System.IO.Directory]::CreateDirectory($destinationParent)
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath
        }
        $catalogRuntimePaths = @(
            $protectedRuntimeFiles | ForEach-Object { Join-Path $catalogFolder $_ }
        )
        $catalogPath = Join-Path $catalogFolder 'SqlUtility.cat'
        New-FileCatalog -Path $catalogRuntimePaths -CatalogFilePath $catalogPath `
            -CatalogVersion 2.0 | Out-Null

        $catalogResult = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $catalogFolder
        Assert-True $catalogResult.IsValid `
            'Runtime catalog accepts the unchanged protected files'

        Set-Content -LiteralPath (Join-Path $catalogFolder 'SqlUtility.config.json') `
            -Value '{}' -Encoding UTF8
        $catalogWithConfig = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $catalogFolder
        Assert-True $catalogWithConfig.IsValid `
            'Runtime catalog ignores mutable application configuration'

        $hiddenCatalogPath = Join-Path $catalogFolder 'SqlUtility.hidden.cat'
        Move-Item -LiteralPath $catalogPath -Destination $hiddenCatalogPath
        $missingCatalogResult = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $catalogFolder
        Assert-True (-not $missingCatalogResult.IsValid) `
            'Runtime catalog validation rejects a missing catalog'
        Move-Item -LiteralPath $hiddenCatalogPath -Destination $catalogPath

        $missingRuntimePath = Join-Path $catalogFolder 'modules\SqlUtility.Excel.ps1'
        $hiddenRuntimePath = Join-Path $catalogFolder 'modules\SqlUtility.Excel.hidden.ps1'
        Move-Item -LiteralPath $missingRuntimePath -Destination $hiddenRuntimePath
        $missingRuntimeResult = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $catalogFolder
        Assert-True (-not $missingRuntimeResult.IsValid) `
            'Runtime catalog validation rejects a missing protected file'
        Move-Item -LiteralPath $hiddenRuntimePath -Destination $missingRuntimePath

        Add-Content -LiteralPath $missingRuntimePath `
            -Value '# accidental change' -Encoding UTF8
        $changedCatalogResult = Test-SqlUtilityRuntimeCatalog -ApplicationRoot $catalogFolder
        Assert-True (-not $changedCatalogResult.IsValid) `
            'Runtime catalog rejects a changed protected file'
    }

    $launcherPath = Join-Path $projectRoot 'StartSqlUtility.cmd'
    if (Test-Path -LiteralPath $launcherPath -PathType Leaf) {
        $launcherCopy = Join-Path $launcherFolder 'StartSqlUtility.cmd'
        $probePath = Join-Path $launcherFolder 'SqlUtility.ps1'
        $observationPath = Join-Path $launcherFolder 'launcher-observation.json'
        Copy-Item -LiteralPath $launcherPath -Destination $launcherCopy
        @'
param([switch] $VerifyCatalog)

$observation = [ordered]@{
    ApartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState().ToString()
    ExecutionPolicy = (Get-ExecutionPolicy -Scope Process).ToString()
    ScriptRoot = $PSScriptRoot
    CurrentDirectory = (Get-Location).ProviderPath
    VerifyCatalog = [bool] $VerifyCatalog
}
$observation | ConvertTo-Json -Compress |
    Set-Content -LiteralPath (Join-Path $PSScriptRoot 'launcher-observation.json') -Encoding UTF8
Set-Content -LiteralPath 'launcher-current-directory.marker' -Value 'caller directory' -Encoding ASCII
exit 37
'@ | Set-Content -LiteralPath $probePath -Encoding UTF8

        Push-Location $alternateWorkingDirectory
        try {
            & $launcherCopy
            $launcherExitCode = $LASTEXITCODE
        }
        finally {
            Pop-Location
        }

        Assert-Equal 37 $launcherExitCode 'Launcher returns the SqlUtility.ps1 exit code'
        Assert-True (Test-Path -LiteralPath $observationPath -PathType Leaf) `
            'Launcher executes SqlUtility.ps1 beside the copied launcher'
        if (Test-Path -LiteralPath $observationPath -PathType Leaf) {
            $observation = Get-Content -LiteralPath $observationPath -Raw | ConvertFrom-Json
            Assert-Equal 'STA' $observation.ApartmentState 'Launcher starts Windows PowerShell in STA mode'
            Assert-Equal 'Bypass' $observation.ExecutionPolicy `
                'Launcher uses a process-only execution-policy override'
            Assert-True $observation.VerifyCatalog `
                'Launcher requests runtime catalog verification'
            Assert-Equal $launcherFolder $observation.ScriptRoot `
                'Launcher resolves SqlUtility.ps1 relative to its own location'
            Assert-True (Test-Path -LiteralPath (Join-Path $alternateWorkingDirectory 'launcher-current-directory.marker')) `
                'Launcher preserves the caller current directory for relative paths'
            Assert-True (-not (Test-Path -LiteralPath (Join-Path $launcherFolder 'launcher-current-directory.marker'))) `
                'Launcher does not change the current directory to its own location'
        }
    }

    $forbiddenMutationPattern = '(?im)\b(?:Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|SetEnvironmentVariable|setx(?:\.exe)?|reg\.exe)\b'
    $forbiddenImportPattern = '(?im)^\s*(?:Import-Module\b|using\s+module\b|#requires\s+-modules?\b)'
    foreach ($relativePath in $protectedRuntimeFiles) {
        $fullPath = Join-Path $projectRoot $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }
        $source = Get-Content -LiteralPath $fullPath -Raw
        Assert-True ($source -notmatch $forbiddenMutationPattern) `
            "Runtime file does not contain registry or environment mutation APIs: $relativePath"
        Assert-True ($source -notmatch $forbiddenImportPattern) `
            "Runtime file does not import PowerShell modules: $relativePath"
    }

    $filesBefore = Get-TestFileSnapshot -Root $projectRoot
    $environmentBefore = Get-TestEnvironmentSnapshot
    $services = @{
        TestConnection = { throw 'Unexpected TestConnection call during form construction.' }
        WriteConfig = { throw 'Unexpected WriteConfig call during form construction.' }
        ShowMessage = { throw 'Unexpected ShowMessage call during form construction.' }
        Confirm = { throw 'Unexpected Confirm call during form construction.' }
        ValidateQuery = { throw 'Unexpected ValidateQuery call during form construction.' }
        ExecuteOrderedPage = { throw 'Unexpected ExecuteOrderedPage call during form construction.' }
        ExecuteUnordered = { throw 'Unexpected ExecuteUnordered call during form construction.' }
        GetLocalPage = { throw 'Unexpected GetLocalPage call during form construction.' }
        BuildCountSql = { throw 'Unexpected BuildCountSql call during form construction.' }
        ExecuteCount = { throw 'Unexpected ExecuteCount call during form construction.' }
        ExportResult = { throw 'Unexpected ExportResult call during form construction.' }
        PromptSavePath = { throw 'Unexpected PromptSavePath call during form construction.' }
        ListPhysicalTables = { throw 'Unexpected ListPhysicalTables call during form construction.' }
        GetTableColumns = { throw 'Unexpected GetTableColumns call during form construction.' }
        BuildDataExplorerQuery = { throw 'Unexpected BuildDataExplorerQuery call during form construction.' }
        ExecuteDataPreview = { throw 'Unexpected ExecuteDataPreview call during form construction.' }
        ExportPreview = { throw 'Unexpected ExportPreview call during form construction.' }
    }
    $form = New-SqlUtilityMainForm -Config (New-SqlUtilityDefaultConfig) `
        -ConfigPath (Join-Path $projectRoot 'SqlUtility.config.json') -Services $services
    try {
        Assert-True ($null -ne $form) 'Form constructs with an in-memory config and fake services'
    }
    finally {
        if ($null -ne $form) { $form.Dispose() }
    }
    $environmentAfter = Get-TestEnvironmentSnapshot
    $filesAfter = Get-TestFileSnapshot -Root $projectRoot

    Assert-Equal ($filesBefore -join "`n") ($filesAfter -join "`n") `
        'Constructing and disposing the form does not create or modify files'
    Assert-Equal ($environmentBefore -join "`n") ($environmentAfter -join "`n") `
        'Constructing and disposing the form does not change process environment entries'
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        $resolvedTemporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
        $resolvedSystemTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if ($resolvedTemporaryRoot.StartsWith($resolvedSystemTemp, [System.StringComparison]::OrdinalIgnoreCase) -and
            (Split-Path -Leaf $resolvedTemporaryRoot) -like 'SqlUtility.LauncherTest.*') {
            [System.IO.Directory]::Delete($resolvedTemporaryRoot, $true)
        }
    }
}

Complete-TestFile 'Launcher and distribution tests passed.'
