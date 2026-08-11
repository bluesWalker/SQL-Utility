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

$runtimeFiles = @(
    'StartSqlUtility.cmd',
    'SqlUtility.ps1',
    'modules\SqlUtility.Config.ps1',
    'modules\SqlUtility.QueryPolicy.ps1',
    'modules\SqlUtility.Database.ps1',
    'modules\SqlUtility.Excel.ps1'
)
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

    $launcherPath = Join-Path $projectRoot 'StartSqlUtility.cmd'
    if (Test-Path -LiteralPath $launcherPath -PathType Leaf) {
        $launcherCopy = Join-Path $launcherFolder 'StartSqlUtility.cmd'
        $probePath = Join-Path $launcherFolder 'SqlUtility.ps1'
        $observationPath = Join-Path $launcherFolder 'launcher-observation.json'
        Copy-Item -LiteralPath $launcherPath -Destination $launcherCopy
        @'
$observation = [ordered]@{
    ApartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState().ToString()
    ExecutionPolicy = (Get-ExecutionPolicy -Scope Process).ToString()
    ScriptRoot = $PSScriptRoot
    CurrentDirectory = (Get-Location).ProviderPath
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
    foreach ($relativePath in $runtimeFiles) {
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
