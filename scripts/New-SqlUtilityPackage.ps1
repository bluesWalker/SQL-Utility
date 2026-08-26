[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $DestinationPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$relativeRuntimePaths = @(
    'StartSqlUtility.cmd',
    'SqlUtility.ps1',
    'modules\SqlUtility.Config.ps1',
    'modules\SqlUtility.QueryPolicy.ps1',
    'modules\SqlUtility.DataExplorer.ps1',
    'modules\SqlUtility.Database.ps1',
    'modules\SqlUtility.Excel.ps1'
)
$runtimePaths = @(
    $relativeRuntimePaths | ForEach-Object { Join-Path $projectRoot $_ }
)
foreach ($runtimePath in $runtimePaths) {
    if (-not (Test-Path -LiteralPath $runtimePath -PathType Leaf)) {
        throw [System.IO.FileNotFoundException]::new(
            ('Required runtime file is missing: {0}' -f $runtimePath),
            $runtimePath
        )
    }
}

$resolvedDestinationPath = [System.IO.Path]::GetFullPath($DestinationPath)
if ([System.IO.Path]::GetExtension($resolvedDestinationPath) -ine '.zip') {
    throw [System.ArgumentException]::new('DestinationPath must end with .zip.')
}
$destinationDirectory = Split-Path -Parent $resolvedDestinationPath
if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
    throw [System.IO.DirectoryNotFoundException]::new(
        ('The destination directory does not exist: {0}' -f $destinationDirectory)
    )
}
if (Test-Path -LiteralPath $resolvedDestinationPath) {
    throw [System.IO.IOException]::new(
        ('The package destination already exists: {0}' -f $resolvedDestinationPath)
    )
}

$catalogPath = Join-Path $projectRoot 'SqlUtility.cat'
$catalogIsCurrent = $false
if (Test-Path -LiteralPath $catalogPath -PathType Leaf) {
    try {
        $catalogIsCurrent = [string] (Test-FileCatalog -Path $runtimePaths `
            -CatalogFilePath $catalogPath -ErrorAction Stop) -eq 'Valid'
    }
    catch {
        $catalogIsCurrent = $false
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('SqlUtility.Package.{0}' -f [guid]::NewGuid().ToString('N'))
$stagingRoot = Join-Path $temporaryRoot 'SQL Utility'
$temporaryZipPath = Join-Path $destinationDirectory `
    ('.{0}.{1}.tmp' -f [System.IO.Path]::GetFileName($resolvedDestinationPath),
        [guid]::NewGuid().ToString('N'))
$candidateCatalogPath = Join-Path $projectRoot `
    ('.SqlUtility.catalog.{0}.cat' -f [guid]::NewGuid().ToString('N'))
$catalogBackupPath = Join-Path $projectRoot `
    ('.SqlUtility.catalog.{0}.bak' -f [guid]::NewGuid().ToString('N'))

try {
    if (-not $catalogIsCurrent) {
        New-FileCatalog -Path $runtimePaths -CatalogFilePath $candidateCatalogPath `
            -CatalogVersion 2.0 | Out-Null
        $candidateStatus = Test-FileCatalog -Path $runtimePaths `
            -CatalogFilePath $candidateCatalogPath -ErrorAction Stop
        if ([string] $candidateStatus -ne 'Valid') {
            throw [System.IO.InvalidDataException]::new(
                'The generated runtime catalog did not validate the source runtime files.'
            )
        }

        if (Test-Path -LiteralPath $catalogPath -PathType Leaf) {
            [System.IO.File]::Replace(
                $candidateCatalogPath,
                $catalogPath,
                $catalogBackupPath,
                $true
            )
            try {
                [System.IO.File]::Delete($catalogBackupPath)
            }
            catch {
                Write-Warning ('The catalog was updated, but its backup could not be removed: {0}' -f
                    $catalogBackupPath)
            }
        }
        else {
            [System.IO.File]::Move($candidateCatalogPath, $catalogPath)
        }
    }

    [void] [System.IO.Directory]::CreateDirectory($stagingRoot)
    foreach ($relativePath in $relativeRuntimePaths) {
        $sourcePath = Join-Path $projectRoot $relativePath
        $stagingPath = Join-Path $stagingRoot $relativePath
        [void] [System.IO.Directory]::CreateDirectory((Split-Path -Parent $stagingPath))
        [System.IO.File]::Copy($sourcePath, $stagingPath, $false)
    }
    [System.IO.File]::Copy($catalogPath, (Join-Path $stagingRoot 'SqlUtility.cat'), $false)

    $stagedRuntimePaths = @(
        $relativeRuntimePaths | ForEach-Object { Join-Path $stagingRoot $_ }
    )
    $stagedStatus = Test-FileCatalog -Path $stagedRuntimePaths `
        -CatalogFilePath (Join-Path $stagingRoot 'SqlUtility.cat') -ErrorAction Stop
    if ([string] $stagedStatus -ne 'Valid') {
        throw [System.IO.InvalidDataException]::new(
            'The staged runtime files did not match SqlUtility.cat.'
        )
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $stagingRoot,
        $temporaryZipPath,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false
    )
    [System.IO.File]::Move($temporaryZipPath, $resolvedDestinationPath)
    Write-Output $resolvedDestinationPath
}
finally {
    foreach ($transientPath in @(
        $temporaryZipPath,
        $candidateCatalogPath,
        $catalogBackupPath
    )) {
        if (Test-Path -LiteralPath $transientPath -PathType Leaf) {
            try { [System.IO.File]::Delete($transientPath) } catch { }
        }
    }
    if (Test-Path -LiteralPath $temporaryRoot -PathType Container) {
        $resolvedTemporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
        $resolvedSystemTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if ($resolvedTemporaryRoot.StartsWith(
                $resolvedSystemTemp,
                [System.StringComparison]::OrdinalIgnoreCase
            ) -and
            (Split-Path -Leaf $resolvedTemporaryRoot) -like 'SqlUtility.Package.*') {
            try { [System.IO.Directory]::Delete($resolvedTemporaryRoot, $true) } catch { }
        }
    }
}
