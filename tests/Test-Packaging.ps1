$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')

$protectedRuntimeFiles = @(
    'StartSqlUtility.cmd',
    'SqlUtility.ps1',
    'modules\SqlUtility.Config.ps1',
    'modules\SqlUtility.QueryPolicy.ps1',
    'modules\SqlUtility.DataExplorer.ps1',
    'modules\SqlUtility.Database.ps1',
    'modules\SqlUtility.Excel.ps1'
)
$packageScript = Join-Path $projectRoot 'scripts\New-SqlUtilityPackage.ps1'
Assert-True (Test-Path -LiteralPath $packageScript -PathType Leaf) `
    'Packaging script exists'

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('SqlUtility.PackagingTest.{0}' -f [guid]::NewGuid().ToString('N'))
$fixtureRoot = Join-Path $temporaryRoot 'Project'
$packageOutputFolder = Join-Path $temporaryRoot 'Output'
$extractedFolder = Join-Path $temporaryRoot 'Extracted'
[void] [System.IO.Directory]::CreateDirectory($fixtureRoot)
[void] [System.IO.Directory]::CreateDirectory($packageOutputFolder)

try {
    if (Test-Path -LiteralPath $packageScript -PathType Leaf) {
        foreach ($relativePath in $protectedRuntimeFiles) {
            $sourcePath = Join-Path $projectRoot $relativePath
            $destinationPath = Join-Path $fixtureRoot $relativePath
            [void] [System.IO.Directory]::CreateDirectory((Split-Path -Parent $destinationPath))
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath
        }
        $fixtureScriptPath = Join-Path $fixtureRoot 'scripts\New-SqlUtilityPackage.ps1'
        [void] [System.IO.Directory]::CreateDirectory((Split-Path -Parent $fixtureScriptPath))
        Copy-Item -LiteralPath $packageScript -Destination $fixtureScriptPath

        $packagePath = Join-Path $packageOutputFolder 'SQL-Utility-test.zip'
        & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File $fixtureScriptPath -DestinationPath $packagePath
        Assert-Equal 0 $LASTEXITCODE 'Packaging script exits successfully'
        Assert-True (Test-Path -LiteralPath $packagePath -PathType Leaf) `
            'Packaging script creates the requested ZIP file'
        Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot 'SqlUtility.cat') -PathType Leaf) `
            'Packaging script refreshes the tracked catalog when it is missing'

        $catalogHashBeforeRepeat = (Get-FileHash `
            -LiteralPath (Join-Path $fixtureRoot 'SqlUtility.cat') -Algorithm SHA256).Hash
        $repeatPackagePath = Join-Path $packageOutputFolder 'SQL-Utility-repeat.zip'
        & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File $fixtureScriptPath -DestinationPath $repeatPackagePath
        Assert-Equal 0 $LASTEXITCODE 'Repeated packaging exits successfully'
        $catalogHashAfterRepeat = (Get-FileHash `
            -LiteralPath (Join-Path $fixtureRoot 'SqlUtility.cat') -Algorithm SHA256).Hash
        Assert-Equal $catalogHashBeforeRepeat $catalogHashAfterRepeat `
            'Packaging preserves a catalog that still matches the runtime files'

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($packagePath, $extractedFolder)
        $expectedEntries = @($protectedRuntimeFiles + 'SqlUtility.cat') |
            ForEach-Object { $_ -replace '\\', '/' } |
            Sort-Object
        $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
        try {
            $actualEntries = @(
                $archive.Entries |
                    Where-Object { -not [string]::IsNullOrEmpty($_.Name) } |
                    ForEach-Object { $_.FullName -replace '\\', '/' } |
                    Sort-Object
            )
        }
        finally {
            $archive.Dispose()
        }
        Assert-Equal ($expectedEntries -join "`n") ($actualEntries -join "`n") `
            'Package contains exactly the protected runtime files and catalog'

        $extractedRuntimePaths = @(
            $protectedRuntimeFiles | ForEach-Object { Join-Path $extractedFolder $_ }
        )
        $extractedCatalogPath = Join-Path $extractedFolder 'SqlUtility.cat'
        Assert-Equal 'Valid' ([string] (Test-FileCatalog -Path $extractedRuntimePaths `
            -CatalogFilePath $extractedCatalogPath)) `
            'Extracted package matches its catalog'

        Add-Content -LiteralPath (Join-Path $extractedFolder 'modules\SqlUtility.Database.ps1') `
            -Value '# accidental change' -Encoding UTF8
        Assert-Equal 'ValidationFailed' ([string] (Test-FileCatalog `
            -Path $extractedRuntimePaths -CatalogFilePath $extractedCatalogPath)) `
            'Extracted catalog detects a changed runtime module'
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        $resolvedTemporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
        $resolvedSystemTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if ($resolvedTemporaryRoot.StartsWith($resolvedSystemTemp, [System.StringComparison]::OrdinalIgnoreCase) -and
            (Split-Path -Leaf $resolvedTemporaryRoot) -like 'SqlUtility.PackagingTest.*') {
            [System.IO.Directory]::Delete($resolvedTemporaryRoot, $true)
        }
    }
}

Complete-TestFile 'Packaging tests passed.'
