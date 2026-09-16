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

function Invoke-TestGit {
    param([Parameter(Mandatory = $true)][string[]] $Arguments)

    $priorErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & git.exe @Arguments 2>&1
    }
    finally {
        $ErrorActionPreference = $priorErrorActionPreference
    }
    if ($LASTEXITCODE -ne 0) {
        throw ('Git command failed: git {0}{1}{2}' -f
            ($Arguments -join ' '), [Environment]::NewLine, ($output -join [Environment]::NewLine))
    }
}

$packageScript = Join-Path $projectRoot 'scripts\New-SqlUtilityPackage.ps1'
Assert-True (Test-Path -LiteralPath $packageScript -PathType Leaf) `
    'Packaging script exists'

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('SqlUtility.PackagingTest.{0}' -f [guid]::NewGuid().ToString('N'))
$fixtureRoot = Join-Path $temporaryRoot 'Project'
$packageOutputFolder = Join-Path $temporaryRoot 'Output'
$extractedFolder = Join-Path $temporaryRoot 'Extracted'
$checkoutFixtureRoot = Join-Path $temporaryRoot 'Checkout Source'
[void] [System.IO.Directory]::CreateDirectory($fixtureRoot)
[void] [System.IO.Directory]::CreateDirectory($packageOutputFolder)

try {
    $attributesPath = Join-Path $projectRoot '.gitattributes'
    $gitCommand = Get-Command git.exe -ErrorAction SilentlyContinue
    Assert-True ($null -ne $gitCommand) 'Git is available for clean-checkout validation'
    Assert-True (Test-Path -LiteralPath $attributesPath -PathType Leaf) `
        'Protected runtime files have a source-control checkout policy'
    if ($null -ne $gitCommand -and
        (Test-Path -LiteralPath $attributesPath -PathType Leaf)) {
        [void] [System.IO.Directory]::CreateDirectory($checkoutFixtureRoot)
        Copy-Item -LiteralPath $attributesPath -Destination $checkoutFixtureRoot
        Copy-Item -LiteralPath (Join-Path $projectRoot 'SqlUtility.cat') `
            -Destination $checkoutFixtureRoot
        foreach ($relativePath in $protectedRuntimeFiles) {
            $sourcePath = Join-Path $projectRoot $relativePath
            $destinationPath = Join-Path $checkoutFixtureRoot $relativePath
            [void] [System.IO.Directory]::CreateDirectory((Split-Path -Parent $destinationPath))
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath
        }

        Invoke-TestGit -Arguments @('init', '--quiet', $checkoutFixtureRoot)
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'config', 'user.name', 'SQL Utility Tests')
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'config', 'user.email', 'tests@localhost')
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'config', 'core.autocrlf', 'false')
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'config', 'core.safecrlf', 'false')
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'add', '--', '.')
        Invoke-TestGit -Arguments @('-C', $checkoutFixtureRoot, 'commit', '--quiet', '-m', 'checkout fixture')
        $fixtureCommit = (& git.exe -C $checkoutFixtureRoot rev-parse HEAD).Trim()
        Assert-Equal 0 $LASTEXITCODE 'Checkout fixture commit can be resolved'

        foreach ($autoCrlf in @('true', 'false')) {
            $checkoutRoot = Join-Path $temporaryRoot ('Checkout ' + $autoCrlf)
            Invoke-TestGit -Arguments @('clone', '--quiet', '--no-hardlinks', '--no-checkout',
                $checkoutFixtureRoot, $checkoutRoot)
            Invoke-TestGit -Arguments @('-C', $checkoutRoot, 'config', 'core.autocrlf', $autoCrlf)
            Invoke-TestGit -Arguments @('-C', $checkoutRoot, 'checkout', '--quiet',
                '--detach', $fixtureCommit)
            $checkoutRuntimePaths = @(
                $protectedRuntimeFiles | ForEach-Object { Join-Path $checkoutRoot $_ }
            )
            Assert-Equal 'Valid' ([string] (Test-FileCatalog -Path $checkoutRuntimePaths `
                -CatalogFilePath (Join-Path $checkoutRoot 'SqlUtility.cat'))) `
                "Clean checkout matches the catalog when core.autocrlf=$autoCrlf"
        }
    }

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

        $personalTemplates = Join-Path $fixtureRoot 'Templates'
        [void] [System.IO.Directory]::CreateDirectory($personalTemplates)
        [System.IO.File]::WriteAllText((Join-Path $personalTemplates 'private.sql'), 'SELECT private FROM dbo.Items')

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
        $extractedTemplates = Join-Path $extractedFolder 'Templates'
        Assert-True ([System.IO.Directory]::Exists($extractedTemplates)) 'Package includes a Templates folder'
        if ([System.IO.Directory]::Exists($extractedTemplates)) {
            Assert-Equal 0 @(Get-ChildItem -LiteralPath $extractedTemplates -Force).Count 'Packaged Templates folder is empty'
        }
        Assert-True ([System.IO.File]::Exists((Join-Path $personalTemplates 'private.sql'))) 'Packaging preserves source personal templates'
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
            Get-ChildItem -LiteralPath $resolvedTemporaryRoot -Force -Recurse |
                ForEach-Object { $_.Attributes = [System.IO.FileAttributes]::Normal }
            [System.IO.Directory]::Delete($resolvedTemporaryRoot, $true)
        }
    }
}

Complete-TestFile 'Packaging tests passed.'
