$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.Config.ps1')

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SqlUtility.TemplatesTest.' + [guid]::NewGuid().ToString('N'))
[void] [System.IO.Directory]::CreateDirectory($testRoot)
try {
    $templates = Join-Path $testRoot 'Templates'
    Initialize-SqlUtilityTemplateDirectory -Path $templates
    Assert-True ([System.IO.Directory]::Exists($templates)) 'Template folder is created by default'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $templates -Force).Count 'New template folder is empty'

    $path = Join-Path $templates 'lookup.sql'
    $text = "-- $([char]0x4e2d)$([char]0x6587) $([char]0x00e9)`r`nSELECT [Name]`r`nFROM [dbo].[Items]`r`nWHERE [Name] = N'O''Brien' AND [Id] = {{Id}};`r`n  "
    $saved = Write-SqlUtilityTemplate -Path $path -Text $text
    Assert-Equal '' $saved.CleanupWarning 'New save has no cleanup warning'
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $path) 'Values, placeholders, Unicode, and whitespace round-trip exactly'
    $bytes = [System.IO.File]::ReadAllBytes($path)
    Assert-Equal '239,187,191' ($bytes[0..2] -join ',') 'Template saves use a UTF-8 BOM'
    Initialize-SqlUtilityTemplateDirectory -Path $templates
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $path) 'Initializing an existing folder preserves personal templates'

    Assert-Throws { Write-SqlUtilityTemplate -Path $path -Text 'replacement' } $null 'Existing file cannot be replaced without overwrite permission'
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $path) 'Unapproved overwrite preserves the original'
    $replacement = "SELECT Id FROM dbo.Other;`n"
    $saved = Write-SqlUtilityTemplate -Path $path -Text $replacement -AllowOverwrite $true
    Assert-Equal $replacement (Read-SqlUtilityTemplate -Path $path) 'Approved replacement saves exact text'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $templates -Filter '.SqlUtility.template.*' -Force).Count 'Successful saves remove temporary and backup files'

    $outside = Join-Path $testRoot 'outside.sql'
    [void] (Write-SqlUtilityTemplate -Path $outside -Text $text)
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $outside) 'User-selected paths outside Templates are supported'
    Assert-Throws { Write-SqlUtilityTemplate -Path (Join-Path $testRoot 'bad.ps1') -Text $text } $null 'Save rejects a non-SQL extension'
    Assert-Throws { Read-SqlUtilityTemplate -Path (Join-Path $testRoot 'bad.ps1') } $null 'Load rejects a non-SQL extension'
    Assert-Throws { Write-SqlUtilityTemplate -Path $path -Text '  ' -AllowOverwrite $true } $null 'Whitespace cannot replace a saved template'
    Assert-Equal $replacement (Read-SqlUtilityTemplate -Path $path) 'Blank save preserves existing template'

    $utf16 = Join-Path $testRoot 'utf16.sql'
    [System.IO.File]::WriteAllText($utf16, $text, [System.Text.Encoding]::Unicode)
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $utf16) 'BOM-marked UTF-16 SQL files load correctly'
    $utf8 = Join-Path $testRoot 'utf8.sql'
    [System.IO.File]::WriteAllText($utf8, $text, [System.Text.UTF8Encoding]::new($false))
    Assert-Equal $text (Read-SqlUtilityTemplate -Path $utf8) 'UTF-8 without BOM loads correctly'
    $invalid = Join-Path $testRoot 'invalid.sql'
    [System.IO.File]::WriteAllBytes($invalid, [byte[]] @(0xc3, 0x28))
    Assert-Throws { Read-SqlUtilityTemplate -Path $invalid } $null 'Invalid UTF-8 fails instead of silently replacing characters'
    [System.IO.File]::WriteAllText($invalid, "SELECT`0hidden", [System.Text.Encoding]::UTF8)
    Assert-Throws { Read-SqlUtilityTemplate -Path $invalid } $null 'Embedded NUL is rejected instead of truncating the editor text'
    Assert-Throws { Read-SqlUtilityTemplate -Path (Join-Path $testRoot 'missing.sql') } $null 'Missing file is reported as a load failure'

    $beforeFailure = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($path))
    $locked = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        Assert-Throws { Read-SqlUtilityTemplate -Path $path } $null 'Locked files cannot be loaded'
        Assert-Throws { Write-SqlUtilityTemplate -Path $path -Text $text -AllowOverwrite $true } $null 'Locked destination fails safely'
    }
    finally { $locked.Dispose() }
    Assert-Equal $beforeFailure ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($path))) 'Pre-commit failure preserves destination bytes'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $templates -Filter '.SqlUtility.template.*' -Force).Count 'Pre-commit failure cleans temporary siblings'

    # Exercise post-commit cleanup using the existing module's narrow filesystem cleanup boundary.
    $originalCleanup = (Get-Command Remove-SqlUtilityTemplateTransientFile).ScriptBlock
    try {
        Set-Item -Path Function:\Remove-SqlUtilityTemplateTransientFile -Value {
            param([string] $Path)
            if ([System.IO.Path]::GetExtension($Path) -eq '.bak') { throw 'backup is locked' }
            [System.IO.File]::Delete($Path)
        }
        $saved = Write-SqlUtilityTemplate -Path $path -Text $text -AllowOverwrite $true
        Assert-Equal $text (Read-SqlUtilityTemplate -Path $path) 'Cleanup failure does not undo a committed template'
        Assert-True ($saved.CleanupWarning -match 'saved') 'Committed save reports cleanup separately'
        $backups = @(Get-ChildItem -LiteralPath $templates -Filter '.SqlUtility.template.*.bak' -Force)
        Assert-Equal 1 $backups.Count 'Cleanup failure retains one recoverable backup'
        Assert-Equal $replacement ([System.IO.File]::ReadAllText($backups[0].FullName)) 'Recoverable backup contains previous text'
    }
    finally { Set-Item -Path Function:\Remove-SqlUtilityTemplateTransientFile -Value $originalCleanup }
}
finally {
    $resolvedRoot = [System.IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedRoot.StartsWith($tempPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedRoot) -like 'SqlUtility.TemplatesTest.*') {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
Complete-TestFile 'Template persistence tests passed.'
