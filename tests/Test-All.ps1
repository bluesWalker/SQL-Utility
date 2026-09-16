$tests = @(
    'Test-Config.ps1', 'Test-Templates.ps1', 'Test-QueryPolicy.ps1', 'Test-DataExplorer.ps1', 'Test-Database.ps1',
    'Test-Excel.ps1', 'Test-SqlUtilityUi.ps1', 'Test-Launcher.ps1', 'Test-Packaging.ps1'
)
foreach ($test in $tests) {
    $arguments = @('-NoLogo','-NoProfile')
    if ($test -in @('Test-SqlUtilityUi.ps1','Test-Launcher.ps1')) { $arguments += '-STA' }
    $arguments += @('-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot $test))
    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
Write-Host 'All SQL Utility tests passed.'
