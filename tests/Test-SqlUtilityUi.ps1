$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'SqlUtility.ps1') -NoGui

function Get-TestControl($Root, [string] $Name) {
    $matches = @($Root.Controls.Find($Name, $true))
    Assert-Equal 1 $matches.Count "Form contains exactly one $Name control"
    if ($matches.Count -eq 0) { return $null }
    return $matches[0]
}

function New-TestServices {
    $recorder = [pscustomobject]@{
        TestCalls = [System.Collections.Generic.List[object]]::new()
        WriteCalls = [System.Collections.Generic.List[object]]::new()
        Messages = [System.Collections.Generic.List[object]]::new()
        ConfirmCalls = [System.Collections.Generic.List[object]]::new()
        DialogCalls = 0
        TestError = $null
        WriteError = $null
        ConfirmResult = $true
        WriteToDisk = $false
    }

    $services = @{
        TestConnection = {
            param($Server, $Database)
            [void] $recorder.TestCalls.Add([pscustomobject]@{ Server = $Server; Database = $Database })
            if ($null -ne $recorder.TestError) {
                throw [System.InvalidOperationException]::new([string] $recorder.TestError)
            }
        }.GetNewClosure()
        WriteConfig = {
            param($Path, $Config)
            [void] $recorder.WriteCalls.Add([pscustomobject]@{ Path = $Path; Config = $Config })
            if ($null -ne $recorder.WriteError) {
                throw [System.IO.IOException]::new([string] $recorder.WriteError)
            }
            if ($recorder.WriteToDisk) {
                return Write-SqlUtilityConfig -Path $Path -Config $Config
            }
            return ConvertTo-SqlUtilityValidatedConfig -InputObject $Config
        }.GetNewClosure()
        ShowMessage = {
            param($Text, $Caption, $Icon)
            [void] $recorder.Messages.Add([pscustomobject]@{ Text = $Text; Caption = $Caption; Icon = [string] $Icon })
        }.GetNewClosure()
        Confirm = {
            param($Text, $Caption)
            [void] $recorder.ConfirmCalls.Add([pscustomobject]@{ Text = $Text; Caption = $Caption })
            return [bool] $recorder.ConfirmResult
        }.GetNewClosure()
        ShowDialog = {
            param($Form)
            $recorder.DialogCalls++
        }.GetNewClosure()
    }

    return [pscustomobject]@{ Recorder = $recorder; Services = $services }
}

function New-TestConfig {
    param([switch] $WithConnections)
    $config = New-SqlUtilityDefaultConfig
    if ($WithConnections) {
        $config = Add-SqlUtilitySavedConnection -Config $config -Server 'SavedServer' -Database 'SavedDatabase'
        $config = Add-SqlUtilitySavedConnection -Config $config -Server 'SecondServer' -Database 'SecondDatabase'
    }
    return $config
}

function Show-TestForm($Form) {
    $Form.Show()
    [System.Windows.Forms.Application]::DoEvents()
}

$requiredControlNames = @(
    'ConnectionPanel', 'ServerTextBox', 'DatabaseTextBox', 'TestConnectionButton',
    'ConnectButton', 'SavedConnectionsList', 'DeleteConnectionButton',
    'WorkspacePanel', 'ActiveConnectionLabel', 'ChangeConnectionButton',
    'WorkspaceTabs', 'QueryTab', 'SettingsTab', 'UnorderedLimitNumeric',
    'QueryExportTimeoutNumeric', 'SaveSettingsButton', 'MainStatusLabel'
)

# Initial stage, stable control contract, saved-pair selection, and settings bounds.
$initialHarness = New-TestServices
$initialForm = New-SqlUtilityMainForm -Config (New-TestConfig -WithConnections) -ConfigPath 'C:\test\config.json' -Services $initialHarness.Services
try {
    foreach ($controlName in $requiredControlNames) {
        $null = Get-TestControl $initialForm $controlName
    }
    Show-TestForm $initialForm
    $connectionPanel = Get-TestControl $initialForm 'ConnectionPanel'
    $workspacePanel = Get-TestControl $initialForm 'WorkspacePanel'
    $serverTextBox = Get-TestControl $initialForm 'ServerTextBox'
    $databaseTextBox = Get-TestControl $initialForm 'DatabaseTextBox'
    $savedList = Get-TestControl $initialForm 'SavedConnectionsList'
    $unorderedNumeric = Get-TestControl $initialForm 'UnorderedLimitNumeric'
    $timeoutNumeric = Get-TestControl $initialForm 'QueryExportTimeoutNumeric'
    $settingsTab = Get-TestControl $initialForm 'SettingsTab'

    Assert-Equal $true $connectionPanel.Visible 'Connection stage is visible first'
    Assert-Equal $false $workspacePanel.Visible 'Workspace stage is hidden first'
    Assert-Equal '' $serverTextBox.Text 'Server starts blank on every form creation'
    Assert-Equal '' $databaseTextBox.Text 'Database starts blank on every form creation'
    Assert-Equal 2 $savedList.Items.Count 'Saved pairs populate the list'
    Assert-Equal ("SavedServer $([char]0x2013) SavedDatabase") $savedList.GetItemText($savedList.Items[0]) 'Saved pair uses server en dash database display'

    $savedList.SelectedIndex = 0
    [System.Windows.Forms.Application]::DoEvents()
    Assert-Equal 'SavedServer' $serverTextBox.Text 'Selecting a saved pair fills server only after selection'
    Assert-Equal 'SavedDatabase' $databaseTextBox.Text 'Selecting a saved pair fills database only after selection'

    Assert-Equal 100 ([int] $unorderedNumeric.Minimum) 'Maximum unordered rows minimum is exact'
    Assert-Equal 2000 ([int] $unorderedNumeric.Maximum) 'Maximum unordered rows maximum is exact'
    Assert-Equal 1000 ([int] $unorderedNumeric.Value) 'Maximum unordered rows default is exact'
    Assert-Equal 5 ([int] $timeoutNumeric.Minimum) 'Query/Export timeout minimum is exact'
    Assert-Equal 3600 ([int] $timeoutNumeric.Maximum) 'Query/Export timeout maximum is exact'
    Assert-Equal 120 ([int] $timeoutNumeric.Value) 'Query/Export timeout default is exact'
    Assert-True ($settingsTab.Text -eq 'Settings') 'Workspace exposes the Settings tab label'
    $settingsText = ($settingsTab.Controls | ForEach-Object { $_.Text }) -join ' '
    Assert-True ($settingsText -match 'Maximum unordered rows') 'Settings labels the unordered-row limit'
    Assert-True ($settingsText -match 'Query/Export timeout \(seconds\)') 'Settings labels the shared timeout'
    Assert-True ($settingsText -match 'interactive queries') 'Settings help covers interactive queries'
    Assert-True ($settingsText -match 'complete Excel export') 'Settings help covers complete Excel export'
    Assert-True ($settingsText -match 'connection timeout remains fixed and separate') 'Settings help separates connection timeout'
}
finally {
    $initialForm.Close()
    $initialForm.Dispose()
}

# Test Connection trims values, tests once, saves, reports success, and always restores busy state.
$testHarness = New-TestServices
$testForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $testHarness.Services
try {
    Show-TestForm $testForm
    $serverTextBox = Get-TestControl $testForm 'ServerTextBox'
    $databaseTextBox = Get-TestControl $testForm 'DatabaseTextBox'
    $testButton = Get-TestControl $testForm 'TestConnectionButton'
    $savedList = Get-TestControl $testForm 'SavedConnectionsList'
    $serverTextBox.Text = '  ServerA  '
    $databaseTextBox.Text = '  DatabaseA  '
    $testButton.PerformClick()

    Assert-Equal 1 $testHarness.Recorder.TestCalls.Count 'Test Connection invokes SQL test once'
    Assert-Equal 'ServerA' $testHarness.Recorder.TestCalls[0].Server 'Test Connection trims server before SQL'
    Assert-Equal 'DatabaseA' $testHarness.Recorder.TestCalls[0].Database 'Test Connection trims database before SQL'
    Assert-Equal 1 $testHarness.Recorder.WriteCalls.Count 'Test Connection persists after SQL success'
    Assert-Equal 1 $testForm.Tag.Config.connections.Count 'Test Connection enters successful pair into in-memory config'
    Assert-Equal 1 $savedList.Items.Count 'Test Connection refreshes saved list'
    Assert-True (@($testHarness.Recorder.Messages | Where-Object Icon -eq 'Information').Count -eq 1) 'Test Connection shows success information'
    Assert-Equal $false $testForm.Tag.IsBusy 'Test Connection restores state after success'
    Assert-Equal $true $testForm.Enabled 'Test Connection restores controls after success'

    $testHarness.Recorder.TestError = 'SQL unavailable'
    $serverTextBox.Text = 'FailureServer'
    $databaseTextBox.Text = 'FailureDatabase'
    $testButton.PerformClick()
    Assert-Equal 2 $testHarness.Recorder.TestCalls.Count 'Connection failure still invokes SQL once for that attempt'
    Assert-Equal 1 $testHarness.Recorder.WriteCalls.Count 'Connection failure never writes config'
    Assert-Equal 1 $testForm.Tag.Config.connections.Count 'Connection failure does not add a pair'
    Assert-Equal $false $testForm.Tag.IsBusy 'Test Connection restores busy state after fake exception'
    Assert-Equal $true $testForm.Enabled 'Test Connection restores controls after fake exception'
}
finally {
    $testForm.Close()
    $testForm.Dispose()
}

# A save failure after SQL success warns but retains the successful pair in memory and on screen.
$saveFailureHarness = New-TestServices
$saveFailureHarness.Recorder.WriteError = 'disk full'
$saveFailureForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $saveFailureHarness.Services
try {
    Show-TestForm $saveFailureForm
    (Get-TestControl $saveFailureForm 'ServerTextBox').Text = 'TransientServer'
    (Get-TestControl $saveFailureForm 'DatabaseTextBox').Text = 'TransientDatabase'
    (Get-TestControl $saveFailureForm 'TestConnectionButton').PerformClick()
    Assert-Equal 1 $saveFailureForm.Tag.Config.connections.Count 'Persistence warning keeps successful pair in memory'
    Assert-Equal 1 (Get-TestControl $saveFailureForm 'SavedConnectionsList').Items.Count 'Persistence warning keeps successful pair in list'
    Assert-True (@($saveFailureHarness.Recorder.Messages | Where-Object Icon -eq 'Warning').Count -eq 1) 'Persistence failure is a warning'
    Assert-Equal $false $saveFailureForm.Tag.IsBusy 'Persistence exception restores busy state'
}
finally {
    $saveFailureForm.Close()
    $saveFailureForm.Dispose()
}

# Connect repeats test/save and enters Query; failed SQL does neither save nor transition.
$connectHarness = New-TestServices
$connectForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $connectHarness.Services
try {
    Show-TestForm $connectForm
    (Get-TestControl $connectForm 'ServerTextBox').Text = ' ConnectServer '
    (Get-TestControl $connectForm 'DatabaseTextBox').Text = ' ConnectDatabase '
    (Get-TestControl $connectForm 'ConnectButton').PerformClick()
    Assert-Equal 1 $connectHarness.Recorder.TestCalls.Count 'Connect repeats the SQL test'
    Assert-Equal 1 $connectHarness.Recorder.WriteCalls.Count 'Connect repeats persistence'
    Assert-Equal 'ConnectServer' $connectForm.Tag.ActiveServer 'Connect enters trimmed active server'
    Assert-Equal 'ConnectDatabase' $connectForm.Tag.ActiveDatabase 'Connect enters trimmed active database'
    Assert-Equal $true (Get-TestControl $connectForm 'WorkspacePanel').Visible 'Connect enters workspace'
    Assert-Equal $false (Get-TestControl $connectForm 'ConnectionPanel').Visible 'Connect hides connection stage'
    Assert-Equal (Get-TestControl $connectForm 'QueryTab') (Get-TestControl $connectForm 'WorkspaceTabs').SelectedTab 'Connect selects Query tab'
    Assert-Equal ("ConnectServer $([char]0x2013) ConnectDatabase") (Get-TestControl $connectForm 'ActiveConnectionLabel').Text 'Workspace identifies active pair'
}
finally {
    $connectForm.Close()
    $connectForm.Dispose()
}

$connectFailureHarness = New-TestServices
$connectFailureHarness.Recorder.TestError = 'login failed'
$connectFailureForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $connectFailureHarness.Services
try {
    Show-TestForm $connectFailureForm
    (Get-TestControl $connectFailureForm 'ServerTextBox').Text = 'BadServer'
    (Get-TestControl $connectFailureForm 'DatabaseTextBox').Text = 'BadDatabase'
    (Get-TestControl $connectFailureForm 'ConnectButton').PerformClick()
    Assert-Equal 0 $connectFailureHarness.Recorder.WriteCalls.Count 'Failed Connect never writes config'
    Assert-Equal 0 $connectFailureForm.Tag.Config.connections.Count 'Failed Connect never adds saved pair'
    Assert-Equal $true (Get-TestControl $connectFailureForm 'ConnectionPanel').Visible 'Failed Connect remains on connection stage'
    Assert-Equal $false $connectFailureForm.Tag.IsBusy 'Failed Connect restores busy state'
}
finally {
    $connectFailureForm.Close()
    $connectFailureForm.Dispose()
}

# Delete requires confirmation, writes first, and swaps state/list only on success.
$deleteHarness = New-TestServices
$deleteForm = New-SqlUtilityMainForm -Config (New-TestConfig -WithConnections) -ConfigPath 'C:\test\config.json' -Services $deleteHarness.Services
try {
    Show-TestForm $deleteForm
    $savedList = Get-TestControl $deleteForm 'SavedConnectionsList'
    $deleteButton = Get-TestControl $deleteForm 'DeleteConnectionButton'
    $savedList.SelectedIndex = 0
    $deleteHarness.Recorder.ConfirmResult = $false
    $deleteButton.PerformClick()
    Assert-Equal 1 $deleteHarness.Recorder.ConfirmCalls.Count 'Delete asks for confirmation'
    Assert-Equal 0 $deleteHarness.Recorder.WriteCalls.Count 'Declined delete never writes'
    Assert-Equal 2 $savedList.Items.Count 'Declined delete keeps list'

    $deleteHarness.Recorder.ConfirmResult = $true
    $deleteHarness.Recorder.WriteError = 'write denied'
    $deleteButton.PerformClick()
    Assert-Equal 1 $deleteHarness.Recorder.WriteCalls.Count 'Confirmed delete writes candidate once'
    Assert-Equal 2 $deleteForm.Tag.Config.connections.Count 'Failed delete write keeps config state'
    Assert-Equal 2 $savedList.Items.Count 'Failed delete write keeps list state'
    Assert-Equal $false $deleteForm.Tag.IsBusy 'Delete exception restores busy state'

    $deleteHarness.Recorder.WriteError = $null
    $deleteButton.PerformClick()
    Assert-Equal 1 $deleteForm.Tag.Config.connections.Count 'Successful delete swaps config state'
    Assert-Equal 1 $savedList.Items.Count 'Successful delete refreshes list after write'
}
finally {
    $deleteForm.Close()
    $deleteForm.Dispose()
}

# Settings enter application state only after a successful write.
$settingsHarness = New-TestServices
$settingsForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $settingsHarness.Services
try {
    Show-TestForm $settingsForm
    $unorderedNumeric = Get-TestControl $settingsForm 'UnorderedLimitNumeric'
    $timeoutNumeric = Get-TestControl $settingsForm 'QueryExportTimeoutNumeric'
    $saveSettings = Get-TestControl $settingsForm 'SaveSettingsButton'
    Set-SqlUtilityStage -Form $settingsForm -Stage 'Workspace'
    (Get-TestControl $settingsForm 'WorkspaceTabs').SelectedTab = Get-TestControl $settingsForm 'SettingsTab'
    [System.Windows.Forms.Application]::DoEvents()
    $unorderedNumeric.Value = 1500
    $timeoutNumeric.Value = 300
    $settingsHarness.Recorder.WriteError = 'read-only directory'
    $saveSettings.PerformClick()
    Assert-Equal 1000 $settingsForm.Tag.Config.unorderedRowLimit 'Failed settings write keeps active row limit'
    Assert-Equal 120 $settingsForm.Tag.Config.queryExportTimeoutSeconds 'Failed settings write keeps active timeout'
    Assert-Equal $false $settingsForm.Tag.IsBusy 'Settings exception restores busy state'

    $settingsHarness.Recorder.WriteError = $null
    $saveSettings.PerformClick()
    Assert-Equal 1500 $settingsForm.Tag.Config.unorderedRowLimit 'Successful settings write enters row limit state'
    Assert-Equal 300 $settingsForm.Tag.Config.queryExportTimeoutSeconds 'Successful settings write enters timeout state'
}
finally {
    $settingsForm.Close()
    $settingsForm.Dispose()
}

# Change Connection cancellation preserves state; confirmation clears transient query/result state and inputs.
$changeHarness = New-TestServices
$changeForm = New-SqlUtilityMainForm -Config (New-TestConfig -WithConnections) -ConfigPath 'C:\test\config.json' -Services $changeHarness.Services
try {
    Show-TestForm $changeForm
    (Get-TestControl $changeForm 'ServerTextBox').Text = 'ActiveServer'
    (Get-TestControl $changeForm 'DatabaseTextBox').Text = 'ActiveDatabase'
    (Get-TestControl $changeForm 'ConnectButton').PerformClick()
    $sentinelResult = [pscustomobject]@{ Marker = 'result' }
    $changeForm.Tag.ExecutedQuery = [pscustomobject]@{ OriginalEditorSql = 'SELECT 1'; NormalizedSql = 'SELECT 1' }
    $changeForm.Tag.CurrentResult = $sentinelResult
    $changeForm.Tag.CurrentPage = 4
    $changeForm.Tag.IsQueryStale = $true
    $changeHarness.Recorder.ConfirmResult = $false
    (Get-TestControl $changeForm 'ChangeConnectionButton').PerformClick()
    Assert-Equal 'ActiveServer' $changeForm.Tag.ActiveServer 'Canceled change preserves active server'
    Assert-Equal $sentinelResult $changeForm.Tag.CurrentResult 'Canceled change preserves result state'
    Assert-Equal 4 $changeForm.Tag.CurrentPage 'Canceled change preserves page state'
    Assert-Equal $true (Get-TestControl $changeForm 'WorkspacePanel').Visible 'Canceled change preserves workspace stage'

    $changeHarness.Recorder.ConfirmResult = $true
    (Get-TestControl $changeForm 'ChangeConnectionButton').PerformClick()
    Assert-Equal '' $changeForm.Tag.ActiveServer 'Confirmed change clears active server'
    Assert-Equal '' $changeForm.Tag.ActiveDatabase 'Confirmed change clears active database'
    Assert-Equal $null $changeForm.Tag.ExecutedQuery 'Confirmed change clears executed query state'
    Assert-Equal $null $changeForm.Tag.CurrentResult 'Confirmed change clears result state'
    Assert-Equal 0 $changeForm.Tag.CurrentPage 'Confirmed change clears page state'
    Assert-Equal $false $changeForm.Tag.IsQueryStale 'Confirmed change clears stale state'
    Assert-Equal '' (Get-TestControl $changeForm 'ServerTextBox').Text 'Confirmed change blanks server input'
    Assert-Equal '' (Get-TestControl $changeForm 'DatabaseTextBox').Text 'Confirmed change blanks database input'
    Assert-Equal 3 (Get-TestControl $changeForm 'SavedConnectionsList').Items.Count 'Confirmed change retains saved pairs including the successful active pair'
    Assert-Equal 1000 ([int] (Get-TestControl $changeForm 'UnorderedLimitNumeric').Value) 'Confirmed change retains settings'
    Assert-Equal $true (Get-TestControl $changeForm 'ConnectionPanel').Visible 'Confirmed change returns to connection stage'
}
finally {
    $changeForm.Close()
    $changeForm.Dispose()
}

# Public busy helper prevents re-entry and restores the form/status.
$busyHarness = New-TestServices
$busyForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $busyHarness.Services
try {
    Show-TestForm $busyForm
    Set-SqlUtilityBusy -Form $busyForm -Busy $true -Message 'Working...'
    Assert-Equal $true $busyForm.Tag.IsBusy 'Busy helper records busy state'
    Assert-Equal $false $busyForm.Enabled 'Busy helper disables re-entry'
    Assert-Equal 'Working...' (Get-TestControl $busyForm 'MainStatusLabel').Text 'Busy helper updates status'
    Set-SqlUtilityBusy -Form $busyForm -Busy $false -Message 'Ready.'
    Assert-Equal $false $busyForm.Tag.IsBusy 'Busy helper clears busy state'
    Assert-Equal $true $busyForm.Enabled 'Busy helper restores form'
    Assert-Equal 'Ready.' (Get-TestControl $busyForm 'MainStatusLabel').Text 'Busy helper restores status'
}
finally {
    $busyForm.Close()
    $busyForm.Dispose()
}

# Startup recovery is owned by Start-SqlUtilityApplication and never happens under -NoGui.
$startupRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SqlUtilityUiTests-' + [guid]::NewGuid().ToString('N'))
[void] [System.IO.Directory]::CreateDirectory($startupRoot)
try {
    $malformedPath = Join-Path $startupRoot 'malformed.json'
    $malformedBytes = [System.Text.UTF8Encoding]::new($false).GetBytes('{malformed')
    [System.IO.File]::WriteAllBytes($malformedPath, $malformedBytes)
    $declineHarness = New-TestServices
    $declineHarness.Recorder.ConfirmResult = $false
    Start-SqlUtilityApplication -ConfigPath $malformedPath -Services $declineHarness.Services
    Assert-Equal 1 $declineHarness.Recorder.ConfirmCalls.Count 'Malformed startup asks once before reset'
    Assert-Equal 0 $declineHarness.Recorder.WriteCalls.Count 'Declined reset does not write'
    Assert-Equal 0 $declineHarness.Recorder.DialogCalls 'Declined reset exits without showing form'
    Assert-Equal ($malformedBytes -join ',') ([System.IO.File]::ReadAllBytes($malformedPath) -join ',') 'Declined reset preserves bytes'

    [System.IO.File]::WriteAllBytes($malformedPath, $malformedBytes)
    $acceptHarness = New-TestServices
    $acceptHarness.Recorder.ConfirmResult = $true
    $acceptHarness.Recorder.WriteToDisk = $true
    Start-SqlUtilityApplication -ConfigPath $malformedPath -Services $acceptHarness.Services
    $recovered = Read-SqlUtilityConfig -Path $malformedPath
    Assert-Equal 1 $acceptHarness.Recorder.ConfirmCalls.Count 'Malformed startup confirms reset once'
    Assert-Equal 1 $acceptHarness.Recorder.WriteCalls.Count 'Confirmed reset writes defaults once'
    Assert-Equal 1 $acceptHarness.Recorder.DialogCalls 'Confirmed recovery continues to the form'
    Assert-Equal 1000 $recovered.unorderedRowLimit 'Confirmed reset writes default configuration'

    $futurePath = Join-Path $startupRoot 'future.json'
    $futureText = '{"schemaVersion":2,"unorderedRowLimit":1000,"queryExportTimeoutSeconds":120,"connections":[]}'
    $futureBytes = [System.Text.UTF8Encoding]::new($false).GetBytes($futureText)
    [System.IO.File]::WriteAllBytes($futurePath, $futureBytes)
    $futureHarness = New-TestServices
    Start-SqlUtilityApplication -ConfigPath $futurePath -Services $futureHarness.Services
    Assert-Equal 0 $futureHarness.Recorder.ConfirmCalls.Count 'Unsupported schema never offers reset'
    Assert-Equal 0 $futureHarness.Recorder.WriteCalls.Count 'Unsupported schema never writes'
    Assert-Equal 0 $futureHarness.Recorder.DialogCalls 'Unsupported schema exits without showing form'
    Assert-Equal 1 @($futureHarness.Recorder.Messages | Where-Object Icon -eq 'Error').Count 'Unsupported schema shows one error'
    Assert-Equal ($futureBytes -join ',') ([System.IO.File]::ReadAllBytes($futurePath) -join ',') 'Unsupported schema preserves bytes'
}
finally {
    if (Test-Path -LiteralPath $startupRoot) {
        Remove-Item -LiteralPath $startupRoot -Recurse -Force
    }
}

Complete-TestFile 'All SQL Utility UI tests passed.'
