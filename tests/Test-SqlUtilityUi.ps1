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
        ValidationCalls = [System.Collections.Generic.List[object]]::new()
        OrderedCalls = [System.Collections.Generic.List[object]]::new()
        UnorderedCalls = [System.Collections.Generic.List[object]]::new()
        LocalPageCalls = [System.Collections.Generic.List[object]]::new()
        ExportCalls = [System.Collections.Generic.List[object]]::new()
        PromptCalls = 0
        DialogCalls = 0
        TestError = $null
        WriteError = $null
        ExecuteError = $null
        LocalPageError = $null
        ExportError = $null
        PromptError = $null
        ConfirmResult = $true
        WriteToDisk = $false
        ValidationResult = [pscustomobject][ordered]@{
            IsValid = $true
            ErrorMessage = ''
            NormalizedSql = 'SELECT Id FROM dbo.Items ORDER BY Id'
            TableIdentifier = 'dbo.Items'
            HasOrderBy = $true
        }
        OrderedResults = @{}
        UnorderedResult = $null
        PromptPath = $null
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
        ValidateQuery = {
            param($Sql)
            [void] $recorder.ValidationCalls.Add($Sql)
            return $recorder.ValidationResult
        }.GetNewClosure()
        ExecuteOrderedPage = {
            param($Server, $Database, $Sql, $PageNumber, $TimeoutSeconds)
            [void] $recorder.OrderedCalls.Add([pscustomobject]@{
                Server = $Server
                Database = $Database
                Sql = $Sql
                PageNumber = $PageNumber
                TimeoutSeconds = $TimeoutSeconds
            })
            if ($null -ne $recorder.ExecuteError) {
                throw [System.TimeoutException]::new([string] $recorder.ExecuteError)
            }
            return $recorder.OrderedResults[[int] $PageNumber]
        }.GetNewClosure()
        ExecuteUnordered = {
            param($Server, $Database, $Sql, $RowLimit, $TimeoutSeconds)
            [void] $recorder.UnorderedCalls.Add([pscustomobject]@{
                Server = $Server
                Database = $Database
                Sql = $Sql
                RowLimit = $RowLimit
                TimeoutSeconds = $TimeoutSeconds
            })
            if ($null -ne $recorder.ExecuteError) {
                throw [System.InvalidOperationException]::new([string] $recorder.ExecuteError)
            }
            return $recorder.UnorderedResult
        }.GetNewClosure()
        GetLocalPage = {
            param($CachedData, $PageNumber, $IsComplete, $IsTruncated)
            [void] $recorder.LocalPageCalls.Add([pscustomobject]@{
                CachedData = $CachedData
                PageNumber = $PageNumber
                IsComplete = $IsComplete
                IsTruncated = $IsTruncated
            })
            if ($null -ne $recorder.LocalPageError) {
                throw [System.InvalidOperationException]::new([string] $recorder.LocalPageError)
            }
            return Get-SqlUtilityLocalPage -CachedData $CachedData -PageNumber $PageNumber `
                -IsComplete $IsComplete -IsTruncated $IsTruncated
        }.GetNewClosure()
        ExportResult = {
            param($State, $DestinationPath)
            $normalizedSql = $null
            $cachedData = $null
            if ($null -ne $State.ExecutedQuery) {
                $normalizedSql = $State.ExecutedQuery.NormalizedSql
            }
            if ($null -ne $State.CurrentResult) {
                $cachedData = $State.CurrentResult.CachedData
            }
            [void] $recorder.ExportCalls.Add([pscustomobject]@{
                DestinationPath = $DestinationPath
                State = $State
                NormalizedSql = $normalizedSql
                CachedData = $cachedData
            })
            if ($null -ne $recorder.ExportError) {
                throw [System.IO.IOException]::new([string] $recorder.ExportError)
            }
        }.GetNewClosure()
        PromptSavePath = {
            $recorder.PromptCalls++
            if ($null -ne $recorder.PromptError) {
                throw [System.IO.IOException]::new([string] $recorder.PromptError)
            }
            return $recorder.PromptPath
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

function New-TestDataTable {
    param(
        [int] $RowCount,
        [int] $TextLength = 8
    )

    $table = [System.Data.DataTable]::new()
    [void] $table.Columns.Add('Id', [int])
    [void] $table.Columns.Add('Description', [string])
    for ($index = 1; $index -le $RowCount; $index++) {
        $row = $table.NewRow()
        $row.Id = $index
        $row.Description = ('x' * $TextLength) + $index
        [void] $table.Rows.Add($row)
    }
    return (, $table)
}

function New-TestPageResult {
    param(
        [Parameter(Mandatory = $true)][System.Data.DataTable] $Data,
        [AllowNull()][System.Data.DataTable] $CachedData,
        [int] $PageNumber = 1,
        [bool] $HasPrevious = $false,
        [bool] $HasNext = $false,
        [bool] $IsComplete = $false,
        [bool] $IsTruncated = $false
    )

    return [pscustomobject][ordered]@{
        Data = $Data
        CachedData = $CachedData
        PageNumber = $PageNumber
        DisplayedRowCount = $Data.Rows.Count
        HasPrevious = $HasPrevious
        HasNext = $HasNext
        IsComplete = $IsComplete
        IsTruncated = $IsTruncated
    }
}

function Enter-TestWorkspace($Form) {
    $Form.Tag.ActiveServer = 'QueryServer'
    $Form.Tag.ActiveDatabase = 'QueryDatabase'
    Set-SqlUtilityStage -Form $Form -Stage 'Workspace'
    [System.Windows.Forms.Application]::DoEvents()
}

function Assert-TestQueryFailureState($Form, $Recorder, [string] $Prefix) {
    Assert-Equal $true ($null -eq $Form.Tag.CurrentResult) "$Prefix clears current result"
    Assert-Equal 0 $Form.Tag.CurrentPage "$Prefix clears page number"
    Assert-Equal $true ($null -eq (Get-TestControl $Form 'ResultsGrid').DataSource) "$Prefix clears grid"
    Assert-Equal $false (Get-TestControl $Form 'PreviousPageButton').Enabled "$Prefix disables Previous"
    Assert-Equal $false (Get-TestControl $Form 'NextPageButton').Enabled "$Prefix disables Next"
    Assert-Equal $false (Get-TestControl $Form 'ExportButton').Enabled "$Prefix disables Export"
    Assert-Equal $false $Form.Tag.IsBusy "$Prefix restores busy state"
    Assert-Equal $true $Form.Enabled "$Prefix restores form buttons"
    Assert-Equal $true (Get-TestControl $Form 'ExecuteButton').Enabled "$Prefix restores Execute button"
    Assert-Equal $false $Form.UseWaitCursor "$Prefix restores cursor"
    Assert-Equal 1 @($Recorder.Messages | Where-Object Icon -eq 'Error').Count "$Prefix shows one error popup"
}

$requiredControlNames = @(
    'ConnectionPanel', 'ServerTextBox', 'DatabaseTextBox', 'TestConnectionButton',
    'ConnectButton', 'SavedConnectionsList', 'DeleteConnectionButton',
    'WorkspacePanel', 'ActiveConnectionLabel', 'ChangeConnectionButton',
    'WorkspaceTabs', 'QueryTab', 'SettingsTab', 'UnorderedLimitNumeric',
    'QueryExportTimeoutNumeric', 'SaveSettingsButton', 'MainStatusLabel',
    'SqlEditor', 'ExecuteButton', 'ExportButton', 'PreviousPageButton',
    'NextPageButton', 'PageStatusLabel', 'ResultsGrid'
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

# Invalid policy results clear prior results and never reach a database executor.
$invalidHarness = New-TestServices
$invalidHarness.Recorder.ValidationResult = [pscustomobject][ordered]@{
    IsValid = $false
    ErrorMessage = 'Only one read-only SELECT is allowed.'
    NormalizedSql = ''
    TableIdentifier = ''
    HasOrderBy = $false
}
$invalidForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $invalidHarness.Services
try {
    Show-TestForm $invalidForm
    Enter-TestWorkspace $invalidForm
    $priorData = New-TestDataTable -RowCount 1
    $invalidForm.Tag.ExecutedQuery = [pscustomobject]@{ OriginalEditorSql = 'old'; NormalizedSql = 'old'; HasOrderBy = $true }
    $invalidForm.Tag.CurrentResult = New-TestPageResult -Data $priorData -PageNumber 1 -HasNext $true
    Show-SqlUtilityPage -Form $invalidForm -PageResult $invalidForm.Tag.CurrentResult
    (Get-TestControl $invalidForm 'SqlEditor').Text = 'DELETE FROM dbo.Items'
    (Get-TestControl $invalidForm 'ExecuteButton').PerformClick()

    Assert-Equal 1 $invalidHarness.Recorder.ValidationCalls.Count 'Execute validates editor SQL once'
    Assert-Equal 0 $invalidHarness.Recorder.OrderedCalls.Count 'Invalid policy never calls ordered database execution'
    Assert-Equal 0 $invalidHarness.Recorder.UnorderedCalls.Count 'Invalid policy never calls unordered database execution'
    Assert-Equal $null $invalidForm.Tag.ExecutedQuery 'Invalid policy clears executed snapshot'
    Assert-Equal $null $invalidForm.Tag.CurrentResult 'Invalid policy clears result state'
    Assert-Equal 0 $invalidForm.Tag.CurrentPage 'Invalid policy clears page state'
    Assert-Equal $null (Get-TestControl $invalidForm 'ResultsGrid').DataSource 'Invalid policy clears result grid'
    Assert-Equal 'Only one read-only SELECT is allowed.' $invalidHarness.Recorder.Messages[0].Text 'Invalid policy displays concise validator message'
}
finally {
    $invalidForm.Close()
    $invalidForm.Dispose()
}

# Ordered execution binds page metadata, reexecutes the normalized snapshot, caps grid widths, and becomes stale on edit.
$orderedHarness = New-TestServices
$orderedPage1Data = New-TestDataTable -RowCount 500 -TextLength 400
$orderedPage2Data = New-TestDataTable -RowCount 2
$orderedHarness.Recorder.OrderedResults[1] = New-TestPageResult -Data $orderedPage1Data -PageNumber 1 -HasNext $true
$orderedHarness.Recorder.OrderedResults[2] = New-TestPageResult -Data $orderedPage2Data -PageNumber 2 -HasPrevious $true
$orderedForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $orderedHarness.Services
try {
    Show-TestForm $orderedForm
    Enter-TestWorkspace $orderedForm
    $sqlEditor = Get-TestControl $orderedForm 'SqlEditor'
    $executeButton = Get-TestControl $orderedForm 'ExecuteButton'
    $previousButton = Get-TestControl $orderedForm 'PreviousPageButton'
    $nextButton = Get-TestControl $orderedForm 'NextPageButton'
    $exportButton = Get-TestControl $orderedForm 'ExportButton'
    $resultsGrid = Get-TestControl $orderedForm 'ResultsGrid'
    $originalEditorSql = " SELECT Id FROM dbo.Items`r`nORDER BY Id; "
    $sqlEditor.Text = $originalEditorSql
    $executeButton.PerformClick()

    Assert-Equal 1 $orderedHarness.Recorder.OrderedCalls.Count 'Ordered Execute calls ordered service once'
    Assert-Equal 'QueryServer' $orderedHarness.Recorder.OrderedCalls[0].Server 'Ordered Execute uses active server'
    Assert-Equal 'QueryDatabase' $orderedHarness.Recorder.OrderedCalls[0].Database 'Ordered Execute uses active database'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $orderedHarness.Recorder.OrderedCalls[0].Sql 'Ordered Execute uses normalized SQL'
    Assert-Equal 1 $orderedHarness.Recorder.OrderedCalls[0].PageNumber 'Ordered Execute requests page one'
    Assert-Equal 120 $orderedHarness.Recorder.OrderedCalls[0].TimeoutSeconds 'Ordered Execute uses configured timeout'
    Assert-Equal $originalEditorSql $orderedForm.Tag.ExecutedQuery.OriginalEditorSql 'Ordered Execute preserves exact editor snapshot'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $orderedForm.Tag.ExecutedQuery.NormalizedSql 'Ordered Execute stores normalized snapshot'
    Assert-Equal 500 $resultsGrid.Rows.Count 'Ordered page binds 500 displayed rows'
    Assert-Equal $true $resultsGrid.ReadOnly 'Results grid is read-only'
    Assert-Equal $false $resultsGrid.AllowUserToAddRows 'Results grid prevents row insertion'
    Assert-Equal $false $resultsGrid.AllowUserToDeleteRows 'Results grid prevents row deletion'
    Assert-Equal $false $previousButton.Enabled 'Ordered page one disables Previous'
    Assert-Equal $true $nextButton.Enabled 'Ordered page one enables Next from sentinel metadata'
    Assert-Equal $true $exportButton.Enabled 'Fresh ordered result enables export'
    Assert-Equal 'Page 1 - 500 rows' (Get-TestControl $orderedForm 'PageStatusLabel').Text 'Ordered status reports page and displayed rows'
    foreach ($column in $resultsGrid.Columns) {
        Assert-True ($column.Width -le 200) "Grid caps $($column.Name) at 200 pixels"
        Assert-Equal ([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None) $column.AutoSizeMode "Grid leaves $($column.Name) fixed after sizing"
    }

    $nextButton.PerformClick()
    Assert-Equal 2 $orderedHarness.Recorder.OrderedCalls.Count 'Ordered Next reexecutes query'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $orderedHarness.Recorder.OrderedCalls[1].Sql 'Ordered Next uses exact normalized executed snapshot'
    Assert-Equal 2 $orderedHarness.Recorder.OrderedCalls[1].PageNumber 'Ordered Next requests target page'
    Assert-Equal 'Page 2 - 2 rows' (Get-TestControl $orderedForm 'PageStatusLabel').Text 'Ordered Next updates page status'
    $previousButton.PerformClick()
    Assert-Equal 3 $orderedHarness.Recorder.OrderedCalls.Count 'Ordered Previous reexecutes query'
    Assert-Equal 1 $orderedHarness.Recorder.OrderedCalls[2].PageNumber 'Ordered Previous requests target page'

    $snapshotBeforeEdit = $orderedForm.Tag.ExecutedQuery
    $sqlEditor.Text = 'SELECT Id FROM dbo.Items ORDER BY Description'
    [System.Windows.Forms.Application]::DoEvents()
    Assert-Equal $true $orderedForm.Tag.IsQueryStale 'Editor change marks successful result stale'
    Assert-Equal $snapshotBeforeEdit $orderedForm.Tag.ExecutedQuery 'Editor change retains executed snapshot object'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $orderedForm.Tag.ExecutedQuery.NormalizedSql 'Editor change never mutates normalized snapshot'
    Assert-Equal 500 $resultsGrid.Rows.Count 'Editor change preserves displayed rows'
    Assert-Equal $false $previousButton.Enabled 'Stale result disables Previous'
    Assert-Equal $false $nextButton.Enabled 'Stale result disables Next'
    Assert-Equal $false $exportButton.Enabled 'Stale result disables Export'
}
finally {
    $orderedForm.Close()
    $orderedForm.Dispose()
}

# Complete unordered execution pages only through its cache and exports the complete cache.
$unorderedHarness = New-TestServices
$unorderedHarness.Recorder.ValidationResult = [pscustomobject][ordered]@{
    IsValid = $true
    ErrorMessage = ''
    NormalizedSql = 'SELECT Id FROM dbo.Items'
    TableIdentifier = 'dbo.Items'
    HasOrderBy = $false
}
$unorderedCache = New-TestDataTable -RowCount 650
$unorderedFirstPage = Get-SqlUtilityLocalPage -CachedData $unorderedCache -PageNumber 1 -IsComplete $true -IsTruncated $false
$unorderedHarness.Recorder.UnorderedResult = $unorderedFirstPage
$unorderedHarness.Recorder.PromptPath = 'C:\exports\complete.xlsx'
$unorderedForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $unorderedHarness.Services
try {
    Show-TestForm $unorderedForm
    Enter-TestWorkspace $unorderedForm
    (Get-TestControl $unorderedForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items;'
    (Get-TestControl $unorderedForm 'ExecuteButton').PerformClick()

    Assert-Equal 1 $unorderedHarness.Recorder.UnorderedCalls.Count 'Unordered Execute calls unordered service once'
    Assert-Equal 1000 $unorderedHarness.Recorder.UnorderedCalls[0].RowLimit 'Unordered Execute uses configured row limit'
    Assert-Equal 0 $unorderedHarness.Recorder.OrderedCalls.Count 'Unordered Execute never calls ordered service'
    Assert-Equal 0 $unorderedHarness.Recorder.Messages.Count 'Complete unordered result shows no truncation popup'
    Assert-Equal $true (Get-TestControl $unorderedForm 'ExportButton').Enabled 'Complete unordered result enables export'

    (Get-TestControl $unorderedForm 'NextPageButton').PerformClick()
    Assert-Equal 1 $unorderedHarness.Recorder.UnorderedCalls.Count 'Unordered Next never queries SQL again'
    Assert-Equal 1 $unorderedHarness.Recorder.LocalPageCalls.Count 'Unordered Next uses local-page service'
    Assert-True ([object]::ReferenceEquals($unorderedCache, $unorderedHarness.Recorder.LocalPageCalls[0].CachedData)) 'Unordered Next uses retained cache object'
    Assert-Equal 2 $unorderedHarness.Recorder.LocalPageCalls[0].PageNumber 'Unordered Next requests local page two'
    Assert-Equal 'Page 2 - 150 rows' (Get-TestControl $unorderedForm 'PageStatusLabel').Text 'Unordered local page shows remaining rows'

    (Get-TestControl $unorderedForm 'ExportButton').PerformClick()
    Assert-Equal 1 $unorderedHarness.Recorder.PromptCalls 'Export opens save prompt once'
    Assert-Equal 1 $unorderedHarness.Recorder.ExportCalls.Count 'Complete unordered export invokes exporter'
    Assert-Equal 'C:\exports\complete.xlsx' $unorderedHarness.Recorder.ExportCalls[0].DestinationPath 'Export forwards selected destination'
    Assert-True ([object]::ReferenceEquals($unorderedCache, $unorderedHarness.Recorder.ExportCalls[0].CachedData)) 'Complete unordered export passes entire cache through state'
    Assert-Equal 1 @($unorderedHarness.Recorder.Messages | Where-Object Icon -eq 'Information').Count 'Successful export shows one information popup'
    Assert-True (-not ($unorderedForm.Tag.PSObject.Properties.Name -contains 'DestinationPath')) 'Export destination is never retained in state'
}
finally {
    $unorderedForm.Close()
    $unorderedForm.Dispose()
}

# Truncated unordered execution retains the configured maximum, warns once, locally pages, and cannot export.
$truncatedHarness = New-TestServices
$truncatedHarness.Recorder.ValidationResult = [pscustomobject][ordered]@{
    IsValid = $true
    ErrorMessage = ''
    NormalizedSql = 'SELECT Id FROM dbo.Items'
    TableIdentifier = 'dbo.Items'
    HasOrderBy = $false
}
$truncatedCache = New-TestDataTable -RowCount 1000
$truncatedHarness.Recorder.UnorderedResult = Get-SqlUtilityLocalPage -CachedData $truncatedCache -PageNumber 1 -IsComplete $false -IsTruncated $true
$truncatedForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $truncatedHarness.Services
try {
    Show-TestForm $truncatedForm
    Enter-TestWorkspace $truncatedForm
    (Get-TestControl $truncatedForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items'
    (Get-TestControl $truncatedForm 'ExecuteButton').PerformClick()
    Assert-Equal 1000 $truncatedForm.Tag.CurrentResult.CachedData.Rows.Count 'Truncated result retains first configured maximum'
    Assert-Equal 1 @($truncatedHarness.Recorder.Messages | Where-Object Text -match 'ORDER BY').Count 'Truncated result asks for ORDER BY exactly once'
    Assert-Equal $false (Get-TestControl $truncatedForm 'ExportButton').Enabled 'Truncated result disables export'
    (Get-TestControl $truncatedForm 'NextPageButton').PerformClick()
    Assert-Equal 1 $truncatedHarness.Recorder.LocalPageCalls.Count 'Truncated result pages locally'
    Assert-Equal 500 (Get-TestControl $truncatedForm 'ResultsGrid').Rows.Count 'Truncated second page retains remaining bounded rows'
    Assert-Equal 1 @($truncatedHarness.Recorder.Messages | Where-Object Text -match 'ORDER BY').Count 'Local paging does not repeat truncation popup'
}
finally {
    $truncatedForm.Close()
    $truncatedForm.Dispose()
}

# Query exceptions clear page/export state and always restore the form and cursor.
$errorHarness = New-TestServices
$errorHarness.Recorder.ExecuteError = 'query timed out'
$errorHarness.Recorder.OrderedResults[1] = New-TestPageResult -Data (New-TestDataTable -RowCount 1) -PageNumber 1
$errorForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $errorHarness.Services
try {
    Show-TestForm $errorForm
    Enter-TestWorkspace $errorForm
    (Get-TestControl $errorForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items ORDER BY Id'
    (Get-TestControl $errorForm 'ExecuteButton').PerformClick()
    Assert-Equal $null $errorForm.Tag.CurrentResult 'Query exception clears current result'
    Assert-Equal 0 $errorForm.Tag.CurrentPage 'Query exception clears page number'
    Assert-Equal $null (Get-TestControl $errorForm 'ResultsGrid').DataSource 'Query exception clears grid'
    Assert-Equal $false (Get-TestControl $errorForm 'PreviousPageButton').Enabled 'Query exception disables Previous'
    Assert-Equal $false (Get-TestControl $errorForm 'NextPageButton').Enabled 'Query exception disables Next'
    Assert-Equal $false (Get-TestControl $errorForm 'ExportButton').Enabled 'Query exception disables Export'
    Assert-Equal $false $errorForm.Tag.IsBusy 'Query exception restores busy state'
    Assert-Equal $true $errorForm.Enabled 'Query exception restores buttons'
    Assert-Equal $false $errorForm.UseWaitCursor 'Query exception restores cursor'
    Assert-Equal 1 @($errorHarness.Recorder.Messages | Where-Object Icon -eq 'Error').Count 'Query exception shows one error popup'
}
finally {
    $errorForm.Close()
    $errorForm.Dispose()
}

# Ordered page reexecution failure clears every result control and restores the interactive form.
$orderedPageErrorHarness = New-TestServices
$orderedPageErrorHarness.Recorder.OrderedResults[1] = New-TestPageResult `
    -Data (New-TestDataTable -RowCount 500) -PageNumber 1 -HasNext $true
$orderedPageErrorHarness.Recorder.OrderedResults[2] = New-TestPageResult `
    -Data (New-TestDataTable -RowCount 500) -PageNumber 2 -HasPrevious $true -HasNext $true
$orderedPageErrorForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $orderedPageErrorHarness.Services
try {
    Show-TestForm $orderedPageErrorForm
    Enter-TestWorkspace $orderedPageErrorForm
    (Get-TestControl $orderedPageErrorForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items ORDER BY Id'
    (Get-TestControl $orderedPageErrorForm 'ExecuteButton').PerformClick()
    (Get-TestControl $orderedPageErrorForm 'NextPageButton').PerformClick()
    $orderedPageErrorHarness.Recorder.ExecuteError = 'ordered page timed out'
    (Get-TestControl $orderedPageErrorForm 'NextPageButton').PerformClick()

    Assert-Equal 3 $orderedPageErrorHarness.Recorder.OrderedCalls.Count 'Ordered page failure occurs on reexecution call'
    Assert-Equal 3 $orderedPageErrorHarness.Recorder.OrderedCalls[2].PageNumber 'Ordered page failure requested page three'
    Assert-TestQueryFailureState $orderedPageErrorForm $orderedPageErrorHarness.Recorder 'Ordered page failure'
}
finally {
    $orderedPageErrorForm.Close()
    $orderedPageErrorForm.Dispose()
}

# Unordered local-page failure clears every result control and restores the interactive form without another SQL call.
$localPageErrorHarness = New-TestServices
$localPageErrorHarness.Recorder.ValidationResult = [pscustomobject][ordered]@{
    IsValid = $true
    ErrorMessage = ''
    NormalizedSql = 'SELECT Id FROM dbo.Items'
    TableIdentifier = 'dbo.Items'
    HasOrderBy = $false
}
$localPageErrorConfig = New-TestConfig
$localPageErrorConfig.unorderedRowLimit = 1500
$localPageErrorCache = New-TestDataTable -RowCount 1200
$localPageErrorHarness.Recorder.UnorderedResult = Get-SqlUtilityLocalPage `
    -CachedData $localPageErrorCache -PageNumber 1 -IsComplete $true -IsTruncated $false
$localPageErrorForm = New-SqlUtilityMainForm -Config $localPageErrorConfig -ConfigPath 'C:\test\config.json' -Services $localPageErrorHarness.Services
try {
    Show-TestForm $localPageErrorForm
    Enter-TestWorkspace $localPageErrorForm
    (Get-TestControl $localPageErrorForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items'
    (Get-TestControl $localPageErrorForm 'ExecuteButton').PerformClick()
    (Get-TestControl $localPageErrorForm 'NextPageButton').PerformClick()
    $localPageErrorHarness.Recorder.LocalPageError = 'local page unavailable'
    (Get-TestControl $localPageErrorForm 'NextPageButton').PerformClick()

    Assert-Equal 1 $localPageErrorHarness.Recorder.UnorderedCalls.Count 'Local-page failure never repeats unordered SQL execution'
    Assert-Equal 2 $localPageErrorHarness.Recorder.LocalPageCalls.Count 'Local-page failure occurs in local-page service'
    Assert-Equal 3 $localPageErrorHarness.Recorder.LocalPageCalls[1].PageNumber 'Local-page failure requested page three'
    Assert-TestQueryFailureState $localPageErrorForm $localPageErrorHarness.Recorder 'Local-page failure'
}
finally {
    $localPageErrorForm.Close()
    $localPageErrorForm.Dispose()
}

# Export cancellation is inert; export failures report errors without clearing a valid result.
$exportHarness = New-TestServices
$exportHarness.Recorder.OrderedResults[1] = New-TestPageResult -Data (New-TestDataTable -RowCount 2) -PageNumber 1
$exportForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $exportHarness.Services
try {
    Show-TestForm $exportForm
    Enter-TestWorkspace $exportForm
    (Get-TestControl $exportForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items ORDER BY Id;'
    (Get-TestControl $exportForm 'ExecuteButton').PerformClick()
    $resultBeforeExport = $exportForm.Tag.CurrentResult
    (Get-TestControl $exportForm 'ExportButton').PerformClick()
    Assert-Equal 1 $exportHarness.Recorder.PromptCalls 'Canceled export still prompts once'
    Assert-Equal 0 $exportHarness.Recorder.ExportCalls.Count 'Canceled save dialog never invokes exporter'
    Assert-Equal 0 $exportHarness.Recorder.Messages.Count 'Canceled save dialog shows no popup'

    $exportHarness.Recorder.PromptPath = 'C:\exports\ordered.xlsx'
    $exportHarness.Recorder.ExportError = 'destination denied'
    (Get-TestControl $exportForm 'ExportButton').PerformClick()
    Assert-Equal 1 $exportHarness.Recorder.ExportCalls.Count 'Chosen ordered export invokes exporter once'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $exportHarness.Recorder.ExportCalls[0].NormalizedSql 'Ordered export receives exact unpaged normalized snapshot'
    Assert-Equal 1 @($exportHarness.Recorder.Messages | Where-Object Icon -eq 'Error').Count 'Failed export shows one error popup'
    Assert-Equal $resultBeforeExport $exportForm.Tag.CurrentResult 'Failed export preserves valid result state'
    Assert-Equal $false $exportForm.Tag.IsBusy 'Failed export restores busy state'
    Assert-Equal $false $exportForm.UseWaitCursor 'Failed export restores cursor'
}
finally {
    $exportForm.Close()
    $exportForm.Dispose()
}

# Save-path prompt failures are reported as export failures and never invoke the result exporter.
$promptErrorHarness = New-TestServices
$promptErrorHarness.Recorder.OrderedResults[1] = New-TestPageResult -Data (New-TestDataTable -RowCount 2) -PageNumber 1
$promptErrorHarness.Recorder.PromptError = 'save dialog unavailable'
$promptErrorForm = New-SqlUtilityMainForm -Config (New-TestConfig) -ConfigPath 'C:\test\config.json' -Services $promptErrorHarness.Services
try {
    Show-TestForm $promptErrorForm
    Enter-TestWorkspace $promptErrorForm
    (Get-TestControl $promptErrorForm 'SqlEditor').Text = 'SELECT Id FROM dbo.Items ORDER BY Id'
    (Get-TestControl $promptErrorForm 'ExecuteButton').PerformClick()
    $promptFailureResult = $promptErrorForm.Tag.CurrentResult
    (Get-TestControl $promptErrorForm 'ExportButton').PerformClick()

    Assert-Equal 1 $promptErrorHarness.Recorder.PromptCalls 'Prompt failure invokes save-path service once'
    Assert-Equal 0 $promptErrorHarness.Recorder.ExportCalls.Count 'Prompt failure never invokes result exporter'
    Assert-Equal 1 @($promptErrorHarness.Recorder.Messages | Where-Object Icon -eq 'Error').Count 'Prompt failure shows one export error popup'
    Assert-True ($promptErrorHarness.Recorder.Messages[0].Text -match 'save dialog unavailable') 'Prompt failure popup includes the prompt error'
    Assert-Equal $promptFailureResult $promptErrorForm.Tag.CurrentResult 'Prompt failure preserves valid result state'
    Assert-Equal $false $promptErrorForm.Tag.IsBusy 'Prompt failure leaves busy state clear'
    Assert-Equal $true $promptErrorForm.Enabled 'Prompt failure leaves form buttons enabled'
    Assert-Equal $true (Get-TestControl $promptErrorForm 'ExportButton').Enabled 'Prompt failure leaves Export enabled for retry'
    Assert-Equal $false $promptErrorForm.UseWaitCursor 'Prompt failure leaves cursor restored'
}
finally {
    $promptErrorForm.Close()
    $promptErrorForm.Dispose()
}

# The production export workflow composes ordered SQL streaming and unordered cached rows through neutral callbacks.
$script:queryCompositionRecorder = [pscustomobject]@{
    OrderedCalls = [System.Collections.Generic.List[object]]::new()
    Exports = [System.Collections.Generic.List[object]]::new()
}
$originalOrderedRowStream = ${function:Invoke-SqlUtilityOrderedRowStream}
$originalXlsxExport = ${function:Export-SqlUtilityXlsx}
try {
    Set-Item -Path Function:\Invoke-SqlUtilityOrderedRowStream -Value {
        param($Server, $Database, $Sql, $CommandTimeoutSeconds, $OnSchema, $OnRow, $ShouldContinue)
        [void] $script:queryCompositionRecorder.OrderedCalls.Add([pscustomobject]@{
            Server = $Server
            Database = $Database
            Sql = $Sql
            TimeoutSeconds = $CommandTimeoutSeconds
        })
        $null = & $OnSchema @([pscustomobject]@{ Name = 'Id'; DataType = [int]; Ordinal = 0 })
        if (& $ShouldContinue) { $null = & $OnRow ([object[]] @(7)) }
    }
    Set-Item -Path Function:\Export-SqlUtilityXlsx -Value {
        param($DestinationPath, $RowSource, $TimeoutSeconds)
        $capture = [pscustomobject]@{
            DestinationPath = $DestinationPath
            TimeoutSeconds = $TimeoutSeconds
            SchemaCount = 0
            RowCount = 0
        }
        & $RowSource `
            { param($Schema) $capture.SchemaCount = @($Schema).Count }.GetNewClosure() `
            { param($Values) $capture.RowCount++ }.GetNewClosure() `
            { return $true }
        [void] $script:queryCompositionRecorder.Exports.Add($capture)
    }

    $orderedState = [pscustomobject]@{
        ActiveServer = 'CompositionServer'
        ActiveDatabase = 'CompositionDatabase'
        Config = [pscustomobject]@{ queryExportTimeoutSeconds = 321 }
        ExecutedQuery = [pscustomobject]@{
            HasOrderBy = $true
            NormalizedSql = 'SELECT Id FROM dbo.Items ORDER BY Id'
        }
        CurrentResult = New-TestPageResult -Data (New-TestDataTable -RowCount 1) -PageNumber 1
    }
    Invoke-SqlUtilityExportWorkflow -State $orderedState -DestinationPath 'C:\exports\composition-ordered.xlsx'
    Assert-Equal 1 $script:queryCompositionRecorder.OrderedCalls.Count 'Ordered export workflow opens one neutral row stream'
    Assert-Equal 'SELECT Id FROM dbo.Items ORDER BY Id' $script:queryCompositionRecorder.OrderedCalls[0].Sql 'Ordered export workflow streams exact normalized unpaged snapshot'
    Assert-Equal 321 $script:queryCompositionRecorder.OrderedCalls[0].TimeoutSeconds 'Ordered export workflow forwards configured timeout to SQL stream'
    Assert-Equal 1 $script:queryCompositionRecorder.Exports[0].RowCount 'Ordered row stream reaches neutral Excel exporter'

    $completeCache = New-TestDataTable -RowCount 3
    $unorderedState = [pscustomobject]@{
        ActiveServer = 'UnusedServer'
        ActiveDatabase = 'UnusedDatabase'
        Config = [pscustomobject]@{ queryExportTimeoutSeconds = 654 }
        ExecutedQuery = [pscustomobject]@{ HasOrderBy = $false; NormalizedSql = 'SELECT Id FROM dbo.Items' }
        CurrentResult = New-TestPageResult -Data $completeCache -CachedData $completeCache -PageNumber 1 -IsComplete $true
    }
    Invoke-SqlUtilityExportWorkflow -State $unorderedState -DestinationPath 'C:\exports\composition-unordered.xlsx'
    Assert-Equal 1 $script:queryCompositionRecorder.OrderedCalls.Count 'Unordered export workflow never opens SQL row stream'
    Assert-Equal 3 $script:queryCompositionRecorder.Exports[1].RowCount 'Unordered export workflow supplies every cached row'
    Assert-Equal 654 $script:queryCompositionRecorder.Exports[1].TimeoutSeconds 'Unordered export workflow forwards configured timeout to Excel'
}
finally {
    Set-Item -Path Function:\Invoke-SqlUtilityOrderedRowStream -Value $originalOrderedRowStream
    Set-Item -Path Function:\Export-SqlUtilityXlsx -Value $originalXlsxExport
    Remove-Variable -Name queryCompositionRecorder -Scope Script -ErrorAction SilentlyContinue
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
