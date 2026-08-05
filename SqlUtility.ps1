param(
    [switch] $NoGui,
    [string] $ConfigPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

. (Join-Path $PSScriptRoot 'modules\SqlUtility.Config.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.QueryPolicy.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.Database.ps1')
. (Join-Path $PSScriptRoot 'modules\SqlUtility.Excel.ps1')

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $PSScriptRoot 'SqlUtility.config.json'
}

function New-SqlUtilityDefaultServices {
    [CmdletBinding()]
    param()

    return @{
        TestConnection = {
            param($Server, $Database)
            Invoke-SqlUtilityConnectionTest -Server $Server -Database $Database
        }
        WriteConfig = {
            param($Path, $Config)
            return Write-SqlUtilityConfig -Path $Path -Config $Config
        }
        ShowMessage = {
            param($Text, $Caption, $Icon)
            $messageIcon = [System.Windows.Forms.MessageBoxIcon]::None
            if (-not [string]::IsNullOrWhiteSpace([string] $Icon)) {
                $messageIcon = [System.Windows.Forms.MessageBoxIcon] [System.Enum]::Parse(
                    [System.Windows.Forms.MessageBoxIcon],
                    [string] $Icon,
                    $true
                )
            }
            [void] [System.Windows.Forms.MessageBox]::Show(
                [string] $Text,
                [string] $Caption,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $messageIcon
            )
        }
        Confirm = {
            param($Text, $Caption)
            $answer = [System.Windows.Forms.MessageBox]::Show(
                [string] $Text,
                [string] $Caption,
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
            return $answer -eq [System.Windows.Forms.DialogResult]::Yes
        }
    }
}

function Assert-SqlUtilityServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable] $Services
    )

    foreach ($serviceName in @('TestConnection', 'WriteConfig', 'ShowMessage', 'Confirm')) {
        if (-not $Services.ContainsKey($serviceName) -or $Services[$serviceName] -isnot [scriptblock]) {
            throw [System.ArgumentException]::new("Services must contain a '$serviceName' scriptblock.")
        }
    }
}

function Show-SqlUtilityMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $State,
        [Parameter(Mandatory = $true)][string] $Text,
        [Parameter(Mandatory = $true)][string] $Caption,
        [Parameter(Mandatory = $true)][string] $Icon
    )

    $showMessage = $State.Services['ShowMessage']
    & $showMessage $Text $Caption $Icon
}

function Confirm-SqlUtilityAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $State,
        [Parameter(Mandatory = $true)][string] $Text,
        [Parameter(Mandatory = $true)][string] $Caption
    )

    $confirm = $State.Services['Confirm']
    return [bool] (& $confirm $Text $Caption)
}

function Get-SqlUtilityNamedControl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Control] $Root,
        [Parameter(Mandatory = $true)][string] $Name
    )

    $matches = @($Root.Controls.Find($Name, $true))
    if ($matches.Count -eq 0) {
        return $null
    }
    return $matches[0]
}

function Update-SqlUtilitySavedConnections {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $list = Get-SqlUtilityNamedControl -Root $Form -Name 'SavedConnectionsList'
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        foreach ($connection in @($Form.Tag.Config.connections)) {
            $item = [pscustomobject][ordered]@{
                Server = [string] $connection.server
                Database = [string] $connection.database
                Display = ('{0} {1} {2}' -f $connection.server, [char] 0x2013, $connection.database)
            }
            [void] $list.Items.Add($item)
        }
        $list.ClearSelected()
    }
    finally {
        $list.EndUpdate()
    }
}

function Set-SqlUtilityBusy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)][bool] $Busy,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $Message
    )

    $Form.Tag.IsBusy = $Busy
    $Form.UseWaitCursor = $Busy
    $Form.Enabled = -not $Busy
    $statusLabel = Get-SqlUtilityNamedControl -Root $Form -Name 'MainStatusLabel'
    if ($null -ne $statusLabel) {
        $statusLabel.Text = $Message
    }
    $Form.Refresh()
}

function Set-SqlUtilityStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)][ValidateSet('Connection', 'Workspace')][string] $Stage
    )

    $connectionPanel = Get-SqlUtilityNamedControl -Root $Form -Name 'ConnectionPanel'
    $workspacePanel = Get-SqlUtilityNamedControl -Root $Form -Name 'WorkspacePanel'
    if ($Stage -eq 'Connection') {
        $workspacePanel.Visible = $false
        $connectionPanel.Visible = $true
        $connectionPanel.BringToFront()
    }
    else {
        $connectionPanel.Visible = $false
        $workspacePanel.Visible = $true
        $workspacePanel.BringToFront()
        $tabs = Get-SqlUtilityNamedControl -Root $Form -Name 'WorkspaceTabs'
        $queryTab = Get-SqlUtilityNamedControl -Root $Form -Name 'QueryTab'
        $tabs.SelectedTab = $queryTab
        $activeLabel = Get-SqlUtilityNamedControl -Root $Form -Name 'ActiveConnectionLabel'
        $activeLabel.Text = ('{0} {1} {2}' -f $Form.Tag.ActiveServer, [char] 0x2013, $Form.Tag.ActiveDatabase)
    }
}

function Invoke-SqlUtilityConnectionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)][bool] $EnterWorkspace
    )

    $state = $Form.Tag
    if ($state.IsBusy) {
        return
    }

    $serverTextBox = Get-SqlUtilityNamedControl -Root $Form -Name 'ServerTextBox'
    $databaseTextBox = Get-SqlUtilityNamedControl -Root $Form -Name 'DatabaseTextBox'
    $server = $serverTextBox.Text.Trim()
    $database = $databaseTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($database)) {
        Show-SqlUtilityMessage -State $state -Text 'Enter both a server and a database.' -Caption 'Connection Required' -Icon 'Warning'
        return
    }

    $serverTextBox.Text = $server
    $databaseTextBox.Text = $database
    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Testing connection...'
    try {
        $testConnection = $state.Services['TestConnection']
        & $testConnection $server $database

        $candidate = Add-SqlUtilitySavedConnection -Config $state.Config -Server $server -Database $database
        $state.Config = $candidate
        Update-SqlUtilitySavedConnections -Form $Form

        try {
            $writeConfig = $state.Services['WriteConfig']
            $persisted = & $writeConfig $state.ConfigPath $candidate
            $state.Config = ConvertTo-SqlUtilityValidatedConfig -InputObject $persisted
            Update-SqlUtilitySavedConnections -Form $Form
        }
        catch {
            Show-SqlUtilityMessage -State $state `
                -Text ("The connection succeeded, but the saved-connection file could not be updated.`r`n`r`n{0}" -f $_.Exception.Message) `
                -Caption 'Configuration Warning' -Icon 'Warning'
        }

        if ($EnterWorkspace) {
            $state.ActiveServer = $server
            $state.ActiveDatabase = $database
            Set-SqlUtilityStage -Form $Form -Stage 'Workspace'
        }
        else {
            Show-SqlUtilityMessage -State $state -Text 'Connection succeeded.' -Caption 'Test Connection' -Icon 'Information'
        }
    }
    catch {
        Show-SqlUtilityMessage -State $state `
            -Text ("Connection failed.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Connection Error' -Icon 'Error'
    }
    finally {
        Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
    }
}

function Invoke-SqlUtilityDeleteConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    if ($state.IsBusy) {
        return
    }

    $list = Get-SqlUtilityNamedControl -Root $Form -Name 'SavedConnectionsList'
    $selected = $list.SelectedItem
    if ($null -eq $selected) {
        Show-SqlUtilityMessage -State $state -Text 'Select a saved connection to delete.' -Caption 'Saved Connections' -Icon 'Information'
        return
    }

    $display = ('{0} {1} {2}' -f $selected.Server, [char] 0x2013, $selected.Database)
    if (-not (Confirm-SqlUtilityAction -State $state -Text "Delete the saved connection ${display}?" -Caption 'Delete Saved Connection')) {
        return
    }

    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Deleting saved connection...'
    try {
        $candidate = Remove-SqlUtilitySavedConnection -Config $state.Config -Server $selected.Server -Database $selected.Database
        $writeConfig = $state.Services['WriteConfig']
        $persisted = & $writeConfig $state.ConfigPath $candidate
        $state.Config = ConvertTo-SqlUtilityValidatedConfig -InputObject $persisted
        Update-SqlUtilitySavedConnections -Form $Form
    }
    catch {
        Show-SqlUtilityMessage -State $state `
            -Text ("The saved connection could not be deleted.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Configuration Error' -Icon 'Error'
    }
    finally {
        Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
    }
}

function Invoke-SqlUtilitySaveSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    if ($state.IsBusy) {
        return
    }

    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Saving settings...'
    try {
        $unorderedNumeric = Get-SqlUtilityNamedControl -Root $Form -Name 'UnorderedLimitNumeric'
        $timeoutNumeric = Get-SqlUtilityNamedControl -Root $Form -Name 'QueryExportTimeoutNumeric'
        $candidate = ConvertTo-SqlUtilityValidatedConfig -InputObject ([pscustomobject][ordered]@{
            schemaVersion = $state.Config.schemaVersion
            unorderedRowLimit = [int] $unorderedNumeric.Value
            queryExportTimeoutSeconds = [int] $timeoutNumeric.Value
            connections = @($state.Config.connections)
        })

        $writeConfig = $state.Services['WriteConfig']
        $persisted = & $writeConfig $state.ConfigPath $candidate
        $state.Config = ConvertTo-SqlUtilityValidatedConfig -InputObject $persisted
        $unorderedNumeric.Value = $state.Config.unorderedRowLimit
        $timeoutNumeric.Value = $state.Config.queryExportTimeoutSeconds
        Show-SqlUtilityMessage -State $state -Text 'Settings saved.' -Caption 'Settings' -Icon 'Information'
    }
    catch {
        Show-SqlUtilityMessage -State $state `
            -Text ("Settings could not be saved.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Configuration Error' -Icon 'Error'
    }
    finally {
        Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
    }
}

function Reset-SqlUtilityWorkspaceState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    $state.ActiveServer = ''
    $state.ActiveDatabase = ''
    $state.ExecutedQuery = $null
    $state.CurrentResult = $null
    $state.CurrentPage = 0
    $state.IsQueryStale = $false

    foreach ($textBoxName in @('ServerTextBox', 'DatabaseTextBox', 'SqlEditor')) {
        $textBox = Get-SqlUtilityNamedControl -Root $Form -Name $textBoxName
        if ($null -ne $textBox) {
            $textBox.Text = ''
        }
    }

    $resultsGrid = Get-SqlUtilityNamedControl -Root $Form -Name 'ResultsGrid'
    if ($null -ne $resultsGrid) {
        $resultsGrid.DataSource = $null
        $resultsGrid.Rows.Clear()
        $resultsGrid.Columns.Clear()
    }
    foreach ($buttonName in @('ExportButton', 'PreviousPageButton', 'NextPageButton')) {
        $button = Get-SqlUtilityNamedControl -Root $Form -Name $buttonName
        if ($null -ne $button) {
            $button.Enabled = $false
        }
    }
    $pageStatus = Get-SqlUtilityNamedControl -Root $Form -Name 'PageStatusLabel'
    if ($null -ne $pageStatus) {
        $pageStatus.Text = ''
    }
}

function Invoke-SqlUtilityChangeConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    if ($state.IsBusy) {
        return
    }

    $message = 'Change connection? The current SQL text, results, page, and export state will be cleared.'
    if (-not (Confirm-SqlUtilityAction -State $state -Text $message -Caption 'Change Connection')) {
        return
    }

    Reset-SqlUtilityWorkspaceState -Form $Form
    Update-SqlUtilitySavedConnections -Form $Form
    Set-SqlUtilityStage -Form $Form -Stage 'Connection'
}

function New-SqlUtilityMainForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Config,
        [Parameter(Mandatory = $true)][string] $ConfigPath,
        [Parameter(Mandatory = $true)][hashtable] $Services
    )

    Assert-SqlUtilityServices -Services $Services
    $validatedConfig = ConvertTo-SqlUtilityValidatedConfig -InputObject $Config

    $form = [System.Windows.Forms.Form]::new()
    $form.Name = 'SqlUtilityMainForm'
    $form.Text = 'SQL Utility'
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Size = [System.Drawing.Size]::new(960, 680)
    $form.MinimumSize = [System.Drawing.Size]::new(760, 520)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

    $state = [pscustomobject][ordered]@{
        Config = $validatedConfig
        ConfigPath = $ConfigPath
        Services = $Services
        ActiveServer = ''
        ActiveDatabase = ''
        ExecutedQuery = $null
        CurrentResult = $null
        CurrentPage = 0
        IsQueryStale = $false
        IsBusy = $false
    }
    $form.Tag = $state

    $statusLabel = [System.Windows.Forms.Label]::new()
    $statusLabel.Name = 'MainStatusLabel'
    $statusLabel.Text = 'Ready.'
    $statusLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $statusLabel.Height = 24
    $statusLabel.Padding = [System.Windows.Forms.Padding]::new(6, 4, 6, 2)
    $statusLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $form.Controls.Add($statusLabel)

    $connectionPanel = [System.Windows.Forms.Panel]::new()
    $connectionPanel.Name = 'ConnectionPanel'
    $connectionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $connectionPanel.Padding = [System.Windows.Forms.Padding]::new(18)
    $form.Controls.Add($connectionPanel)

    $connectionLayout = [System.Windows.Forms.TableLayoutPanel]::new()
    $connectionLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $connectionLayout.ColumnCount = 2
    $connectionLayout.RowCount = 1
    [void] $connectionLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 52))
    [void] $connectionLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 48))
    $connectionPanel.Controls.Add($connectionLayout)

    $inputGroup = [System.Windows.Forms.GroupBox]::new()
    $inputGroup.Text = 'Connection'
    $inputGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
    $inputGroup.Padding = [System.Windows.Forms.Padding]::new(14)
    $inputGroup.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 0)
    $connectionLayout.Controls.Add($inputGroup, 0, 0)

    $inputLayout = [System.Windows.Forms.TableLayoutPanel]::new()
    $inputLayout.Dock = [System.Windows.Forms.DockStyle]::Top
    $inputLayout.AutoSize = $true
    $inputLayout.ColumnCount = 2
    $inputLayout.RowCount = 4
    [void] $inputLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void] $inputLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    $inputGroup.Controls.Add($inputLayout)

    $serverLabel = [System.Windows.Forms.Label]::new()
    $serverLabel.Text = 'Server'
    $serverLabel.AutoSize = $true
    $serverLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $serverLabel.Margin = [System.Windows.Forms.Padding]::new(0, 7, 12, 7)
    $inputLayout.Controls.Add($serverLabel, 0, 0)

    $serverTextBox = [System.Windows.Forms.TextBox]::new()
    $serverTextBox.Name = 'ServerTextBox'
    $serverTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $serverTextBox.Margin = [System.Windows.Forms.Padding]::new(0, 4, 0, 4)
    $inputLayout.Controls.Add($serverTextBox, 1, 0)

    $databaseLabel = [System.Windows.Forms.Label]::new()
    $databaseLabel.Text = 'Database'
    $databaseLabel.AutoSize = $true
    $databaseLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $databaseLabel.Margin = [System.Windows.Forms.Padding]::new(0, 7, 12, 7)
    $inputLayout.Controls.Add($databaseLabel, 0, 1)

    $databaseTextBox = [System.Windows.Forms.TextBox]::new()
    $databaseTextBox.Name = 'DatabaseTextBox'
    $databaseTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $databaseTextBox.Margin = [System.Windows.Forms.Padding]::new(0, 4, 0, 4)
    $inputLayout.Controls.Add($databaseTextBox, 1, 1)

    $testConnectionButton = [System.Windows.Forms.Button]::new()
    $testConnectionButton.Name = 'TestConnectionButton'
    $testConnectionButton.Text = 'Test Connection'
    $testConnectionButton.AutoSize = $true
    $testConnectionButton.Margin = [System.Windows.Forms.Padding]::new(0, 12, 8, 0)
    $inputLayout.Controls.Add($testConnectionButton, 0, 2)

    $connectButton = [System.Windows.Forms.Button]::new()
    $connectButton.Name = 'ConnectButton'
    $connectButton.Text = 'Connect'
    $connectButton.AutoSize = $true
    $connectButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $connectButton.Margin = [System.Windows.Forms.Padding]::new(0, 12, 0, 0)
    $inputLayout.Controls.Add($connectButton, 1, 2)

    $savedGroup = [System.Windows.Forms.GroupBox]::new()
    $savedGroup.Text = 'Saved Connections'
    $savedGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
    $savedGroup.Padding = [System.Windows.Forms.Padding]::new(12)
    $savedGroup.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
    $connectionLayout.Controls.Add($savedGroup, 1, 0)

    $deleteConnectionButton = [System.Windows.Forms.Button]::new()
    $deleteConnectionButton.Name = 'DeleteConnectionButton'
    $deleteConnectionButton.Text = 'Delete'
    $deleteConnectionButton.AutoSize = $true
    $deleteConnectionButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $deleteConnectionButton.Margin = [System.Windows.Forms.Padding]::new(0, 8, 0, 0)
    $savedGroup.Controls.Add($deleteConnectionButton)

    $savedConnectionsList = [System.Windows.Forms.ListBox]::new()
    $savedConnectionsList.Name = 'SavedConnectionsList'
    $savedConnectionsList.DisplayMember = 'Display'
    $savedConnectionsList.Dock = [System.Windows.Forms.DockStyle]::Fill
    $savedGroup.Controls.Add($savedConnectionsList)

    $workspacePanel = [System.Windows.Forms.Panel]::new()
    $workspacePanel.Name = 'WorkspacePanel'
    $workspacePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $workspacePanel.Padding = [System.Windows.Forms.Padding]::new(12)
    $workspacePanel.Visible = $false
    $form.Controls.Add($workspacePanel)

    $workspaceHeader = [System.Windows.Forms.Panel]::new()
    $workspaceHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $workspaceHeader.Height = 42
    $workspacePanel.Controls.Add($workspaceHeader)

    $changeConnectionButton = [System.Windows.Forms.Button]::new()
    $changeConnectionButton.Name = 'ChangeConnectionButton'
    $changeConnectionButton.Text = 'Change Connection'
    $changeConnectionButton.AutoSize = $true
    $changeConnectionButton.Dock = [System.Windows.Forms.DockStyle]::Right
    $workspaceHeader.Controls.Add($changeConnectionButton)

    $activeConnectionLabel = [System.Windows.Forms.Label]::new()
    $activeConnectionLabel.Name = 'ActiveConnectionLabel'
    $activeConnectionLabel.Text = ''
    $activeConnectionLabel.AutoSize = $false
    $activeConnectionLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $activeConnectionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $activeConnectionLabel.Font = [System.Drawing.Font]::new($form.Font, [System.Drawing.FontStyle]::Bold)
    $workspaceHeader.Controls.Add($activeConnectionLabel)

    $workspaceTabs = [System.Windows.Forms.TabControl]::new()
    $workspaceTabs.Name = 'WorkspaceTabs'
    $workspaceTabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $workspacePanel.Controls.Add($workspaceTabs)
    $workspaceTabs.BringToFront()

    $queryTab = [System.Windows.Forms.TabPage]::new()
    $queryTab.Name = 'QueryTab'
    $queryTab.Text = 'Query'
    $queryTab.UseVisualStyleBackColor = $true
    [void] $workspaceTabs.TabPages.Add($queryTab)

    $queryPlaceholder = [System.Windows.Forms.Label]::new()
    $queryPlaceholder.Text = 'Query controls are added in the next stage.'
    $queryPlaceholder.AutoSize = $true
    $queryPlaceholder.Location = [System.Drawing.Point]::new(16, 18)
    $queryTab.Controls.Add($queryPlaceholder)

    $settingsTab = [System.Windows.Forms.TabPage]::new()
    $settingsTab.Name = 'SettingsTab'
    $settingsTab.Text = 'Settings'
    $settingsTab.UseVisualStyleBackColor = $true
    $settingsTab.Padding = [System.Windows.Forms.Padding]::new(18)
    [void] $workspaceTabs.TabPages.Add($settingsTab)

    $unorderedLabel = [System.Windows.Forms.Label]::new()
    $unorderedLabel.Text = 'Maximum unordered rows'
    $unorderedLabel.AutoSize = $true
    $unorderedLabel.Location = [System.Drawing.Point]::new(22, 26)
    $settingsTab.Controls.Add($unorderedLabel)

    $unorderedLimitNumeric = [System.Windows.Forms.NumericUpDown]::new()
    $unorderedLimitNumeric.Name = 'UnorderedLimitNumeric'
    $unorderedLimitNumeric.Minimum = 100
    $unorderedLimitNumeric.Maximum = 2000
    $unorderedLimitNumeric.Increment = 100
    $unorderedLimitNumeric.Value = $validatedConfig.unorderedRowLimit
    $unorderedLimitNumeric.Location = [System.Drawing.Point]::new(260, 22)
    $unorderedLimitNumeric.Width = 110
    $settingsTab.Controls.Add($unorderedLimitNumeric)

    $timeoutLabel = [System.Windows.Forms.Label]::new()
    $timeoutLabel.Text = 'Query/Export timeout (seconds)'
    $timeoutLabel.AutoSize = $true
    $timeoutLabel.Location = [System.Drawing.Point]::new(22, 68)
    $settingsTab.Controls.Add($timeoutLabel)

    $queryExportTimeoutNumeric = [System.Windows.Forms.NumericUpDown]::new()
    $queryExportTimeoutNumeric.Name = 'QueryExportTimeoutNumeric'
    $queryExportTimeoutNumeric.Minimum = 5
    $queryExportTimeoutNumeric.Maximum = 3600
    $queryExportTimeoutNumeric.Value = $validatedConfig.queryExportTimeoutSeconds
    $queryExportTimeoutNumeric.Location = [System.Drawing.Point]::new(260, 64)
    $queryExportTimeoutNumeric.Width = 110
    $settingsTab.Controls.Add($queryExportTimeoutNumeric)

    $settingsHelp = [System.Windows.Forms.Label]::new()
    $settingsHelp.Text = 'Query/Export timeout (seconds) covers interactive queries and complete Excel export; connection timeout remains fixed and separate.'
    $settingsHelp.AutoSize = $false
    $settingsHelp.Location = [System.Drawing.Point]::new(22, 108)
    $settingsHelp.Size = [System.Drawing.Size]::new(680, 42)
    $settingsHelp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsTab.Controls.Add($settingsHelp)

    $saveSettingsButton = [System.Windows.Forms.Button]::new()
    $saveSettingsButton.Name = 'SaveSettingsButton'
    $saveSettingsButton.Text = 'Save Settings'
    $saveSettingsButton.AutoSize = $true
    $saveSettingsButton.Location = [System.Drawing.Point]::new(22, 164)
    $settingsTab.Controls.Add($saveSettingsButton)

    $savedConnectionsList.Add_SelectedIndexChanged({
        if ($null -ne $savedConnectionsList.SelectedItem) {
            $serverTextBox.Text = [string] $savedConnectionsList.SelectedItem.Server
            $databaseTextBox.Text = [string] $savedConnectionsList.SelectedItem.Database
        }
    }.GetNewClosure())
    $testConnectionButton.Add_Click({ Invoke-SqlUtilityConnectionAction -Form $form -EnterWorkspace $false }.GetNewClosure())
    $connectButton.Add_Click({ Invoke-SqlUtilityConnectionAction -Form $form -EnterWorkspace $true }.GetNewClosure())
    $deleteConnectionButton.Add_Click({ Invoke-SqlUtilityDeleteConnection -Form $form }.GetNewClosure())
    $saveSettingsButton.Add_Click({ Invoke-SqlUtilitySaveSettings -Form $form }.GetNewClosure())
    $changeConnectionButton.Add_Click({ Invoke-SqlUtilityChangeConnection -Form $form }.GetNewClosure())

    Update-SqlUtilitySavedConnections -Form $form
    $serverTextBox.Text = ''
    $databaseTextBox.Text = ''
    Set-SqlUtilityStage -Form $form -Stage 'Connection'
    return $form
}

function Start-SqlUtilityApplication {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ConfigPath,
        [hashtable] $Services
    )

    if ($null -eq $Services) {
        $Services = New-SqlUtilityDefaultServices
    }
    Assert-SqlUtilityServices -Services $Services

    try {
        $config = Read-SqlUtilityConfig -Path $ConfigPath
    }
    catch [System.NotSupportedException] {
        $showMessage = $Services['ShowMessage']
        & $showMessage $_.Exception.Message 'SQL Utility Configuration Error' 'Error'
        return
    }
    catch {
        $showMessage = $Services['ShowMessage']
        & $showMessage ("The configuration file could not be read.`r`n`r`n{0}" -f $_.Exception.Message) 'SQL Utility Configuration Error' 'Error'
        $confirm = $Services['Confirm']
        $shouldReset = [bool] (& $confirm 'Reset the malformed configuration file to defaults? The existing file will be overwritten.' 'Reset Configuration')
        if (-not $shouldReset) {
            return
        }

        try {
            $defaults = New-SqlUtilityDefaultConfig
            $writeConfig = $Services['WriteConfig']
            $config = & $writeConfig $ConfigPath $defaults
            $config = ConvertTo-SqlUtilityValidatedConfig -InputObject $config
        }
        catch {
            & $showMessage ("The configuration file could not be reset.`r`n`r`n{0}" -f $_.Exception.Message) 'SQL Utility Configuration Error' 'Error'
            return
        }
    }

    $form = $null
    try {
        $form = New-SqlUtilityMainForm -Config $config -ConfigPath $ConfigPath -Services $Services
        if ($Services.ContainsKey('ShowDialog') -and $Services['ShowDialog'] -is [scriptblock]) {
            $showDialog = $Services['ShowDialog']
            & $showDialog $form
        }
        else {
            [void] $form.ShowDialog()
        }
    }
    finally {
        if ($null -ne $form) {
            $form.Dispose()
        }
    }
}

if (-not $NoGui) {
    Start-SqlUtilityApplication -ConfigPath $ConfigPath -Services (New-SqlUtilityDefaultServices)
}
