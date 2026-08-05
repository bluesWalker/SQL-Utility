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

function Invoke-SqlUtilityExportWorkflow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $State,
        [Parameter(Mandatory = $true)][string] $DestinationPath
    )

    if ($State.ExecutedQuery.HasOrderBy) {
        $rowSource = {
            param($OnSchema, $OnRow, $ShouldContinue)
            Invoke-SqlUtilityOrderedRowStream -Server $State.ActiveServer `
                -Database $State.ActiveDatabase -Sql $State.ExecutedQuery.NormalizedSql `
                -CommandTimeoutSeconds $State.Config.queryExportTimeoutSeconds `
                -OnSchema $OnSchema -OnRow $OnRow -ShouldContinue $ShouldContinue
        }.GetNewClosure()
    }
    else {
        $rowSource = New-SqlUtilityDataTableRowSource -DataTable $State.CurrentResult.CachedData
    }

    Export-SqlUtilityXlsx -DestinationPath $DestinationPath -RowSource $rowSource `
        -TimeoutSeconds $State.Config.queryExportTimeoutSeconds
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
        ValidateQuery = {
            param($Sql)
            return Test-SqlUtilityQuery -Sql $Sql
        }
        ExecuteOrderedPage = {
            param($Server, $Database, $Sql, $PageNumber, $TimeoutSeconds)
            return Invoke-SqlUtilityOrderedPage -Server $Server -Database $Database -Sql $Sql `
                -PageNumber $PageNumber -CommandTimeoutSeconds $TimeoutSeconds
        }
        ExecuteUnordered = {
            param($Server, $Database, $Sql, $RowLimit, $TimeoutSeconds)
            return Invoke-SqlUtilityUnorderedQuery -Server $Server -Database $Database -Sql $Sql `
                -RowLimit $RowLimit -CommandTimeoutSeconds $TimeoutSeconds
        }
        GetLocalPage = {
            param($CachedData, $PageNumber, $IsComplete, $IsTruncated)
            return Get-SqlUtilityLocalPage -CachedData $CachedData -PageNumber $PageNumber `
                -IsComplete $IsComplete -IsTruncated $IsTruncated
        }
        ExportResult = {
            param($State, $DestinationPath)
            Invoke-SqlUtilityExportWorkflow -State $State -DestinationPath $DestinationPath
        }
        PromptSavePath = {
            $dialog = [System.Windows.Forms.SaveFileDialog]::new()
            try {
                $dialog.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
                $dialog.DefaultExt = 'xlsx'
                $dialog.AddExtension = $true
                $dialog.OverwritePrompt = $true
                if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                    return $null
                }
                return [string] $dialog.FileName
            }
            finally {
                $dialog.Dispose()
            }
        }
    }
}

function Assert-SqlUtilityServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable] $Services
    )

    foreach ($serviceName in @(
        'TestConnection', 'WriteConfig', 'ShowMessage', 'Confirm', 'ValidateQuery',
        'ExecuteOrderedPage', 'ExecuteUnordered', 'GetLocalPage', 'ExportResult', 'PromptSavePath'
    )) {
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

function Clear-SqlUtilityQueryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    $state.ExecutedQuery = $null
    $state.CurrentResult = $null
    $state.CurrentPage = 0
    $state.IsQueryStale = $false

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

function Show-SqlUtilityPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)] $PageResult
    )

    $state = $Form.Tag
    $resultsGrid = Get-SqlUtilityNamedControl -Root $Form -Name 'ResultsGrid'
    $previousButton = Get-SqlUtilityNamedControl -Root $Form -Name 'PreviousPageButton'
    $nextButton = Get-SqlUtilityNamedControl -Root $Form -Name 'NextPageButton'
    $exportButton = Get-SqlUtilityNamedControl -Root $Form -Name 'ExportButton'
    $pageStatus = Get-SqlUtilityNamedControl -Root $Form -Name 'PageStatusLabel'

    $resultsGrid.DataSource = $null
    $resultsGrid.Columns.Clear()
    $resultsGrid.AutoGenerateColumns = $true
    $resultsGrid.DataSource = $PageResult.Data
    $previousButton.Enabled = (-not $state.IsQueryStale) -and [bool] $PageResult.HasPrevious
    $nextButton.Enabled = (-not $state.IsQueryStale) -and [bool] $PageResult.HasNext
    $pageStatus.Text = 'Page {0} - {1} rows' -f $PageResult.PageNumber, $PageResult.DisplayedRowCount

    $canExport = $false
    if (-not $state.IsQueryStale -and $null -ne $state.ExecutedQuery) {
        $canExport = [bool] $state.ExecutedQuery.HasOrderBy -or `
            ([bool] $PageResult.IsComplete -and -not [bool] $PageResult.IsTruncated)
    }
    $exportButton.Enabled = $canExport

    foreach ($column in $resultsGrid.Columns) {
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        $resultsGrid.AutoResizeColumn($column.Index, [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
        $width = [Math]::Min(200, $column.Width)
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
        $column.Width = $width
    }
}

function Invoke-SqlUtilityQueryAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    if ($state.IsBusy) {
        return
    }

    $sqlEditor = Get-SqlUtilityNamedControl -Root $Form -Name 'SqlEditor'
    $editorSql = [string] $sqlEditor.Text
    Clear-SqlUtilityQueryResult -Form $Form
    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Executing query...'
    try {
        $validateQuery = $state.Services['ValidateQuery']
        $validation = & $validateQuery $editorSql
        if (-not [bool] $validation.IsValid) {
            Show-SqlUtilityMessage -State $state -Text ([string] $validation.ErrorMessage) `
                -Caption 'Query Validation' -Icon 'Warning'
            return
        }

        $snapshot = [pscustomobject][ordered]@{
            OriginalEditorSql = $editorSql
            NormalizedSql = [string] $validation.NormalizedSql
            HasOrderBy = [bool] $validation.HasOrderBy
        }
        if ($snapshot.HasOrderBy) {
            $executeOrderedPage = $state.Services['ExecuteOrderedPage']
            $pageResult = & $executeOrderedPage $state.ActiveServer $state.ActiveDatabase `
                $snapshot.NormalizedSql 1 $state.Config.queryExportTimeoutSeconds
        }
        else {
            $executeUnordered = $state.Services['ExecuteUnordered']
            $pageResult = & $executeUnordered $state.ActiveServer $state.ActiveDatabase `
                $snapshot.NormalizedSql $state.Config.unorderedRowLimit $state.Config.queryExportTimeoutSeconds
        }

        $state.ExecutedQuery = $snapshot
        $state.CurrentResult = $pageResult
        $state.CurrentPage = [int] $pageResult.PageNumber
        $state.IsQueryStale = $false
        Show-SqlUtilityPage -Form $Form -PageResult $pageResult
        if ([bool] $pageResult.IsTruncated) {
            Show-SqlUtilityMessage -State $state `
                -Text 'The result reached the configured maximum. Add ORDER BY for complete server-side paging.' `
                -Caption 'Result Limit Reached' -Icon 'Warning'
        }
    }
    catch {
        Clear-SqlUtilityQueryResult -Form $Form
        Show-SqlUtilityMessage -State $state `
            -Text ("Query failed.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Query Error' -Icon 'Error'
    }
    finally {
        Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
    }
}

function Invoke-SqlUtilityPageAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)][ValidateSet(-1, 1)][int] $PageDelta
    )

    $state = $Form.Tag
    if ($state.IsBusy -or $state.IsQueryStale -or $null -eq $state.ExecutedQuery -or $null -eq $state.CurrentResult) {
        return
    }

    $targetPage = $state.CurrentPage + $PageDelta
    if ($targetPage -lt 1) {
        return
    }

    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Loading page...'
    try {
        if ([bool] $state.ExecutedQuery.HasOrderBy) {
            $executeOrderedPage = $state.Services['ExecuteOrderedPage']
            $pageResult = & $executeOrderedPage $state.ActiveServer $state.ActiveDatabase `
                $state.ExecutedQuery.NormalizedSql $targetPage $state.Config.queryExportTimeoutSeconds
        }
        else {
            $getLocalPage = $state.Services['GetLocalPage']
            $pageResult = & $getLocalPage $state.CurrentResult.CachedData $targetPage `
                $state.CurrentResult.IsComplete $state.CurrentResult.IsTruncated
        }

        $state.CurrentResult = $pageResult
        $state.CurrentPage = [int] $pageResult.PageNumber
        Show-SqlUtilityPage -Form $Form -PageResult $pageResult
    }
    catch {
        Clear-SqlUtilityQueryResult -Form $Form
        Show-SqlUtilityMessage -State $state `
            -Text ("Query failed.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Query Error' -Icon 'Error'
    }
    finally {
        Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
    }
}

function Invoke-SqlUtilityExportAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    if ($state.IsBusy -or $state.IsQueryStale -or $null -eq $state.CurrentResult -or $null -eq $state.ExecutedQuery) {
        return
    }

    $busyStarted = $false
    try {
        $promptSavePath = $state.Services['PromptSavePath']
        $destinationPath = & $promptSavePath
        if ([string]::IsNullOrWhiteSpace([string] $destinationPath)) {
            return
        }

        Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Exporting result...'
        $busyStarted = $true
        $exportResult = $state.Services['ExportResult']
        & $exportResult $state ([string] $destinationPath)
        Show-SqlUtilityMessage -State $state -Text 'Export completed.' -Caption 'Export' -Icon 'Information'
    }
    catch {
        Show-SqlUtilityMessage -State $state `
            -Text ("Export failed.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Export Error' -Icon 'Error'
    }
    finally {
        if ($busyStarted) {
            Set-SqlUtilityBusy -Form $Form -Busy $false -Message 'Ready.'
        }
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
    Clear-SqlUtilityQueryResult -Form $Form

    foreach ($textBoxName in @('ServerTextBox', 'DatabaseTextBox', 'SqlEditor')) {
        $textBox = Get-SqlUtilityNamedControl -Root $Form -Name $textBoxName
        if ($null -ne $textBox) {
            $textBox.Text = ''
        }
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

    $querySplit = [System.Windows.Forms.SplitContainer]::new()
    $querySplit.Name = 'QuerySplitContainer'
    $querySplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $querySplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $querySplit.SplitterDistance = 205
    $querySplit.Panel1MinSize = 150
    $querySplit.Panel2MinSize = 120
    $queryTab.Controls.Add($querySplit)

    $queryActionPanel = [System.Windows.Forms.Panel]::new()
    $queryActionPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $queryActionPanel.Height = 64
    $querySplit.Panel1.Controls.Add($queryActionPanel)

    $executeButton = [System.Windows.Forms.Button]::new()
    $executeButton.Name = 'ExecuteButton'
    $executeButton.Text = 'Execute'
    $executeButton.AutoSize = $true
    $executeButton.Location = [System.Drawing.Point]::new(8, 5)
    $queryActionPanel.Controls.Add($executeButton)

    $exportButton = [System.Windows.Forms.Button]::new()
    $exportButton.Name = 'ExportButton'
    $exportButton.Text = 'Export to Excel'
    $exportButton.AutoSize = $true
    $exportButton.Enabled = $false
    $exportButton.Location = [System.Drawing.Point]::new(96, 5)
    $queryActionPanel.Controls.Add($exportButton)

    $pagingHelp = [System.Windows.Forms.Label]::new()
    $pagingHelp.Text = 'Paging requires ORDER BY on stable, preferably unique columns.'
    $pagingHelp.AutoSize = $true
    $pagingHelp.Location = [System.Drawing.Point]::new(225, 10)
    $queryActionPanel.Controls.Add($pagingHelp)

    $previousPageButton = [System.Windows.Forms.Button]::new()
    $previousPageButton.Name = 'PreviousPageButton'
    $previousPageButton.Text = 'Previous'
    $previousPageButton.AutoSize = $true
    $previousPageButton.Enabled = $false
    $previousPageButton.Location = [System.Drawing.Point]::new(8, 35)
    $queryActionPanel.Controls.Add($previousPageButton)

    $nextPageButton = [System.Windows.Forms.Button]::new()
    $nextPageButton.Name = 'NextPageButton'
    $nextPageButton.Text = 'Next'
    $nextPageButton.AutoSize = $true
    $nextPageButton.Enabled = $false
    $nextPageButton.Location = [System.Drawing.Point]::new(96, 35)
    $queryActionPanel.Controls.Add($nextPageButton)

    $pageStatusLabel = [System.Windows.Forms.Label]::new()
    $pageStatusLabel.Name = 'PageStatusLabel'
    $pageStatusLabel.Text = ''
    $pageStatusLabel.AutoSize = $true
    $pageStatusLabel.Location = [System.Drawing.Point]::new(175, 40)
    $queryActionPanel.Controls.Add($pageStatusLabel)

    $sqlEditor = [System.Windows.Forms.TextBox]::new()
    $sqlEditor.Name = 'SqlEditor'
    $sqlEditor.Multiline = $true
    $sqlEditor.AcceptsReturn = $true
    $sqlEditor.AcceptsTab = $true
    $sqlEditor.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $sqlEditor.WordWrap = $false
    $sqlEditor.Dock = [System.Windows.Forms.DockStyle]::Fill
    $sqlEditor.Font = [System.Drawing.Font]::new('Consolas', 10)
    $querySplit.Panel1.Controls.Add($sqlEditor)
    $sqlEditor.BringToFront()

    $resultsGrid = [System.Windows.Forms.DataGridView]::new()
    $resultsGrid.Name = 'ResultsGrid'
    $resultsGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $resultsGrid.ReadOnly = $true
    $resultsGrid.AllowUserToAddRows = $false
    $resultsGrid.AllowUserToDeleteRows = $false
    $resultsGrid.AllowUserToOrderColumns = $false
    $resultsGrid.AutoGenerateColumns = $true
    $resultsGrid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $resultsGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $querySplit.Panel2.Controls.Add($resultsGrid)

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
    $executeButton.Add_Click({ Invoke-SqlUtilityQueryAction -Form $form }.GetNewClosure())
    $previousPageButton.Add_Click({ Invoke-SqlUtilityPageAction -Form $form -PageDelta -1 }.GetNewClosure())
    $nextPageButton.Add_Click({ Invoke-SqlUtilityPageAction -Form $form -PageDelta 1 }.GetNewClosure())
    $exportButton.Add_Click({ Invoke-SqlUtilityExportAction -Form $form }.GetNewClosure())
    $sqlEditor.Add_TextChanged({
        if ($null -ne $form.Tag.ExecutedQuery -and -not $form.Tag.IsQueryStale) {
            $matchesExecutedText = [string]::Equals(
                [string] $sqlEditor.Text,
                [string] $form.Tag.ExecutedQuery.OriginalEditorSql,
                [System.StringComparison]::Ordinal
            )
            if (-not $matchesExecutedText) {
                $form.Tag.IsQueryStale = $true
                $previousPageButton.Enabled = $false
                $nextPageButton.Enabled = $false
                $exportButton.Enabled = $false
            }
        }
    }.GetNewClosure())

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
