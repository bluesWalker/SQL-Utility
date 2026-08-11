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
        BuildCountSql = {
            param($CountSourceSql, $OutputColumnCount)
            return New-SqlUtilityCountSql -CountSourceSql $CountSourceSql -OutputColumnCount $OutputColumnCount
        }
        ExecuteCount = {
            param($Server, $Database, $CountSql, $TimeoutSeconds)
            return Invoke-SqlUtilityExactCount -Server $Server -Database $Database -CountSql $CountSql `
                -CommandTimeoutSeconds $TimeoutSeconds
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
        'ExecuteOrderedPage', 'ExecuteUnordered', 'GetLocalPage', 'BuildCountSql', 'ExecuteCount',
        'ExportResult', 'PromptSavePath'
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
    Update-SqlUtilityQueryActionState -Form $Form
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
    $state.ExplicitTotalRowCount = $null

    $resultsGrid = Get-SqlUtilityNamedControl -Root $Form -Name 'ResultsGrid'
    if ($null -ne $resultsGrid) {
        $resultsGrid.DataSource = $null
        $resultsGrid.Rows.Clear()
        $resultsGrid.Columns.Clear()
    }
    foreach ($buttonName in @('ExportButton', 'CountButton', 'PreviousPageButton', 'NextPageButton')) {
        $button = Get-SqlUtilityNamedControl -Root $Form -Name $buttonName
        if ($null -ne $button) {
            $button.Enabled = $false
        }
    }
    $pageStatus = Get-SqlUtilityNamedControl -Root $Form -Name 'PageStatusLabel'
    if ($null -ne $pageStatus) {
        $pageStatus.Text = ''
    }
    Update-SqlUtilityQueryActionState -Form $Form
}

function Get-SqlUtilityPageStatusText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $State,
        [Parameter(Mandatory = $true)] $PageResult
    )

    $pageNumber = [int] $PageResult.PageNumber
    $displayedRowCount = [int] $PageResult.DisplayedRowCount
    $prefix = [string]::Format(
        [System.Globalization.CultureInfo]::InvariantCulture,
        'Page {0} - {1}',
        $pageNumber,
        $displayedRowCount
    )

    $explicitTotal = $null
    if ($null -ne $State.PSObject.Properties['ExplicitTotalRowCount']) {
        $explicitTotal = $State.ExplicitTotalRowCount
    }
    if ($null -ne $explicitTotal) {
        return '{0} of {1}' -f $prefix, ([long] $explicitTotal).ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    $hasOrderBy = $false
    if ($null -ne $State.PSObject.Properties['ExecutedQuery'] -and $null -ne $State.ExecutedQuery) {
        $hasOrderBy = [bool] $State.ExecutedQuery.HasOrderBy
    }
    if (-not $hasOrderBy -and [bool] $PageResult.IsComplete -and -not [bool] $PageResult.IsTruncated -and
        $null -ne $PageResult.CachedData) {
        return '{0} of {1}' -f $prefix, ([int] $PageResult.CachedData.Rows.Count).ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if (-not $hasOrderBy -and [bool] $PageResult.IsTruncated) {
        return '{0} of {1}+' -f $prefix, ([int] $State.Config.unorderedRowLimit).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($hasOrderBy -and $pageNumber -eq 1 -and $displayedRowCount -eq 0) {
        return '{0} of 0' -f $prefix
    }
    return $prefix
}

function Update-SqlUtilityQueryActionState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    $previousButton = Get-SqlUtilityNamedControl -Root $Form -Name 'PreviousPageButton'
    $nextButton = Get-SqlUtilityNamedControl -Root $Form -Name 'NextPageButton'
    $exportButton = Get-SqlUtilityNamedControl -Root $Form -Name 'ExportButton'
    $countButton = Get-SqlUtilityNamedControl -Root $Form -Name 'CountButton'

    $hasFreshResult = -not $state.IsBusy -and -not $state.IsQueryStale -and `
        $null -ne $state.ExecutedQuery -and $null -ne $state.CurrentResult
    $exactCacheKnown = $hasFreshResult -and -not [bool] $state.ExecutedQuery.HasOrderBy -and `
        [bool] $state.CurrentResult.IsComplete -and -not [bool] $state.CurrentResult.IsTruncated

    if ($null -ne $previousButton) {
        $previousButton.Enabled = $hasFreshResult -and [bool] $state.CurrentResult.HasPrevious
    }
    if ($null -ne $nextButton) {
        $nextButton.Enabled = $hasFreshResult -and [bool] $state.CurrentResult.HasNext
    }
    if ($null -ne $exportButton) {
        $exportButton.Enabled = $hasFreshResult -and (
            [bool] $state.ExecutedQuery.HasOrderBy -or
            ([bool] $state.CurrentResult.IsComplete -and -not [bool] $state.CurrentResult.IsTruncated)
        )
    }
    if ($null -ne $countButton) {
        $countButton.Enabled = $hasFreshResult -and -not $exactCacheKnown
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
    $pageStatus = Get-SqlUtilityNamedControl -Root $Form -Name 'PageStatusLabel'

    $resultsGrid.DataSource = $null
    $resultsGrid.Columns.Clear()
    $resultsGrid.AutoGenerateColumns = $true
    $resultsGrid.DataSource = $PageResult.Data
    $pageStatus.Text = Get-SqlUtilityPageStatusText -State $state -PageResult $PageResult
    Update-SqlUtilityQueryActionState -Form $Form

    foreach ($column in $resultsGrid.Columns) {
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        $resultsGrid.AutoResizeColumn($column.Index, [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
        $width = [Math]::Min(300, $column.Width)
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

        $normalizedSql = [string] $validation.NormalizedSql
        $hasOrderBy = [bool] $validation.HasOrderBy
        if ($hasOrderBy) {
            $executeOrderedPage = $state.Services['ExecuteOrderedPage']
            $pageResult = & $executeOrderedPage $state.ActiveServer $state.ActiveDatabase `
                $normalizedSql 1 $state.Config.queryExportTimeoutSeconds
        }
        else {
            $executeUnordered = $state.Services['ExecuteUnordered']
            $pageResult = & $executeUnordered $state.ActiveServer $state.ActiveDatabase `
                $normalizedSql $state.Config.unorderedRowLimit $state.Config.queryExportTimeoutSeconds
        }

        $buildCountSql = $state.Services['BuildCountSql']
        $countSql = & $buildCountSql ([string] $validation.CountSourceSql) ([int] $pageResult.Data.Columns.Count)
        $snapshot = [pscustomobject][ordered]@{
            OriginalEditorSql = $editorSql
            NormalizedSql = $normalizedSql
            HasOrderBy = $hasOrderBy
            CountSql = [string] $countSql
        }

        $state.ExecutedQuery = $snapshot
        $state.CurrentResult = $pageResult
        $state.CurrentPage = [int] $pageResult.PageNumber
        $state.IsQueryStale = $false
        $state.ExplicitTotalRowCount = $null
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

function Invoke-SqlUtilityCountAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Form] $Form
    )

    $state = $Form.Tag
    $countButton = Get-SqlUtilityNamedControl -Root $Form -Name 'CountButton'
    $exactCacheKnown = $false
    if ($null -ne $state.ExecutedQuery -and $null -ne $state.CurrentResult) {
        $exactCacheKnown = -not [bool] $state.ExecutedQuery.HasOrderBy -and `
            [bool] $state.CurrentResult.IsComplete -and -not [bool] $state.CurrentResult.IsTruncated
    }
    if ($null -eq $countButton -or -not $countButton.Enabled -or $state.IsBusy -or $state.IsQueryStale -or
        $null -eq $state.ExecutedQuery -or $null -eq $state.CurrentResult -or $exactCacheKnown) {
        return
    }

    $priorExplicitTotal = $state.ExplicitTotalRowCount
    Set-SqlUtilityBusy -Form $Form -Busy $true -Message 'Counting rows...'
    try {
        $executeCount = $state.Services['ExecuteCount']
        $total = & $executeCount $state.ActiveServer $state.ActiveDatabase $state.ExecutedQuery.CountSql `
            $state.Config.queryExportTimeoutSeconds
        $state.ExplicitTotalRowCount = [long] $total
        $pageStatus = Get-SqlUtilityNamedControl -Root $Form -Name 'PageStatusLabel'
        $pageStatus.Text = Get-SqlUtilityPageStatusText -State $state -PageResult $state.CurrentResult
    }
    catch {
        $state.ExplicitTotalRowCount = $priorExplicitTotal
        Show-SqlUtilityMessage -State $state `
            -Text ("The row count could not be retrieved.`r`n`r`n{0}" -f $_.Exception.Message) `
            -Caption 'Row Count Failed' -Icon 'Error'
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
    if (
        -not [bool] $state.ExecutedQuery.HasOrderBy -and
        (-not [bool] $state.CurrentResult.IsComplete -or [bool] $state.CurrentResult.IsTruncated)
    ) {
        Show-SqlUtilityMessage -State $state `
            -Text 'This unordered result is incomplete and cannot be exported. Add ORDER BY and run the query again for a complete export.' `
            -Caption 'Export Unavailable' -Icon 'Warning'
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

function Initialize-SqlUtilityVisualStyles {
    if (-not (Get-Variable -Name SqlUtilityVisualStylesInitialized -Scope Script -ErrorAction SilentlyContinue) -or
        -not $script:SqlUtilityVisualStylesInitialized) {
        [System.Windows.Forms.Application]::EnableVisualStyles()
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
        $script:SqlUtilityVisualStylesInitialized = $true
    }
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
    Initialize-SqlUtilityVisualStyles

    $form = [System.Windows.Forms.Form]::new()
    $ownedFonts = [System.Collections.Generic.List[System.Drawing.Font]]::new()
    $interfaceFont = [System.Drawing.Font]::new('Segoe UI', 9.0)
    [void] $ownedFonts.Add($interfaceFont)
    $form.Font = $interfaceFont
    $form.Add_Disposed({
        foreach ($ownedFont in $ownedFonts) {
            $ownedFont.Dispose()
        }
        $ownedFonts.Clear()
    }.GetNewClosure())
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
        ExplicitTotalRowCount = $null
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
    $activeConnectionFont = [System.Drawing.Font]::new($interfaceFont, [System.Drawing.FontStyle]::Bold)
    [void] $ownedFonts.Add($activeConnectionFont)
    $activeConnectionLabel.Font = $activeConnectionFont
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
    $querySplit.FixedPanel = [System.Windows.Forms.FixedPanel]::None
    $querySplit.IsSplitterFixed = $false
    $queryTab.Controls.Add($querySplit)

    $queryActionLayout = [System.Windows.Forms.TableLayoutPanel]::new()
    $queryActionLayout.Name = 'QueryActionLayout'
    $queryActionLayout.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $queryActionLayout.AutoSize = $true
    $queryActionLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $queryActionLayout.BackColor = [System.Drawing.SystemColors]::Control
    $queryActionLayout.Padding = [System.Windows.Forms.Padding]::new(4)
    $queryActionLayout.ColumnCount = 7
    $queryActionLayout.RowCount = 1
    [void] $queryActionLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void] $queryActionLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    foreach ($columnIndex in 2..6) {
        [void] $queryActionLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    }
    [void] $queryActionLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    $querySplit.Panel1.Controls.Add($queryActionLayout)

    $executeButton = [System.Windows.Forms.Button]::new()
    $executeButton.Name = 'ExecuteButton'
    $executeButton.Text = 'Execute'
    $executeButton.AutoSize = $true
    $executeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $executeButton.Margin = [System.Windows.Forms.Padding]::new(0, 0, 6, 0)
    $executeButton.AccessibleName = 'Execute query'
    $executeButton.AccessibleDescription = 'Run the SQL query in the editor.'
    $executeButton.UseVisualStyleBackColor = $false
    $executeButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $executeButton.ForeColor = [System.Drawing.Color]::White
    $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 84, 153)
    $executeButton.FlatAppearance.BorderSize = 1
    $queryActionLayout.Controls.Add($executeButton, 0, 0)

    $pagingHelp = [System.Windows.Forms.Label]::new()
    $pagingHelp.Name = 'PagingHelpLabel'
    $pagingHelp.Text = 'ORDER BY required for paging.'
    $pagingHelp.AutoSize = $false
    $pagingHelp.AutoEllipsis = $true
    $pagingHelp.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pagingHelp.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $pagingHelp.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $pagingHelp.ForeColor = [System.Drawing.SystemColors]::ControlText
    $pagingHelp.AccessibleName = 'Paging requirement'
    $pagingHelp.AccessibleDescription = 'ORDER BY is required for result paging.'
    $queryActionLayout.Controls.Add($pagingHelp, 1, 0)

    $pageStatusLabel = [System.Windows.Forms.Label]::new()
    $pageStatusLabel.Name = 'PageStatusLabel'
    $pageStatusLabel.Text = ''
    $pageStatusLabel.AutoSize = $true
    $pageStatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $pageStatusLabel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $pageStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $pageStatusLabel.AccessibleName = 'Page status'
    $queryActionLayout.Controls.Add($pageStatusLabel, 2, 0)

    $countButton = [System.Windows.Forms.Button]::new()
    $countButton.Name = 'CountButton'
    $countButton.Text = 'Count'
    $countButton.AutoSize = $true
    $countButton.Enabled = $false
    $countButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $countButton.Margin = [System.Windows.Forms.Padding]::new(0, 0, 4, 0)
    $countButton.AccessibleName = 'Count result rows'
    $countButton.AccessibleDescription = 'Retrieve the exact row count for the current result.'
    $countButton.UseVisualStyleBackColor = $true
    $queryActionLayout.Controls.Add($countButton, 3, 0)

    $previousPageButton = [System.Windows.Forms.Button]::new()
    $previousPageButton.Name = 'PreviousPageButton'
    $previousPageButton.Text = '<'
    $previousPageButton.AutoSize = $true
    $previousPageButton.Enabled = $false
    $previousPageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $previousPageButton.Margin = [System.Windows.Forms.Padding]::new(0, 0, 4, 0)
    $previousPageButton.AccessibleName = 'Previous page'
    $previousPageButton.AccessibleDescription = 'Show the previous result page.'
    $previousPageButton.UseVisualStyleBackColor = $true
    $queryActionLayout.Controls.Add($previousPageButton, 4, 0)

    $nextPageButton = [System.Windows.Forms.Button]::new()
    $nextPageButton.Name = 'NextPageButton'
    $nextPageButton.Text = '>'
    $nextPageButton.AutoSize = $true
    $nextPageButton.Enabled = $false
    $nextPageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $nextPageButton.Margin = [System.Windows.Forms.Padding]::new(0, 0, 4, 0)
    $nextPageButton.AccessibleName = 'Next page'
    $nextPageButton.AccessibleDescription = 'Show the next result page.'
    $nextPageButton.UseVisualStyleBackColor = $true
    $queryActionLayout.Controls.Add($nextPageButton, 5, 0)

    $exportButton = [System.Windows.Forms.Button]::new()
    $exportButton.Name = 'ExportButton'
    $exportButton.Text = 'Export'
    $exportButton.AutoSize = $true
    $exportButton.Enabled = $false
    $exportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $exportButton.Margin = [System.Windows.Forms.Padding]::new(0)
    $exportButton.AccessibleName = 'Export to Excel'
    $exportButton.AccessibleDescription = 'Export the complete current result to an Excel workbook.'
    $exportButton.UseVisualStyleBackColor = $true
    $queryActionLayout.Controls.Add($exportButton, 6, 0)

    $sqlEditor = [System.Windows.Forms.TextBox]::new()
    $sqlEditor.Name = 'SqlEditor'
    $sqlEditor.Multiline = $true
    $sqlEditor.AcceptsReturn = $true
    $sqlEditor.AcceptsTab = $true
    $sqlEditor.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $sqlEditor.WordWrap = $false
    $sqlEditor.Dock = [System.Windows.Forms.DockStyle]::Fill
    $editorFont = [System.Drawing.Font]::new('Consolas', 10.0)
    [void] $ownedFonts.Add($editorFont)
    $sqlEditor.Font = $editorFont
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
    $resultsGrid.BackgroundColor = [System.Drawing.SystemColors]::Window
    $resultsGrid.GridColor = [System.Drawing.SystemColors]::ControlDark
    $resultsGrid.DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Window
    $resultsGrid.DefaultCellStyle.ForeColor = [System.Drawing.SystemColors]::WindowText
    $resultsGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
    $resultsGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.SystemColors]::ControlText
    $resultsGrid.EnableHeadersVisualStyles = $false
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
    $countButton.Add_Click({ Invoke-SqlUtilityCountAction -Form $form }.GetNewClosure())
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
                Update-SqlUtilityQueryActionState -Form $form
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
    Initialize-SqlUtilityVisualStyles

    try {
        $config = Read-SqlUtilityConfig -Path $ConfigPath
    }
    catch [System.NotSupportedException] {
        $showMessage = $Services['ShowMessage']
        & $showMessage $_.Exception.Message 'SQL Utility Configuration Error' 'Error'
        return
    }
    catch [System.IO.InvalidDataException] {
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
    catch {
        $showMessage = $Services['ShowMessage']
        & $showMessage ("The configuration file could not be read.`r`n`r`n{0}" -f $_.Exception.Message) 'SQL Utility Configuration Error' 'Error'
        return
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
