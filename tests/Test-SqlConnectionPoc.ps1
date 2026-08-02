$ErrorActionPreference = 'Stop'
$script:Failures = 0

function Assert-True([bool] $Condition, [string] $Message) {
    if (-not $Condition) {
        $script:Failures++
        Write-Error -ErrorAction Continue "FAIL: $Message"
    }
}

function Assert-Equal($Expected, $Actual, [string] $Message) {
    Assert-True ($Expected -eq $Actual) "$Message (expected '$Expected', got '$Actual')"
}

function Assert-Throws([scriptblock] $Action, [string] $Message) {
    $threw = $false
    try { & $Action } catch { $threw = $true }
    Assert-True $threw $Message
}

& {
$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $projectRoot 'SqlConnectionPoc.ps1') -NoGui

Assert-Throws { New-PocConnectionString -Server '' -Database 'db' } 'Blank server is rejected'
Assert-Throws { New-PocConnectionString -Server 'server' -Database '  ' } 'Blank database is rejected'

$connectionString = New-PocConnectionString -Server 'sample-host' -Database 'sample-db'
$builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $connectionString
Assert-Equal 'sample-host' $builder.DataSource 'Server is copied into DataSource'
Assert-Equal 'sample-db' $builder.InitialCatalog 'Database is copied into InitialCatalog'
Assert-True $builder.IntegratedSecurity 'Integrated security is enabled'
Assert-Equal '' $builder.UserID 'No SQL user is present'
Assert-Equal '' $builder.Password 'No SQL password is present'
Assert-Equal 10 $builder.ConnectTimeout 'Connection timeout is bounded'

$expectedQuery = 'SELECT @@SERVERNAME AS ServerName, DB_NAME() AS DatabaseName, SYSTEM_USER AS LoginName, GETDATE() AS ServerTime;'
$actualQuery = ((Get-PocDiagnosticQuery) -replace '\s+', ' ').Trim()
Assert-Equal $expectedQuery $actualQuery 'Only the approved diagnostic query is returned'

function Find-ControlByName([System.Windows.Forms.Control] $Parent, [string] $Name) {
    foreach ($control in $Parent.Controls) {
        if ($control.Name -eq $Name) { return $control }
        $nested = Find-ControlByName -Parent $control -Name $Name
        if ($null -ne $nested) { return $nested }
    }
    return $null
}

$form = New-PocMainForm
try {
    $form.Show()
    $form.CreateControl()

    $serverBox = Find-ControlByName $form 'ServerTextBox'
    $databaseBox = Find-ControlByName $form 'DatabaseTextBox'
    $runButton = Find-ControlByName $form 'RunButton'
    $statusLabel = Find-ControlByName $form 'StatusLabel'
    $resultsGrid = Find-ControlByName $form 'ResultsGrid'

    Assert-True ($null -ne $serverBox) 'Server field exists'
    Assert-True ($null -ne $databaseBox) 'Database field exists'
    Assert-True ($null -ne $runButton) 'Run button exists'
    Assert-True ($null -ne $statusLabel) 'Status label exists'
    Assert-True ($null -ne $resultsGrid) 'Results grid exists'
    Assert-Equal ([System.Windows.Forms.Button]) $runButton.GetType() 'Run button uses the standard WinForms button implementation'
    Assert-Equal '' $serverBox.Text 'Server starts blank'
    Assert-Equal '' $databaseBox.Text 'Database starts blank'
    Assert-True $resultsGrid.ReadOnly 'Result editing is disabled'
    Assert-True (-not $resultsGrid.AllowUserToAddRows) 'Result row insertion is disabled'
    Assert-True (-not $resultsGrid.AllowUserToDeleteRows) 'Result row deletion is disabled'
    $runButton.PerformClick()
    Assert-True ($statusLabel.Text -like 'Failed: Enter a SQL Server name.*') 'Button click validates blank input without contacting SQL Server'
}
finally {
    $form.Dispose()
}

$blankDatabaseRunnerState = @{ CallCount = 0 }
$blankDatabaseRunner = {
    param([string] $Server, [string] $Database)
    $blankDatabaseRunnerState.CallCount++
    throw 'The query runner must not be called for a blank database.'
}.GetNewClosure()

$blankDatabaseForm = New-PocMainForm -QueryRunner $blankDatabaseRunner
try {
    $blankDatabaseForm.Show()
    $blankDatabaseForm.CreateControl()

    $blankDatabaseServerBox = Find-ControlByName $blankDatabaseForm 'ServerTextBox'
    $blankDatabaseBox = Find-ControlByName $blankDatabaseForm 'DatabaseTextBox'
    $blankDatabaseRunButton = Find-ControlByName $blankDatabaseForm 'RunButton'
    $blankDatabaseStatusLabel = Find-ControlByName $blankDatabaseForm 'StatusLabel'

    $blankDatabaseServerBox.Text = 'sample-host'
    $blankDatabaseBox.Text = '  '
    $blankDatabaseRunButton.PerformClick()

    Assert-Equal 0 $blankDatabaseRunnerState.CallCount 'Blank database validation does not call the query runner'
    Assert-True ($blankDatabaseStatusLabel.Text -like 'Failed: Enter a database name.*') 'Button click validates a blank database'
    Assert-True $blankDatabaseRunButton.Enabled 'Blank database validation restores the run button'
    Assert-Equal ([System.Windows.Forms.Cursors]::Default) $blankDatabaseForm.Cursor 'Blank database validation restores the cursor'
}
finally {
    $blankDatabaseForm.Dispose()
}

$successfulRunnerState = @{
    CallCount = 0
    Server = $null
    Database = $null
    Table = $null
}
$successfulRunner = {
    param([string] $Server, [string] $Database)

    $successfulRunnerState.CallCount++
    $successfulRunnerState.Server = $Server
    $successfulRunnerState.Database = $Database

    $table = New-Object System.Data.DataTable
    [void] $table.Columns.Add('ServerName', [string])
    [void] $table.Columns.Add('DatabaseName', [string])
    [void] $table.Columns.Add('LoginName', [string])
    [void] $table.Columns.Add('ServerTime', [datetime])
    [void] $table.Rows.Add('sample-server', 'sample-database', 'DOMAIN\sample-user', [datetime] '2026-08-02T10:30:00')
    $successfulRunnerState.Table = $table
    return (, $table)
}.GetNewClosure()

$successfulForm = New-PocMainForm -QueryRunner $successfulRunner
try {
    $successfulForm.Show()
    $successfulForm.CreateControl()

    $successfulServerBox = Find-ControlByName $successfulForm 'ServerTextBox'
    $successfulDatabaseBox = Find-ControlByName $successfulForm 'DatabaseTextBox'
    $successfulRunButton = Find-ControlByName $successfulForm 'RunButton'
    $successfulStatusLabel = Find-ControlByName $successfulForm 'StatusLabel'
    $successfulResultsGrid = Find-ControlByName $successfulForm 'ResultsGrid'

    $successfulServerBox.Text = 'sample-host'
    $successfulDatabaseBox.Text = 'sample-db'
    $successfulRunButton.PerformClick()

    Assert-Equal 1 $successfulRunnerState.CallCount 'Valid input calls the query runner once'
    Assert-Equal 'sample-host' $successfulRunnerState.Server 'Valid input passes the server to the query runner'
    Assert-Equal 'sample-db' $successfulRunnerState.Database 'Valid input passes the database to the query runner'
    $boundTable = $successfulResultsGrid.DataSource -as [System.Data.DataTable]
    $returnedTableIsBound = ($null -ne $successfulRunnerState.Table) -and [object]::ReferenceEquals($successfulRunnerState.Table, $boundTable)
    Assert-True $returnedTableIsBound 'The returned in-memory DataTable is bound to the results grid'
    $boundColumnCount = if ($null -eq $boundTable) { -1 } else { $boundTable.Columns.Count }
    $boundColumnNames = if ($null -eq $boundTable) { '' } else { ($boundTable.Columns | ForEach-Object { $_.ColumnName }) -join ',' }
    Assert-Equal 4 $boundColumnCount 'The bound diagnostic result has four columns'
    Assert-Equal 'ServerName,DatabaseName,LoginName,ServerTime' $boundColumnNames 'The bound diagnostic columns are correct'
    Assert-Equal 'Success: connected with Windows authentication and returned the diagnostic result.' $successfulStatusLabel.Text 'Successful execution updates the status'
    Assert-True $successfulRunButton.Enabled 'Successful execution restores the run button'
    Assert-Equal ([System.Windows.Forms.Cursors]::Default) $successfulForm.Cursor 'Successful execution restores the cursor'
}
finally {
    $successfulForm.Dispose()
}

function Get-WorkspaceFileState([string] $Path) {
    return @(
        Get-ChildItem -LiteralPath $Path -File -Recurse |
            Sort-Object FullName |
            ForEach-Object { '{0}|{1}|{2}' -f $_.FullName, $_.Length, $_.LastWriteTimeUtc.Ticks }
    )
}

$beforeState = @(Get-WorkspaceFileState $projectRoot)
$persistenceForm = New-PocMainForm
$persistenceForm.Dispose()
$afterState = @(Get-WorkspaceFileState $projectRoot)
Assert-Equal ($beforeState -join "`n") ($afterState -join "`n") 'Creating and closing the POC does not create or modify workspace files'
}

if ($script:Failures -gt 0) { exit 1 }
Write-Host 'All SQL Connection POC tests passed.'
