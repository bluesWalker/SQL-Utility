param([switch] $NoGui)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

function New-PocConnectionString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database
    )

    if ([string]::IsNullOrWhiteSpace($Server)) {
        throw [System.ArgumentException]::new('Enter a SQL Server name.')
    }
    if ([string]::IsNullOrWhiteSpace($Database)) {
        throw [System.ArgumentException]::new('Enter a database name.')
    }

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $Server.Trim()
    $builder['Initial Catalog'] = $Database.Trim()
    $builder['Integrated Security'] = $true
    $builder['Application Name'] = 'SQL Connection POC'
    $builder['Connect Timeout'] = 10
    return $builder.ConnectionString
}

function Get-PocDiagnosticQuery {
    return @'
SELECT
    @@SERVERNAME AS ServerName,
    DB_NAME() AS DatabaseName,
    SYSTEM_USER AS LoginName,
    GETDATE() AS ServerTime;
'@
}

function Invoke-PocDiagnosticQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Server,
        [Parameter(Mandatory = $true)][string] $Database
    )

    $connectionString = New-PocConnectionString -Server $Server -Database $Database
    $connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $command = $connection.CreateCommand()
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $table = New-Object System.Data.DataTable

    try {
        $command.CommandText = Get-PocDiagnosticQuery
        $command.CommandTimeout = 15
        $connection.Open()
        [void] $adapter.Fill($table)
        return (, $table)
    }
    finally {
        $adapter.Dispose()
        $command.Dispose()
        $connection.Dispose()
    }
}

function New-PocMainForm {
    param(
        [scriptblock] $QueryRunner = ${function:Invoke-PocDiagnosticQuery}
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'SQL Connection POC'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.MinimumSize = New-Object System.Drawing.Size(520, 320)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $topLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $topLayout.Dock = [System.Windows.Forms.DockStyle]::Top
    $topLayout.AutoSize = $true
    $topLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $topLayout.Padding = New-Object System.Windows.Forms.Padding(8)
    $topLayout.ColumnCount = 2
    $topLayout.RowCount = 4
    [void] $topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void] $topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $serverLabel = New-Object System.Windows.Forms.Label
    $serverLabel.Text = 'Server:'
    $serverLabel.AutoSize = $true
    $serverLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left

    $serverTextBox = New-Object System.Windows.Forms.TextBox
    $serverTextBox.Name = 'ServerTextBox'
    $serverTextBox.Text = ''
    $serverTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill

    $databaseLabel = New-Object System.Windows.Forms.Label
    $databaseLabel.Text = 'Database:'
    $databaseLabel.AutoSize = $true
    $databaseLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left

    $databaseTextBox = New-Object System.Windows.Forms.TextBox
    $databaseTextBox.Name = 'DatabaseTextBox'
    $databaseTextBox.Text = ''
    $databaseTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Name = 'RunButton'
    $runButton.Text = 'Test connection and query'
    $runButton.AutoSize = $true
    $runButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Name = 'StatusLabel'
    $statusLabel.Text = 'Enter a server and database.'
    $statusLabel.AutoSize = $true
    $statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left

    $resultsGrid = New-Object System.Windows.Forms.DataGridView
    $resultsGrid.Name = 'ResultsGrid'
    $resultsGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $resultsGrid.ReadOnly = $true
    $resultsGrid.AllowUserToAddRows = $false
    $resultsGrid.AllowUserToDeleteRows = $false
    $resultsGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $resultsGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

    [void] $topLayout.Controls.Add($serverLabel, 0, 0)
    [void] $topLayout.Controls.Add($serverTextBox, 1, 0)
    [void] $topLayout.Controls.Add($databaseLabel, 0, 1)
    [void] $topLayout.Controls.Add($databaseTextBox, 1, 1)
    [void] $topLayout.Controls.Add($runButton, 0, 2)
    [void] $topLayout.SetColumnSpan($runButton, 2)
    [void] $topLayout.Controls.Add($statusLabel, 0, 3)
    [void] $topLayout.SetColumnSpan($statusLabel, 2)

    $clickHandler = {
        $runButton.Enabled = $false
        $resultsGrid.DataSource = $null
        $resultsGrid.Rows.Clear()
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $statusLabel.Text = 'Connecting and running diagnostic query...'
        $form.Refresh()

        try {
            if ([string]::IsNullOrWhiteSpace($serverTextBox.Text)) {
                throw [System.ArgumentException]::new('Enter a SQL Server name.')
            }
            if ([string]::IsNullOrWhiteSpace($databaseTextBox.Text)) {
                throw [System.ArgumentException]::new('Enter a database name.')
            }
            $table = & $QueryRunner -Server $serverTextBox.Text -Database $databaseTextBox.Text
            $resultsGrid.DataSource = $table
            $statusLabel.Text = 'Success: connected with Windows authentication and returned the diagnostic result.'
        }
        catch {
            $statusLabel.Text = "Failed: $($_.Exception.Message)"
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $runButton.Enabled = $true
        }
    }.GetNewClosure()
    $runButton.Add_Click($clickHandler)

    [void] $form.Controls.Add($resultsGrid)
    [void] $form.Controls.Add($topLayout)
    return $form
}

if (-not $NoGui) {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $mainForm = New-PocMainForm
    try {
        [void] $mainForm.ShowDialog()
    }
    finally {
        $mainForm.Dispose()
    }
}
