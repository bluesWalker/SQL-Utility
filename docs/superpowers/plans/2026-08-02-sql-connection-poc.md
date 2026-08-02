# SQL Connection GUI POC Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a portable Windows PowerShell GUI that uses the current Windows identity to run one fixed diagnostic query against a user-entered SQL Server and database and display the result without persisting runtime data.

**Architecture:** `SqlConnectionPoc.ps1` owns three focused units: safe connection-string construction, fixed query execution into an in-memory `DataTable`, and WinForms presentation. `StartSqlPoc.cmd` is a location-independent launcher. A dependency-free PowerShell test harness verifies security, blank defaults, UI restrictions, and launcher behavior without contacting a database.

**Tech Stack:** Windows PowerShell 5.1; .NET Framework `System.Windows.Forms`, `System.Drawing`, and `System.Data.SqlClient`; Windows batch launcher.

## Global Constraints

- Target Windows PowerShell 5.1 in the Citrix environment.
- Use Windows integrated authentication only; never accept or include SQL credentials.
- Server and database fields must be blank on every launch and absent as connection defaults from source.
- Execute only the approved read-only diagnostic query.
- Do not create runtime files, settings, history, logs, exports, cache entries, registry entries, or temporary files.
- Require no installer, compilation, administrator rights, third-party module, NuGet package, Python, Office, or `sqlcmd`.
- Keep the two runtime deliverables portable: `StartSqlPoc.cmd` and `SqlConnectionPoc.ps1`.
- The workspace is not currently a Git repository. Commit steps below must report a skip unless a repository exists when executed; they must not initialize one without user authorization.
- The launcher is a four-line configuration artifact; the user approved focused review and manual execution in place of a source-text test. All executable PowerShell behavior remains test-first.

## File Structure

- Create `SqlConnectionPoc.ps1`: connection-string builder, fixed query provider, query executor, WinForms construction, and guarded GUI entry point.
- Create `StartSqlPoc.cmd`: double-click launcher that finds the PowerShell script beside itself and applies process-only execution-policy bypass.
- Create `tests/Test-SqlConnectionPoc.ps1`: built-in assertion harness; no Pester or other module dependency.
- Use `docs/superpowers/specs/2026-08-02-sql-connection-poc-design.md` as the acceptance source; do not alter it unless implementation exposes a genuine design conflict.

---

### Task 1: Secure connection and query core

**Files:**
- Create: `SqlConnectionPoc.ps1`
- Create: `tests/Test-SqlConnectionPoc.ps1`

**Interfaces:**
- Produces: `New-PocConnectionString([string] $Server, [string] $Database) -> [string]`
- Produces: `Get-PocDiagnosticQuery() -> [string]`
- Produces: `Invoke-PocDiagnosticQuery([string] $Server, [string] $Database) -> [System.Data.DataTable]`
- Consumes: .NET Framework `System.Data.SqlClient` types only.

- [ ] **Step 1: Write the failing dependency-free core tests**

Create `tests/Test-SqlConnectionPoc.ps1` with a small assertion harness and these exact behavioral checks:

```powershell
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

if ($script:Failures -gt 0) { exit 1 }
Write-Host 'All SQL Connection POC tests passed.'
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: nonzero exit because `SqlConnectionPoc.ps1` or `New-PocConnectionString` does not exist.

- [ ] **Step 3: Implement the minimal connection and query functions**

Create `SqlConnectionPoc.ps1` with a `-NoGui` switch, strict error behavior, required assembly loads, and these implementations:

```powershell
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
    $builder.DataSource = $Server.Trim()
    $builder.InitialCatalog = $Database.Trim()
    $builder.IntegratedSecurity = $true
    $builder.ApplicationName = 'SQL Connection POC'
    $builder.ConnectTimeout = 10
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
```

Do not add a GUI entry call yet; `-NoGui` exists so the test contract remains stable when the GUI is added.

- [ ] **Step 4: Run the core tests and verify they pass**

Run the Step 2 command again.

Expected: exit code `0` and `All SQL Connection POC tests passed.` No SQL Server connection is attempted because the executor is not called by the tests.

- [ ] **Step 5: Record the task checkpoint**

Run:

```powershell
if (Test-Path .git) {
    git add SqlConnectionPoc.ps1 tests/Test-SqlConnectionPoc.ps1
    git commit -m "feat: add secure SQL diagnostic core"
} else {
    Write-Host 'Git checkpoint skipped: workspace is not a repository.'
}
```

Expected in the current workspace: the explicit skip message and no repository initialization.

---

### Task 2: Minimal read-only WinForms interface

**Files:**
- Modify: `SqlConnectionPoc.ps1`
- Modify: `tests/Test-SqlConnectionPoc.ps1`

**Interfaces:**
- Consumes: `Invoke-PocDiagnosticQuery([string], [string]) -> [System.Data.DataTable]` from Task 1.
- Produces: `New-PocMainForm() -> [System.Windows.Forms.Form]`.
- Produces named controls for verification: `ServerTextBox`, `DatabaseTextBox`, `RunButton`, `StatusLabel`, and `ResultsGrid`.

- [ ] **Step 1: Extend the test harness with failing GUI checks**

Insert the following before the final failure check in `tests/Test-SqlConnectionPoc.ps1`. The helper recursively finds controls inside layout containers:

```powershell
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: nonzero exit with `New-PocMainForm` not recognized.

- [ ] **Step 3: Implement the GUI and guarded entry point**

Append `New-PocMainForm` to `SqlConnectionPoc.ps1`. Build a resizable form with a top `TableLayoutPanel`, two blank text boxes, the named button and status label, and a fill-docked named `DataGridView`. Configure the grid exactly as follows:

```powershell
$resultsGrid.Name = 'ResultsGrid'
$resultsGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$resultsGrid.ReadOnly = $true
$resultsGrid.AllowUserToAddRows = $false
$resultsGrid.AllowUserToDeleteRows = $false
$resultsGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$resultsGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
```

Name the blank fields and other controls exactly:

```powershell
$serverTextBox.Name = 'ServerTextBox'
$serverTextBox.Text = ''
$databaseTextBox.Name = 'DatabaseTextBox'
$databaseTextBox.Text = ''
$runButton.Name = 'RunButton'
$runButton.Text = 'Test connection and query'
$statusLabel.Name = 'StatusLabel'
$statusLabel.Text = 'Enter a server and database.'
```

Wire the button click to disable the button, clear the grid, set the wait cursor, and call:

```powershell
$table = Invoke-PocDiagnosticQuery -Server $serverTextBox.Text -Database $databaseTextBox.Text
$resultsGrid.DataSource = $table
$statusLabel.Text = 'Success: connected with Windows authentication and returned the diagnostic result.'
```

Catch exceptions and display `Failed: <exception message>` in `StatusLabel`; do not write or log the exception. In `finally`, restore the default cursor and re-enable the button. Call `$form.Refresh()` after setting the in-progress status so it is painted before the synchronous connection attempt.

End the script with the guarded entry point:

```powershell
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
```

- [ ] **Step 4: Run the complete core and GUI tests**

Run the Step 2 command again.

Expected: exit code `0` and the all-tests-passed message. No GUI remains open and no database connection is attempted.

- [ ] **Step 5: Manually smoke-test GUI startup without entering connection data**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\SqlConnectionPoc.ps1
```

Expected: a resizable window opens with both fields blank. Clicking the button with blank values reports `Failed: Enter a SQL Server name.` in the window, creates no files, and keeps the window usable. Close the window manually.

- [ ] **Step 6: Record the task checkpoint**

Run:

```powershell
if (Test-Path .git) {
    git add SqlConnectionPoc.ps1 tests/Test-SqlConnectionPoc.ps1
    git commit -m "feat: add SQL connection POC GUI"
} else {
    Write-Host 'Git checkpoint skipped: workspace is not a repository.'
}
```

Expected in the current workspace: the explicit skip message.

---

### Task 3: Portable launcher and behavioral persistence guardrail

**Files:**
- Create: `StartSqlPoc.cmd`
- Modify: `tests/Test-SqlConnectionPoc.ps1`

**Interfaces:**
- Consumes: `SqlConnectionPoc.ps1` beside the launcher.
- Produces: a double-clickable process-local launch command and propagated exit code.

- [ ] **Step 1: Add a failing runtime no-persistence test**

Insert these checks before the final failure check in `tests/Test-SqlConnectionPoc.ps1`. They exercise the real script and form, then compare observable workspace state rather than inspecting source text:

```powershell
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
```

- [ ] **Step 2: Run the test and verify the new guard fails for the expected reason**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Before adding the comparison helper, temporarily make `New-PocMainForm` write a sentinel filename only inside the test's controlled setup, confirm the assertion reports that workspace state changed, then remove the sentinel-producing setup before continuing. The application source must never contain that write. Expected final RED evidence: a nonzero exit with `FAIL: Creating and closing the POC does not create or modify workspace files` caused by the test-only sentinel.

- [ ] **Step 3: Create the minimal launcher**

Create `StartSqlPoc.cmd` with exactly:

```bat
@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0SqlConnectionPoc.ps1"
set "pocExitCode=%ERRORLEVEL%"
endlocal & exit /b %pocExitCode%
```

`-ExecutionPolicy Bypass` applies only to this PowerShell process; the launcher must not call `Set-ExecutionPolicy` or change machine/user policy.

Review the four lines directly against those requirements. This user-approved configuration-file exception avoids a brittle test that merely greps launcher source.

- [ ] **Step 4: Run the full automated verification**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlConnectionPoc.ps1
```

Expected: exit code `0` and `All SQL Connection POC tests passed.`

- [ ] **Step 5: Verify the two-file distribution boundary**

Run:

```powershell
Get-Item .\StartSqlPoc.cmd, .\SqlConnectionPoc.ps1 | Select-Object Name, Length
```

Expected: exactly the two runtime files are listed. The `tests` and `docs` directories are development artifacts and do not need to be copied to Citrix.

- [ ] **Step 6: Record the final checkpoint**

Run:

```powershell
if (Test-Path .git) {
    git add StartSqlPoc.cmd SqlConnectionPoc.ps1 tests/Test-SqlConnectionPoc.ps1
    git commit -m "feat: add portable SQL POC launcher"
} else {
    Write-Host 'Git checkpoint skipped: workspace is not a repository.'
}
```

Expected in the current workspace: the explicit skip message.

---

## Citrix Acceptance Test

After local implementation and verification:

1. Copy only `StartSqlPoc.cmd` and `SqlConnectionPoc.ps1` to the allowed cloud-drive folder.
2. Double-click `StartSqlPoc.cmd`. If command-file launch is blocked, open an allowed Windows PowerShell session and run `powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\SqlConnectionPoc.ps1` from that folder.
3. Confirm Server and Database are blank.
4. Enter the server alias and database name supplied through the existing operational process.
5. Click **Test connection and query**.
6. Accept only a grid containing `ServerName`, `DatabaseName`, `LoginName`, and `ServerTime`, with values supplied by the SQL Server session.
7. Close and reopen the POC; confirm both fields and the grid are blank again.
8. If the test fails, retain the on-screen error text for diagnosis; the POC itself will not create a log file.
