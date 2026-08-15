# SQL Utility Column Selection Navigation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Data Explorer output-column checks change only through checkbox-glyph clicks or All/None, and add case-insensitive one-second buffered prefix navigation that changes only the highlighted column.

**Architecture:** Keep the standard `CheckedListBox` and all new behavior inside `SqlUtility.ps1`, which already owns WinForms state and orchestration. A transient form-owned interaction object gates application-owned check changes and holds the navigation prefix and timer; no builder, preview, configuration, database, or module contract changes.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework WinForms, the existing script-based UI test harness, and Git.

## Global Constraints

- Work only in `C:\Users\bluesWalker\Documents\SQL Utility\.worktrees\column-selection-navigation` on `feature/column-selection-navigation`.
- Preserve the seven-file runtime distribution, Windows integrated authentication, physical-table-only scope, preview/snapshot independence, and all module boundaries.
- Do not add a search bar, owner-drawn control, compiled helper, dependency, persistence, setting, query behavior, or database call.
- Use regression-first TDD for each behavior change. Run the focused UI suite after each production correction and the aggregate suite before completion.
- Treat checkbox hit-area feel/scaling and timeout feel in Citrix as pending external acceptance.

---

### Task 1: Allow check changes only through the checkbox glyph and All/None

**Files:**

- Modify: `tests/Test-SqlUtilityUi.ps1` (test helpers near lines 6-20; Data Explorer workflow near lines 1540-1725; preview-export workflow near lines 1730-1790)
- Modify: `SqlUtility.ps1` (Data Explorer helpers near lines 547-565; metadata population near line 625; form state near lines 963-985; output-list construction near line 1281; event wiring near lines 1390-1410)

- [ ] **Step 1: Add deterministic real-control input helpers to the UI test**

Add a general protected-event invoker and output-list click helper after `Get-TestControl`:

```powershell
function Invoke-TestProtectedControlEvent($Control, [string] $MethodName, [System.EventArgs] $EventArgs) {
    $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
    $method = $Control.GetType().GetMethod($MethodName, $flags)
    Assert-True ($null -ne $method) "$($Control.GetType().Name) exposes $MethodName"
    [void] $method.Invoke($Control, @($EventArgs))
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-TestOutputColumnClick($Form, [int] $Index, [switch] $Text) {
    $list = Get-TestControl $Form 'OutputColumnsList'
    $rectangle = $list.GetItemRectangle($Index)
    $checkWidth = [System.Windows.Forms.SystemInformation]::MenuCheckSize.Width
    $x = if ($Text) { $rectangle.Left + $checkWidth + 8 } else { $rectangle.Left + [Math]::Max(1, [Math]::Floor($checkWidth / 2)) }
    $y = $rectangle.Top + [Math]::Max(1, [Math]::Floor($rectangle.Height / 2))
    $eventArgs = [System.Windows.Forms.MouseEventArgs]::new(
        [System.Windows.Forms.MouseButtons]::Left, 1, $x, $y, 0
    )
    Invoke-TestProtectedControlEvent $list 'OnMouseDown' $eventArgs
}
```

Use `GetItemRectangle` and `SystemInformation.MenuCheckSize.Width`, matching the production DPI-aware hit-test contract. Keep the helper test-only.

- [ ] **Step 2: Write the failing checkbox-interaction regressions**

Immediately after the existing first Preview assertions, capture the exact checked-name set and assert:

```powershell
$outputList = Get-TestControl $explorerForm 'OutputColumnsList'
$initialCheckedNames = @($outputList.CheckedItems | ForEach-Object Name) -join ','
Assert-Equal $false $outputList.CheckOnClick 'Output text clicks cannot use native check-on-click behavior'

Invoke-TestOutputColumnClick $explorerForm 1 -Text
Invoke-TestOutputColumnClick $explorerForm 1 -Text
Assert-Equal $initialCheckedNames (@($outputList.CheckedItems | ForEach-Object Name) -join ',') `
    'Repeated output-column text clicks preserve every check'

Invoke-TestOutputColumnClick $explorerForm 1
Assert-Equal $false $outputList.GetItemChecked(1) 'Checkbox glyph click unchecks exactly one column'
Invoke-TestOutputColumnClick $explorerForm 1
Assert-Equal $true $outputList.GetItemChecked(1) 'Second checkbox glyph click checks exactly once'

$outputList.SetItemChecked(1, $false)
Assert-Equal $true $outputList.GetItemChecked(1) 'Unguarded programmatic or native check changes are rejected'
```

Replace the two existing direct `SetItemChecked(...)` calls that model user interaction with `Invoke-TestOutputColumnClick`. Add explicit All/None count assertions around the existing bulk-action and preview-snapshot checks.

- [ ] **Step 3: Run the focused UI test and confirm RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: exit `1`, first on `CheckOnClick` being `True` (or on a following text-click/check-gate assertion). Record the exact failing assertion before editing production code.

- [ ] **Step 4: Add the transient check-state gate and guarded check helpers**

Add this form-state member in `New-SqlUtilityMainForm`:

```powershell
DataExplorerOutputInteraction = [pscustomobject][ordered]@{
    AllowCheckChange = $false
    Prefix = ''
    ResetTimer = $null
}
```

Add focused UI helpers near the existing Data Explorer helpers:

```powershell
function Invoke-SqlUtilityDataExplorerOutputCheckChange {
    param(
        [System.Windows.Forms.Form] $Form,
        [scriptblock] $Action
    )

    $interaction = $Form.Tag.DataExplorerOutputInteraction
    $previous = [bool] $interaction.AllowCheckChange
    try {
        $interaction.AllowCheckChange = $true
        & $Action
    }
    finally {
        $interaction.AllowCheckChange = $previous
    }
}

function Set-SqlUtilityDataExplorerOutputColumnChecked {
    param(
        [System.Windows.Forms.Form] $Form,
        [int] $Index,
        [bool] $Checked
    )

    $outputList = Get-SqlUtilityNamedControl $Form 'OutputColumnsList'
    if ($Index -lt 0 -or $Index -ge $outputList.Items.Count) { return }
    Invoke-SqlUtilityDataExplorerOutputCheckChange $Form {
        $outputList.SetItemChecked($Index, $Checked)
    }.GetNewClosure()
    Update-SqlUtilityDataExplorerSendState $Form
}

function Set-SqlUtilityDataExplorerAllOutputColumnsChecked {
    param(
        [System.Windows.Forms.Form] $Form,
        [bool] $Checked
    )

    $outputList = Get-SqlUtilityNamedControl $Form 'OutputColumnsList'
    Invoke-SqlUtilityDataExplorerOutputCheckChange $Form {
        for ($index = 0; $index -lt $outputList.Items.Count; $index++) {
            $outputList.SetItemChecked($index, $Checked)
        }
    }.GetNewClosure()
    Update-SqlUtilityDataExplorerSendState $Form
}

function Invoke-SqlUtilityDataExplorerOutputColumnMouseDown {
    param(
        [System.Windows.Forms.Form] $Form,
        [System.Windows.Forms.MouseEventArgs] $EventArgs
    )

    if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $outputList = Get-SqlUtilityNamedControl $Form 'OutputColumnsList'
    $index = $outputList.IndexFromPoint($EventArgs.Location)
    if ($index -lt 0) { return }

    $rectangle = $outputList.GetItemRectangle($index)
    $checkRight = $rectangle.Left + [System.Windows.Forms.SystemInformation]::MenuCheckSize.Width
    if ($EventArgs.X -ge $rectangle.Left -and $EventArgs.X -lt $checkRight) {
        $outputList.SelectedIndex = $index
        Set-SqlUtilityDataExplorerOutputColumnChecked $Form $index (-not $outputList.GetItemChecked($index))
    }
}
```

Do not put this state in `DataExplorerBuilder` or `DataExplorerPreview`.

- [ ] **Step 5: Route every allowed and disallowed check path through the gate**

Make these exact orchestration changes:

- Set `$outputs.CheckOnClick = $false`.
- In the `ItemCheck` handler, set `$eventArgs.NewValue = $eventArgs.CurrentValue` and return unless `AllowCheckChange` is true; retain the deferred Send-state update only for allowed changes.
- Wire `MouseDown` to `Invoke-SqlUtilityDataExplorerOutputColumnMouseDown`.
- Replace the All/None loops with `Set-SqlUtilityDataExplorerAllOutputColumnsChecked`.
- Wrap the metadata-population loop that calls `Items.Add(..., $true)` in `Invoke-SqlUtilityDataExplorerOutputCheckChange`, retaining the current ordinal order, display text, default-all-selected behavior, and Preview execution.

The gate must restore its prior value in `finally`, even if a control operation fails.

- [ ] **Step 6: Run GREEN verification for Task 1**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
git diff --check
```

Expected: UI suite exits `0`; text clicks and direct/native check attempts preserve checks; glyph clicks toggle once; initial all-checked and All/None/snapshot assertions remain green; diff check reports no errors.

- [ ] **Step 7: Review and commit Task 1**

Inspect `git diff -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1`, confirm no query/builder/preview behavior widened, then commit only those two files:

```powershell
git add -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git diff --cached --check
git diff --cached --stat
git commit -m "fix: restrict output column check changes"
```

---

### Task 2: Add one-second buffered prefix navigation and reset its transient state

**Files:**

- Modify: `tests/Test-SqlUtilityUi.ps1` (input helper near the Task 1 helpers; new focused Data Explorer navigation fixture near the main Explorer workflow)
- Modify: `SqlUtility.ps1` (Data Explorer helpers near Task 1 helpers; list-clear paths near lines 620, 625, 881, and 1394; timer construction/disposal and event wiring in `New-SqlUtilityMainForm`)
- Modify: `README.md` (Data Explorer behavior near lines 169-186)

- [ ] **Step 1: Add a key-input helper and focused navigation fixture**

Add this test helper beside the protected-event helper:

```powershell
function Invoke-TestOutputColumnKeyPress($Form, [char] $Character) {
    $list = Get-TestControl $Form 'OutputColumnsList'
    $eventArgs = [System.Windows.Forms.KeyPressEventArgs]::new($Character)
    Invoke-TestProtectedControlEvent $list 'OnKeyPress' $eventArgs
    return $eventArgs
}
```

Create a dedicated Data Explorer harness whose ordinal columns are `PlantID`, `ProductID`, `CustomerID`, `Description`, and `Region`. Load metadata with Preview, then assert after each rapid character:

```powershell
$checkedBeforeTyping = @($outputList.CheckedItems | ForEach-Object Name) -join ','
foreach ($step in @(
    [pscustomobject]@{ Character=[char]'P'; Expected='PlantID'; Prefix='P' },
    [pscustomobject]@{ Character=[char]'R'; Expected='ProductID'; Prefix='PR' },
    [pscustomobject]@{ Character=[char]'O'; Expected='ProductID'; Prefix='PRO' },
    [pscustomobject]@{ Character=[char]'D'; Expected='ProductID'; Prefix='PROD' }
)) {
    $eventArgs = Invoke-TestOutputColumnKeyPress $navigationForm $step.Character
    Assert-Equal $true $eventArgs.Handled "Typing $($step.Character) suppresses native matching"
    Assert-Equal $step.Expected $outputList.SelectedItem.Name "Typing $($step.Prefix) selects its first prefix match"
    Assert-Equal $step.Prefix $navigationForm.Tag.DataExplorerOutputInteraction.Prefix "Typing stores prefix $($step.Prefix)"
    Assert-Equal $checkedBeforeTyping (@($outputList.CheckedItems | ForEach-Object Name) -join ',') `
        "Typing $($step.Prefix) preserves every output check"
}
```

Also assert:

- the stored timer interval is exactly `1000`;
- `C` after `PROD` retries as `C` and highlights `CustomerID` without changing checks;
- an unmatched `X` leaves the `CustomerID` highlight and checks unchanged while storing `X` until reset;
- invoking the timer's protected `OnTick` through `Invoke-TestProtectedControlEvent` clears only the prefix and stops it, leaving highlight/checks intact;
- a printable character after reset starts a new one-character prefix;
- space is handled and cannot toggle a check;
- selecting another table, metadata repopulation, Refresh, and confirmed Change Connection leave the prefix empty and timer stopped at their existing state boundaries.

- [ ] **Step 2: Run the focused UI test and confirm RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: exit `1` because the timer and custom prefix behavior do not exist. Confirm the failure belongs to the new navigation fixture, not Task 1.

- [ ] **Step 3: Implement reset and prefix-matching helpers**

Add these functions beside the Task 1 interaction helpers:

```powershell
function Reset-SqlUtilityDataExplorerOutputColumnNavigation {
    param([System.Windows.Forms.Form] $Form)

    $interaction = $Form.Tag.DataExplorerOutputInteraction
    $interaction.Prefix = ''
    if ($null -ne $interaction.ResetTimer) {
        $interaction.ResetTimer.Stop()
    }
}

function Find-SqlUtilityDataExplorerOutputColumnPrefix {
    param(
        [System.Windows.Forms.CheckedListBox] $OutputList,
        [string] $Prefix
    )

    for ($index = 0; $index -lt $OutputList.Items.Count; $index++) {
        $name = [string] $OutputList.Items[$index].Name
        if ($name.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $index
        }
    }
    return -1
}

function Invoke-SqlUtilityDataExplorerOutputColumnKeyPress {
    param(
        [System.Windows.Forms.Form] $Form,
        [System.Windows.Forms.KeyPressEventArgs] $EventArgs
    )

    if ([char]::IsControl($EventArgs.KeyChar)) { return }
    $EventArgs.Handled = $true

    $interaction = $Form.Tag.DataExplorerOutputInteraction
    $interaction.ResetTimer.Stop()
    $character = [string] $EventArgs.KeyChar
    $candidate = [string] $interaction.Prefix + $character
    $outputList = Get-SqlUtilityNamedControl $Form 'OutputColumnsList'
    $match = Find-SqlUtilityDataExplorerOutputColumnPrefix $outputList $candidate
    if ($match -lt 0) {
        $candidate = $character
        $match = Find-SqlUtilityDataExplorerOutputColumnPrefix $outputList $candidate
    }
    if ($match -ge 0) {
        $outputList.SelectedIndex = $match
    }
    $interaction.Prefix = $candidate
    $interaction.ResetTimer.Start()
}
```

Search only catalog item `.Name` values and keep first-item ordinal order. Do not alter `CheckedItems`, `DataExplorerBuilder.SelectedColumnNames`, or preview state.

- [ ] **Step 4: Construct, wire, reset, and dispose the timer**

After assigning `$form.Tag`, create one form-owned WinForms timer:

```powershell
$outputNavigationTimer = [System.Windows.Forms.Timer]::new()
$outputNavigationTimer.Interval = 1000
$state.DataExplorerOutputInteraction.ResetTimer = $outputNavigationTimer
$outputNavigationTimer.Add_Tick({
    Reset-SqlUtilityDataExplorerOutputColumnNavigation $form
}.GetNewClosure())
$form.Add_Disposed({
    $outputNavigationTimer.Stop()
    $outputNavigationTimer.Dispose()
}.GetNewClosure())
```

Wire `$outputs.Add_KeyPress(...)` to `Invoke-SqlUtilityDataExplorerOutputColumnKeyPress`.

Call `Reset-SqlUtilityDataExplorerOutputColumnNavigation` at the start of `Invoke-SqlUtilityLoadDataExplorerTables` after its busy guard, and immediately before every other application-owned output-list clear/repopulation path:

- catalog loading/Refresh in `Invoke-SqlUtilityLoadDataExplorerTables`, including a failed explicit refresh;
- first/reloaded metadata population in `Invoke-SqlUtilityDataExplorerPreview`;
- `Reset-SqlUtilityWorkspaceState`, which covers confirmed Change Connection;
- selected-table changes in the `PhysicalTablesList.SelectedIndexChanged` handler.

Do not reset on table-filter edits that retain the same selected object, ordinary column highlight/check changes, filter edits, Preview reruns with cached metadata, Send to Query, Export Preview, settings changes, or Query execution.

- [ ] **Step 5: Document the user-visible behavior**

In `README.md` under **Data Explorer**, add a concise paragraph stating:

- first Preview still checks all columns;
- only checkbox-glyph clicks or All/None change output checks;
- text/keyboard interaction only moves the highlight;
- rapid printable characters form a case-insensitive column-name prefix, reset after one second, with newest-character fallback when the full prefix has no match.

Do not describe the interval as a setting and do not add persistence/search-bar claims.

- [ ] **Step 6: Run GREEN verification for Task 2**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
git diff --check
```

Expected: UI suite exits `0`; the `PROD` progression, newest-character fallback, one-second reset, reset lifecycle, check invariants, and pre-existing Explorer workflows all pass.

- [ ] **Step 7: Review and commit Task 2**

Inspect the complete Task 2 diff, verify the timer is disposed and no state crossed into a module/config object, then commit:

```powershell
git add -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1 README.md
git diff --cached --check
git diff --cached --stat
git commit -m "feat: add buffered output column navigation"
```

---

### Task 3: Whole-change review and verification

**Files:**

- Review: `SqlUtility.ps1`
- Review: `tests/Test-SqlUtilityUi.ps1`
- Review: `README.md`
- Review: `docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-design.md`
- Review: `docs/superpowers/specs/2026-08-15-sql-utility-column-selection-navigation-design.md`

- [ ] **Step 1: Perform a spec-to-diff review**

Check every approved requirement against code and tests: default all checked; text/keyboard never check; glyph click once; All/None; `PROD`; one-second reset; newest-character fallback; unmatched highlight preservation; lifecycle resets; timer disposal; snapshot independence; no search bar/persistence/module change.

Search for accidental alternate check paths:

```powershell
rg -n "CheckOnClick|SetItemChecked|Items\.Add\(.+,\s*\$true|Add_ItemCheck|Add_MouseDown|Add_KeyPress|DataExplorerOutputInteraction" SqlUtility.ps1 tests\Test-SqlUtilityUi.ps1
```

Every production check-changing call must be inside the application-owned guard.

- [ ] **Step 2: Run final verification from the feature worktree root**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

Expected: both suites exit `0`; diff check has no errors; status contains only intentional report artifacts, if any. Preserve the exact output for the handoff.

- [ ] **Step 3: Inspect commit scope and history**

Run:

```powershell
git log --oneline --decorate -5
git diff main...HEAD --stat
git diff main...HEAD -- SqlUtility.ps1 tests\Test-SqlUtilityUi.ps1 README.md docs\superpowers\specs
```

Confirm the branch contains the approved design/plan plus the two scoped implementation commits and no unrelated changes.

- [ ] **Step 4: Report completion and external acceptance boundary**

Report the behavior outcome, exact test commands/results, commit IDs, and changed files. State that Citrix DPI checkbox-hit feel and the one-second typing feel remain pending manual acceptance. Do not push, merge, or remove the worktree until the user chooses an integration action.
