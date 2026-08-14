# SQL Utility Data Explorer Layout Correction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the clipped flat Data Explorer layout with responsive table, builder, filter, and preview panes whose scrollbars remain accessible throughout resizing.

**Architecture:** Keep every existing named workflow control and event handler in `SqlUtility.ps1`, but reparent them into nested dock-filled `SplitContainer` and `TableLayoutPanel` controls. Make dynamic filter rows width-responsive so the filter pane owns vertical scrolling without horizontal overflow. Protect the contract with real WinForms geometry, scroll-position, splitter-limit, and state-preservation tests.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework WinForms, dependency-free PowerShell regression scripts.

## Global Constraints

- Preserve `StartSqlUtility.cmd` and `powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass`.
- Preserve Windows PowerShell 5.1, .NET Framework, Windows integrated authentication, read-only query boundaries, and the exact seven-file runtime distribution.
- Modify only `SqlUtility.ps1` and `tests/Test-SqlUtilityUi.ps1` for production/test behavior; the approved design documents are already committed.
- Do not change table discovery, Preview, output checks, filter semantics, Send to Query, Export Preview, configuration, QueryPolicy, Database, or Excel behavior.
- Keep form default size `960 x 680`, minimum size `760 x 520`, and `AutoScaleMode = Dpi`.
- Keep standard `CheckedListBox` matching navigation unchanged; issue 4 remains a separate design discussion.
- Use regression-first TDD: observe the focused UI suite fail for the intended geometry/scroll reason before changing production code.
- Treat Citrix DPI/scaling and live database behavior as pending external acceptance.

---

### Task 1: Responsive Data Explorer Pane Hierarchy

**Files:**
- Modify: `tests/Test-SqlUtilityUi.ps1:1-330,1379-1521`
- Modify: `SqlUtility.ps1:1249-1264,1332-1354`

**Interfaces:**
- Produces named containers discoverable through existing recursive `Get-SqlUtilityNamedControl`/`Get-TestControl` helpers:
  - `DataExplorerMainSplit`
  - `DataExplorerTableLayout`
  - `DataExplorerRightSplit`
  - `DataExplorerBuilderLayout`
  - `DataExplorerBuilderSplit`
  - `DataExplorerColumnsLayout`
  - `DataExplorerFiltersLayout`
  - `DataExplorerActionLayout`
  - `DataExplorerPreviewLayout`
- Preserves every existing named control and handler.
- Uses `SplitContainer.FixedPanel = None`; all three splitters remain user-adjustable.

- [ ] **Step 1: Add a containment assertion helper and failing default-size regression**

Add this test-only helper near `Get-TestControl`:

```powershell
function Assert-TestControlContained($Control, [string] $Message) {
    $parent = $Control.Parent
    $contained = $null -ne $parent -and
        $Control.Left -ge 0 -and $Control.Top -ge 0 -and
        $Control.Right -le $parent.ClientSize.Width -and
        $Control.Bottom -le $parent.ClientSize.Height
    Assert-True $contained $Message
}
```

In the Data Explorer UI block, after selecting the tab and processing events, assert containment using hand-derived expectations:

```powershell
foreach ($name in @(
    'PhysicalTablesList',
    'OutputColumnsList',
    'DataExplorerFiltersPanel',
    'PreviewGrid'
)) {
    Assert-TestControlContained (Get-TestControl $explorerForm $name) `
        "$name remains inside its owning pane at default size"
}
```

The production mutation caught by this test is reintroducing direct bottom/right anchors or any layout that places a scrollbar edge outside its immediate parent.

- [ ] **Step 2: Run the focused suite and verify RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: exit `1`; the current `PhysicalTablesList` containment assertion fails because its bottom exceeds its parent client height. Record the actual failure before production edits.

- [ ] **Step 3: Add failing hierarchy and scroll-configuration assertions**

Extend `$requiredControlNames` with the nine container names above. Assert the approved structure and orientations:

```powershell
$mainSplit = Get-TestControl $explorerForm 'DataExplorerMainSplit'
$rightSplit = Get-TestControl $explorerForm 'DataExplorerRightSplit'
$builderSplit = Get-TestControl $explorerForm 'DataExplorerBuilderSplit'

Assert-Equal ([System.Windows.Forms.Orientation]::Vertical) $mainSplit.Orientation `
    'Table pane is left of the main Data Explorer area'
Assert-Equal ([System.Windows.Forms.Orientation]::Horizontal) $rightSplit.Orientation `
    'Builder is above Preview'
Assert-Equal ([System.Windows.Forms.Orientation]::Vertical) $builderSplit.Orientation `
    'Columns are left of filters'
Assert-Equal ([System.Windows.Forms.FixedPanel]::None) $mainSplit.FixedPanel `
    'Main panes resize without a fixed panel'
Assert-True ([object]::ReferenceEquals(
    (Get-TestControl $explorerForm 'PhysicalTablesList').Parent,
    (Get-TestControl $explorerForm 'DataExplorerTableLayout')
)) 'Table list is owned by the table layout'
Assert-True ([object]::ReferenceEquals(
    (Get-TestControl $explorerForm 'PreviewGrid').Parent,
    (Get-TestControl $explorerForm 'DataExplorerPreviewLayout')
)) 'Preview grid is owned by the lower preview layout'

Assert-Equal $false (Get-TestControl $explorerForm 'PhysicalTablesList').HorizontalScrollbar `
    'Table list is vertical-scroll only'
Assert-Equal $false (Get-TestControl $explorerForm 'OutputColumnsList').HorizontalScrollbar `
    'Output list is vertical-scroll only'
Assert-Equal $true (Get-TestControl $explorerForm 'DataExplorerFiltersPanel').AutoScroll `
    'Filter pane owns scrolling'
Assert-Equal ([System.Windows.Forms.ScrollBars]::Both) `
    (Get-TestControl $explorerForm 'PreviewGrid').ScrollBars `
    'Preview grid supports horizontal and vertical scrolling'
```

Expected before implementation: missing named-container failures.

- [ ] **Step 4: Replace direct TabPage placement with nested dock-filled containers**

Replace the flat Data Explorer construction block. Create the existing action/content controls without `Location`, `Size`, or direct `TabPage.Controls.Add` calls. Construct this hierarchy:

```powershell
$dataExplorerMainSplit = [System.Windows.Forms.SplitContainer]::new()
$dataExplorerMainSplit.Name = 'DataExplorerMainSplit'
$dataExplorerMainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataExplorerMainSplit.Orientation = [System.Windows.Forms.Orientation]::Vertical
$dataExplorerMainSplit.FixedPanel = [System.Windows.Forms.FixedPanel]::None
$dataExplorerMainSplit.Panel1MinSize = 180
$dataExplorerMainSplit.Panel2MinSize = 420
$dataExplorerTab.Controls.Add($dataExplorerMainSplit)

$dataExplorerRightSplit = [System.Windows.Forms.SplitContainer]::new()
$dataExplorerRightSplit.Name = 'DataExplorerRightSplit'
$dataExplorerRightSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataExplorerRightSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$dataExplorerRightSplit.FixedPanel = [System.Windows.Forms.FixedPanel]::None
$dataExplorerRightSplit.Panel1MinSize = 150
$dataExplorerRightSplit.Panel2MinSize = 120
$dataExplorerMainSplit.Panel2.Controls.Add($dataExplorerRightSplit)

$dataExplorerBuilderSplit = [System.Windows.Forms.SplitContainer]::new()
$dataExplorerBuilderSplit.Name = 'DataExplorerBuilderSplit'
$dataExplorerBuilderSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataExplorerBuilderSplit.Orientation = [System.Windows.Forms.Orientation]::Vertical
$dataExplorerBuilderSplit.FixedPanel = [System.Windows.Forms.FixedPanel]::None
$dataExplorerBuilderSplit.Panel1MinSize = 150
$dataExplorerBuilderSplit.Panel2MinSize = 260
```

Use named `TableLayoutPanel` controls for the table pane, builder rows, columns pane, filters pane, action toolbar, and preview pane. Every root/content control uses `Dock = Fill`; action rows use `AutoSize`. Put:

- table filter and Refresh in the table layout's first row and `PhysicalTablesList` in its percentage-height second row;
- `OutputColumnsList` above All/None in the builder split's left panel;
- Add Filter/Clear above `DataExplorerFiltersPanel` in the builder split's right panel;
- the builder split above Preview/Export Preview/Send to Query in the upper-right builder layout;
- source/status labels above `PreviewGrid` in the lower-right preview layout.

Set `PhysicalTablesList.HorizontalScrollbar` and `OutputColumnsList.HorizontalScrollbar` explicitly to `$false`; set `PreviewGrid.ScrollBars` explicitly to `Both`.

Initialize splitter distances only after the controls have nonzero final client sizes. Use one shared closure attached to `DataExplorerTab.Layout`, guarded by a Boolean flag, and clamp each distance between its `Panel1MinSize` and available size minus `Panel2MinSize` and `SplitterWidth`. Initial targets are `230` pixels for the table pane, `200` pixels for columns, and half the right-pane height for the builder.

- [ ] **Step 5: Run the focused suite and verify GREEN for hierarchy/containment**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: exit `0`, including all pre-existing Data Explorer workflow tests.

- [ ] **Step 6: Commit Task 1**

Run:

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git diff --cached --check
git diff --cached --stat
git commit -m "fix: make data explorer panes responsive"
```

---

### Task 2: Filter-Row Scrolling and Resize Stability

**Files:**
- Modify: `tests/Test-SqlUtilityUi.ps1` Data Explorer geometry block
- Modify: `SqlUtility.ps1:559-581` dynamic filter-row construction and Data Explorer panel resize handler

**Interfaces:**
- Produces: `Resize-SqlUtilityDataExplorerFilterRows -Form`, which sizes existing dynamic filter rows to the visible filter-panel client width without changing their state or order.
- Preserves `Add-SqlUtilityDataExplorerFilterRow` return value and row `Tag` members: `RowNumber`, `ColumnCombo`, `OperatorCombo`, `ValueText`, `ValueBitCombo`, `RemoveButton`.
- Keeps text and bit value controls mutually exclusive in one value cell.

- [ ] **Step 1: Add failing dynamic-row and resize-cycle regressions**

Add a separate geometry-only `$layoutHarness` and `$layoutForm` so the existing workflow fixture remains unchanged. Before constructing the form:

- Set `TablesResult` to 100 physical-table objects named `Table001` through `Table100`.
- Set `ColumnsResult` to 80 nullable `nvarchar(100)` column objects named `Column001` through `Column080`.
- Set `PreviewResult` to a 100-row `DataTable` with those same 80 string columns; give every cell a 40-character value so both grid axes overflow at every tested size.

Show the real form, enter the workspace, select Data Explorer, select `Table001`, and click Preview. Add exactly eight filter rows, select `Column001` through `Column008`, keep `Contains`, and set their values to `value-01` through `value-08`. Exercise the real controls, not a mock layout. Set nonzero scroll positions:

```powershell
$tableList = Get-TestControl $explorerForm 'PhysicalTablesList'
$outputList = Get-TestControl $explorerForm 'OutputColumnsList'
$filtersPanel = Get-TestControl $explorerForm 'DataExplorerFiltersPanel'
$previewGrid = Get-TestControl $explorerForm 'PreviewGrid'

$tableList.TopIndex = 10
$outputList.TopIndex = 10
$filtersPanel.AutoScrollPosition = [System.Drawing.Point]::new(0, 40)
```

Capture the builder object, preview object, checked names, filter values, list top indexes, and panel scroll position. Resize through this literal sequence, processing WinForms events after every size:

```powershell
foreach ($size in @(
    [System.Drawing.Size]::new(760, 520),
    [System.Drawing.Size]::new(860, 600),
    [System.Drawing.Size]::new(960, 680),
    [System.Drawing.Size]::new(1280, 800),
    [System.Drawing.Size]::new(760, 520)
)) {
    $explorerForm.Size = $size
    [System.Windows.Forms.Application]::DoEvents()
    foreach ($name in @(
        'PhysicalTablesList',
        'OutputColumnsList',
        'DataExplorerFiltersPanel',
        'PreviewGrid'
    )) {
        Assert-TestControlContained (Get-TestControl $explorerForm $name) `
            "$name remains contained at $($size.Width)x$($size.Height)"
    }
}
```

Then, at `960 x 680`, set each split container to its literal allowed extremes (`Panel1MinSize`, then `Width/Height - Panel2MinSize - SplitterWidth`) and re-run containment assertions after each change. Restore the three original splitter distances before state comparisons.

Assert:

- Builder and preview object references are unchanged.
- Checked column names and all eight filter values are unchanged and in order.
- List top indexes and filter vertical scroll remain nonzero; the fixed 100-table, 80-column, and eight-filter fixture guarantees overflow throughout the size cycle.
- `filtersPanel.HorizontalScroll.Visible` is false and `VerticalScroll.Visible` is true.
- `previewGrid.ScrollBars` remains `Both`; set `FirstDisplayedScrollingRowIndex` and `HorizontalScrollingOffset` to nonzero values where the bound test data exceeds the viewport, resize, and assert they remain nonzero.

The production mutations caught are recreating/rebinding controls during resize, returning to fixed-width filter rows, or allowing a scrollbar edge outside its pane.

- [ ] **Step 2: Run the focused suite and verify RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
```

Expected: exit `1` because current filter rows retain fixed width `565`, use five independent columns, and can introduce horizontal overflow in the side-by-side filter pane.

- [ ] **Step 3: Implement responsive four-cell filter rows**

Add:

```powershell
function Resize-SqlUtilityDataExplorerFilterRows {
    param([System.Windows.Forms.Form] $Form)
    $panel = Get-SqlUtilityNamedControl $Form 'DataExplorerFiltersPanel'
    if ($null -eq $panel -or $panel.ClientSize.Width -le 0) { return }
    $rowWidth = [Math]::Max(1,
        $panel.ClientSize.Width - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth - 4)
    foreach ($row in @($panel.Controls)) {
        $row.Width = $rowWidth
    }
}
```

Change each dynamic row to four cells: column, operator, value host, and Remove. Add the text and bit controls to one dock-filled value-host `Panel`; only the type-appropriate control remains visible. Define the row's first three `ColumnStyle` entries as percentages and the Remove column as `AutoSize`. Dock the combo/value controls to `Fill`, preserve their existing names and `Tag` references, and keep the row height `31`.

Call the resize helper after adding a row and from `DataExplorerFiltersPanel.ClientSizeChanged`. Do not rebuild controls during resize. Preserve current visual order and `Get-SqlUtilityDataExplorerFilters` iteration.

- [ ] **Step 4: Run focused and aggregate verification**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-DataExplorer.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

Expected: every command exits `0`; aggregate reports `All SQL Utility tests passed.`; only Task 2 files are modified before staging.

- [ ] **Step 5: Run a local geometry diagnostic**

Instantiate the real form under Windows PowerShell 5.1 STA with `-NoGui`, enter the workspace, select Data Explorer with catalog loading suppressed, and print the tab/client bounds at default, minimum, and enlarged sizes. Confirm every content control's right/bottom is within its immediate parent and the three splitters remain draggable.

This local check does not replace Citrix DPI acceptance.

- [ ] **Step 6: Commit Task 2**

Run:

```powershell
git add SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1
git diff --cached --check
git diff --cached --stat
git commit -m "fix: keep explorer scrolling stable on resize"
```

---

### Task 3: Final Review and Draft-PR Update

**Files:**
- Verify: `SqlUtility.ps1`, `tests/Test-SqlUtilityUi.ps1`, all test suites
- Verify: `docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-layout-correction-design.md`

**Interfaces:**
- No new runtime interface.
- Draft PR remains `main <- feature/data-explorer`; the feature worktree remains preserved.

- [ ] **Step 1: Inspect the complete correction diff**

Run:

```powershell
git diff 1b2b1ca..HEAD -- SqlUtility.ps1 tests/Test-SqlUtilityUi.ps1 docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-design.md docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-layout-correction-design.md
git diff --check
git status --short
```

Confirm no matching-navigation handler, database behavior, query behavior, configuration, or packaging change entered the diff.

- [ ] **Step 2: Run fresh final verification**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
git diff --check
git status --short
```

Expected: all tests pass, diff check is silent, and worktree is clean.

- [ ] **Step 3: Request focused code review**

Review the correction range from `1b2b1ca` to `HEAD` against the layout-correction design. Resolve every Critical/Important finding and re-run the focused and aggregate suites after any correction.

- [ ] **Step 4: Push the verified branch**

After clean review and verification, push `feature/data-explorer` to its existing upstream outside the sandbox/keyring boundary. Read back draft PR #3 and confirm its head SHA matches local `HEAD`. Preserve the worktree for further visual/Citrix feedback.

Report Citrix DPI/scaling as pending external acceptance.
