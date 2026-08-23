# SQL Utility Data Explorer Layout Correction Design

## Purpose

Correct the Data Explorer layout and scrolling defects discovered during default-size GUI acceptance without changing its query, state, persistence, export, or matching-navigation behavior.

This specification extends the approved [SQL Utility Data Explorer Design](2026-08-14-sql-utility-data-explorer-design.md). When the two documents differ on control geometry or pane hierarchy, this correction governs. All other Data Explorer contracts remain owned by the original design.

## Confirmed Defects

The current tab constructs every control directly on an initially unlaid-out `TabPage` with fixed pixel locations and `Top,Bottom,Left,Right` anchors. At 96 DPI and the default `960 x 680` window, the visible Data Explorer client area is `912 x 523`, but anchored controls expand beyond it during first layout:

- `PhysicalTablesList` ends at Y `961`.
- `OutputColumnsList` ends at X `1307`.
- `DataExplorerFiltersPanel` ends at X `1617`.
- `PreviewGrid` ends at X `1617` and Y `965`.

The controls technically provide scrolling, but their scrollbar edges are outside the visible tab. The current flat vertical stack also places output columns above filters rather than side by side in the upper builder area.

The existing UI tests assert control presence and workflow behavior but do not assert the Data Explorer parent hierarchy, default/minimum-size containment, or visible scrollbar ownership.

## Scope

This correction includes only:

- A responsive left table pane.
- An upper-right builder split into output columns on the left and filters on the right.
- A single builder toolbar with compact column-selection actions on the left and preview actions on the right.
- A lower-right preview pane.
- A 10-pixel builder/preview splitter that is easier to drag over a remote connection.
- Visible vertical scrolling for tables, output columns, and filters.
- Visible horizontal and vertical scrolling for the preview grid.
- Default-size and minimum-size layout regression coverage.

The following remain unchanged:

- Physical-table loading, substring table filtering, Refresh, metadata retrieval, Preview, filter semantics, output checks, Send to Query, Export Preview, settings, busy behavior, reset behavior, and snapshot independence.
- Every service interface and module boundary.
- Window default size `960 x 680`, minimum size `760 x 520`, and DPI autoscaling.
- Standard `CheckedListBox` matching navigation. Its selection/check behavior and multi-character matching are deferred to a separate design discussion.
- QueryPolicy, Database, Config, and Excel production modules.

## Chosen Layout

The Data Explorer tab uses nested WinForms layout containers rather than absolute coordinates:

```text
DataExplorerTab
└─ DataExplorerMainSplit (vertical)
   ├─ Panel1: table pane
   │  └─ TableLayoutPanel
   │     ├─ table filter + Refresh
   │     └─ PhysicalTablesList (fill)
   └─ Panel2: DataExplorerRightSplit (horizontal)
      ├─ Panel1: upper builder
      │  └─ TableLayoutPanel
      │     ├─ DataExplorerBuilderSplit (vertical)
      │     │  ├─ output-column pane
      │     │  │  └─ OutputColumnsList (fill)
      │     │  └─ filter pane
      │     │     ├─ Add Filter + Clear
      │     │     └─ DataExplorerFiltersPanel (fill)
      │     └─ All + None | flexible space | Preview + Export Preview + Send to Query
      └─ Panel2: lower preview
         └─ TableLayoutPanel
            ├─ PreviewSourceLabel + PreviewStatusLabel
            └─ PreviewGrid (fill)
```

All root containers and content controls use `Dock = Fill`. Split containers remain user-adjustable, use `FixedPanel = None`, and define minimum panel sizes that keep their contained actions usable. The horizontal builder/preview splitter is 10 pixels wide. Initial splitter distances favor a compact table pane, a larger filter pane than output pane, and approximately balanced builder/preview heights at the default window size.

All and None use their compact native widths at the left of the shared toolbar. A percentage-width spacer absorbs the remaining room before Preview, Export Preview, and Send to Query, keeping those actions aligned to the right. The filter editor remains entirely above the toolbar; adding rows cannot displace or overlap its actions.

No control is anchored directly to an unlaid-out `TabPage`.

## Scrolling Contract

- `PhysicalTablesList` uses its native vertical list scrollbar and no horizontal scrollbar.
- `OutputColumnsList` uses its native vertical checked-list scrollbar and no horizontal scrollbar.
- `DataExplorerFiltersPanel` scrolls vertically. Dynamic filter rows size to the panel client width and do not require horizontal scrolling.
- Filter scrolling activates automatically when added rows exceed the editor viewport, while the shared toolbar remains visible below it.
- Each filter row remains a `TableLayoutPanel`, but column, operator, value, and Remove controls share the available row width. Text and bit value controls occupy the same value cell because only one is visible at a time.
- `PreviewGrid.ScrollBars` is explicitly `Both`. Its existing 300-pixel column-width cap remains unchanged so wide schemas scroll horizontally rather than forcing the grid outside its pane.

Scrollbars must remain inside their owning pane throughout resizing at and above the minimum window size. They may disappear when all content fits and reappear when content exceeds the viewport, but they must never move outside the visible pane. Resizing or moving a splitter must not recreate or rebind controls, reset tables, checked columns, filters, or preview state, or unnecessarily reset the current scroll position. DPI-scaled Citrix rendering remains an external acceptance check, but the container hierarchy must scale without fixed bottom/right offsets.

## Alternatives Considered

### Chosen: nested `SplitContainer` and `TableLayoutPanel` controls

This matches the approved visual hierarchy, gives the user adjustable pane proportions, and makes containment a property of docking rather than manual arithmetic.

### Rejected: repair the current fixed coordinates and anchors

This would be the smallest textual diff but would retain the construction-time anchor defect and remain fragile under DPI scaling, minimum-size resizing, and future control additions.

### Rejected: one flat `TableLayoutPanel` for the whole tab

This would contain controls correctly but would not give independent, intuitive resizing between the table list, builder, and preview grid. It also makes the upper side-by-side builder and lower preview relationship less explicit.

## State and Event Preservation

Existing named controls and event handlers remain authoritative. Reparenting controls must not replace their names or change how helper functions locate them recursively. Dynamic filter rows continue to emit the same neutral `ColumnName`, `Operator`, and `ValueText` objects in visual order.

Reparenting and resizing must not trigger database work, change checked output columns, remove filter rows, clear the preview snapshot, or alter button eligibility. Table selection remains the only builder reset trigger, subject to the existing same-object preservation rule.

## Testing

Extend `tests/Test-SqlUtilityUi.ps1` with regression-first tests that fail against the flat anchored layout and assert real control behavior:

- Named split/layout containers exist with the approved orientation and parent hierarchy.
- The builder/preview splitter is 10 pixels wide, and all five toolbar actions remain on one row with compact All/None actions on the left and preview actions on the right.
- At minimum `760 x 520`, default `960 x 680`, intermediate, and enlarged window sizes, the table list, output list, filter panel, and preview grid bounds are fully contained by their immediate parent client rectangles.
- Repeated shrink, expand, and shrink cycles preserve containment, visible scrollbar ownership, builder/preview state, and the current scroll position when the content still requires scrolling.
- Moving each splitter to its allowed minimum and maximum distances preserves pane containment and does not reset or rebind Data Explorer controls.
- After enough table/column/filter/preview content is added, table, output, and filter scrolling remains vertical-only and visible within its pane; the preview grid is configured for both directions and remains contained.
- Adding filter rows until the editor overflows enables its vertical scroller without moving or covering the shared toolbar.
- Filter rows resize with the filter panel, retain visual order and values, and do not introduce a horizontal scrollbar.
- Existing first/later Preview, filter, Send, snapshot, Export Preview, refresh, and connection-reset tests remain unchanged and pass.

Run the focused UI suite first, followed by `tests/Test-All.ps1`, `git diff --check`, and worktree status inspection. A manual default/minimum-size local check is required. Citrix DPI/scaling and live database acceptance remain pending externally.

## Documentation

The owning Data Explorer design links to this correction. README behavior does not need expansion because query and user workflow contracts do not change; its testing and external-acceptance statements remain current.
