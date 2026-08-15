# SQL Utility Column Selection Navigation Design

## Purpose

Correct the Data Explorer output-column list so navigation never changes output checks, while preserving fast keyboard access for tables with many columns. A user can type a column-name prefix such as `PROD` to reach `ProductID`, and only a checkbox-glyph click or the existing All/None actions can change checked output columns.

This specification extends the approved [SQL Utility Data Explorer Design](2026-08-14-sql-utility-data-explorer-design.md). It replaces only that design's standard `CheckedListBox` matching-navigation contract. All table discovery, metadata, builder, preview, filter, export, query-handoff, layout, persistence, and module boundaries remain unchanged.

## Confirmed Native Behavior

The current output list is a standard .NET Framework `CheckedListBox` with `CheckOnClick = true` and no custom keyboard or mouse handlers. Local Windows PowerShell 5.1 WinForms probes confirmed:

- Clicking an item's text toggles its checkbox.
- Typing `PROD` is handled as four independent characters, navigating through matches for `P`, `R`, `O`, and `D` instead of accumulating `PROD`.
- With `CheckOnClick = true`, those typed navigation changes can also alter checks.
- Setting `CheckOnClick = false` prevents typed navigation from altering checks, but it does not provide prefix buffering and repeated text clicks can still toggle a check.

The exact approved interaction therefore cannot be obtained from a `CheckedListBox` property alone.

## Approved Interaction

- Output columns still default to all checked after the first Preview.
- Clicking directly on a checkbox glyph toggles that column once.
- Clicking column text may move the list highlight, but it never changes any checkbox.
- Keyboard navigation never changes any checkbox.
- Printable typed characters accumulate into a case-insensitive column-name prefix while the output list has focus.
- The prefix buffer resets after `1000` milliseconds of inactivity.
- Matching uses the catalog column `Name`, not its longer display text or SQL type annotation.
- Matching always selects the first column in ordinal display order whose name starts with the complete prefix. For example, `P` selects `PlantID`, then `PR`, `PRO`, and `PROD` select `ProductID`.
- If the complete prefix has no match, navigation retries with only the newest character. If that character also has no match, the current highlight and every checkbox remain unchanged.
- Resetting the buffer does not clear or move the current highlight.
- All and None remain explicit bulk check actions and do not execute Preview or modify the displayed preview snapshot.

The one-second interval is intentionally easy to adjust after hands-on use, but no user-facing setting or persistence is added.

## Chosen Design

### Check-state gate

Set `OutputColumnsList.CheckOnClick` to false. Keep one transient, mutable UI state object that records whether an application-owned check change is currently permitted.

The list's `ItemCheck` handler cancels any state change unless that guard is active. Checkbox-glyph clicks and the existing All/None handlers activate the guard only around their synchronous `SetItemChecked` calls and restore it in `finally`. Initial metadata population uses the same guarded path when adding or checking items.

This prevents keyboard input, column-text clicks, repeated clicks on a highlighted row, and native space-bar behavior from changing checks. Programmatic changes remain explicit and auditable.

### Checkbox hit testing

The list's mouse handler identifies the item under the pointer with `IndexFromPoint` and its item rectangle. A click counts as a checkbox click only when its X coordinate is at least the item rectangle's left edge and less than that edge plus `SystemInformation.MenuCheckSize.Width`; the application remains left-to-right. A qualifying click selects/highlights the row and toggles it once through the guarded check path. Other clicks retain normal row highlighting but do not change checks.

No owner-drawn list, compiled helper, custom control assembly, or hard-coded unscaled checkbox width is introduced.

### Incremental prefix navigation

The list's `KeyPress` handler consumes printable characters before native single-character matching can run. It restarts a WinForms `Timer` with `Interval = 1000` after every accepted character, performs an ordinal-ignore-case prefix search over item `Name` values, and updates only `SelectedIndex` for the first match.

The timer tick clears only the prefix buffer. Table/builder reset, metadata repopulation, Refresh, and Change Connection also clear the buffer and stop the timer. The timer is stopped and disposed with the form so no event source outlives its controls.

The navigation state is process-only. It is not part of builder state, preview state, configuration, or any module interface.

## Alternatives Considered

### Chosen: custom behavior on the existing checked list

This satisfies both checkbox-only changes and multi-character navigation without adding visible UI or changing selected-column state ownership. The change remains within `SqlUtility.ps1` and its UI tests.

### Rejected: native `CheckedListBox` behavior

Neither `CheckOnClick` setting provides both required click semantics and prefix matching. Native character navigation treats `PROD` as independent searches in the target runtime.

### Rejected: column filter/search box

Filtering would require preserving checks for hidden columns independently from `CheckedItems`, updating query-generation selection collection, and defining how All/None apply to hidden results. That is a larger state-model and UI change than the approved navigation requirement.

## Scope and Architecture

Production behavior changes only in `SqlUtility.ps1`, which already owns WinForms construction, Data Explorer UI state, and event orchestration. No WinForms dependency enters `SqlUtility.DataExplorer.ps1` or any other module.

There is no database call, query generation, Preview execution, export, configuration, persistence, runtime-file, authentication, or QueryPolicy change. Standard table-list substring filtering remains unchanged.

README and the owning Data Explorer design must describe the new checkbox and prefix-navigation behavior. No instructional label or search box is added to the application.

## Testing

Extend `tests/Test-SqlUtilityUi.ps1` with regression-first real WinForms behavior checks:

- First Preview still checks every output column.
- Rapid `P`, `R`, `O`, `D` input accumulates and finishes on `ProductID`, while the exact checked-column set remains unchanged after every character.
- The timer interval is exactly `1000` milliseconds, and input after reset starts a new prefix.
- A no-match prefix retries from the newest character; a completely unmatched character preserves the highlight and checks.
- Repeated column-text clicks may highlight but never toggle the item.
- Each checkbox-glyph click toggles exactly once.
- Native keyboard navigation and space do not change checks.
- All and None still update all columns, including after keyboard navigation, without changing the displayed preview snapshot.
- Table change, metadata repopulation, Refresh, and Change Connection reset transient navigation state without affecting their existing builder/snapshot contracts.

Run the focused UI suite, aggregate `tests/Test-All.ps1`, `git diff --check`, and worktree/staged status inspection. Local WinForms tests cover the behavior contract; checkbox hit-area feel and scaling remain manual Citrix/DPI acceptance.

## External Acceptance

In the target Citrix environment, verify checkbox-only toggling and text-click non-toggling at the deployed DPI/scaling, then try the one-second prefix reset with representative wide schemas. The reset interval may be adjusted in a separately reviewed follow-up based on that hands-on result.
