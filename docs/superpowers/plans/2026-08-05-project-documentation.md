# SQL Utility Project Documentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an accurate repository-root `README.md` for users and maintainers plus an AI-agent-focused repository-root `AGENTS.md` that preserves SQL Utility's approved architecture and engineering constraints.

**Architecture:** Use a layered documentation model. `README.md` is the explanatory source for current application behavior, operation, architecture, security, testing, and limitations; `AGENTS.md` is a concise binding instruction layer that links to the README and approved design documents rather than duplicating their detail.

**Tech Stack:** GitHub-flavored Markdown, Mermaid, Windows PowerShell 5.1 validation commands, existing dependency-free PowerShell test suite.

## Global Constraints

- Describe the final code at the current feature-branch head, including release-review fixes; do not copy superseded code snippets from the original plan.
- Preserve the portable `StartSqlUtility.cmd` to Windows PowerShell 5.1 architecture and the six-file runtime distribution.
- Document only Windows and .NET Framework components normally available in the target Citrix environment; do not introduce dependencies or code changes.
- State that authentication is Windows integrated only and that credentials are never requested or persisted.
- State that the only automatic persistent file is app-local `SqlUtility.config.json`; Excel output goes only to the user-selected location.
- Use exact values: unordered-row limit `100` through `2000`, default `1000`; Query/Export timeout `5` through `3600` seconds, default `120`; connection timeout `10` seconds; display page size `500`.
- Distinguish local automated verification from pending external Citrix, live SQL Server, cloud-drive, and desktop Excel acceptance.
- Keep `README.md` explanatory and `AGENTS.md` prescriptive; link instead of duplicating long explanations.
- Do not modify production scripts, modules, tests, configuration schema, query grammar, packaging, or runtime behavior.
- Use relative Markdown links for repository documents.

---

### Task 1: Create the Project README

**Files:**

- Create: `README.md`
- Reference: `StartSqlUtility.cmd`
- Reference: `SqlUtility.ps1`
- Reference: `modules/SqlUtility.Config.ps1`
- Reference: `modules/SqlUtility.QueryPolicy.ps1`
- Reference: `modules/SqlUtility.Database.ps1`
- Reference: `modules/SqlUtility.Excel.ps1`
- Reference: `tests/Test-All.ps1`
- Reference: `docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md`
- Reference: `docs/superpowers/plans/2026-08-02-sql-utility-v1.md`
- Reference: `docs/superpowers/specs/2026-08-05-project-documentation-design.md`

**Interfaces:**

- Produces: the primary project introduction, operation guide, final design overview, implementation approach, verification guide, and limitation reference.
- Consumed by: internal Citrix users, human maintainers, and repository agents directed there by `AGENTS.md`.

- [ ] **Step 1: Verify the README is currently absent**

Run from the repository root:

```powershell
if (Test-Path -LiteralPath .\README.md) {
    throw 'README.md already exists; inspect it before following this create-only task.'
}
Write-Host 'RED: README.md is absent.'
```

Expected: output `RED: README.md is absent.`.

- [ ] **Step 2: Create the README with the approved structure**

Create `README.md` with these exact top-level sections in this order:

```markdown
# SQL Utility

## Overview
## Version 1 Features
## Requirements and Portability
## Portable Distribution
## Running the Application
## User Workflow
## Architecture and Implementation
## Configuration and Persistence
## Query Policy
## Query Execution and Paging
## Excel Export
## Security and Data Boundaries
## Error Handling and Runtime Model
## Testing
## Citrix Acceptance Checklist
## Version 1 Limitations
## Future Extensions
## Design References
```

Populate the sections as follows:

- **Overview:** Identify the app as a portable internal Windows GUI for Windows-authenticated SQL Server access, restricted read-only querying, paging, and `.xlsx` export. State that v1 is intentionally small and synchronous.
- **Version 1 Features:** Cover connection-first startup, saved server/database pairs, Query and Settings tabs, single-table read-only `SELECT`, two paging models, complete-result export, and global settings.
- **Requirements and Portability:** List Windows PowerShell 5.1 and built-in .NET Framework components: WinForms, Drawing, Data.SqlClient, IO.Compression, Xml, and built-in JSON commands. State no installer, executable compilation, administrator rights, third-party module, NuGet, Office installation, `sqlcmd`, Python, registry write, or environment mutation.
- **Portable Distribution:** Show this exact tree and state that all six files must remain together:

  ```text
  StartSqlUtility.cmd
  SqlUtility.ps1
  modules/
    SqlUtility.Config.ps1
    SqlUtility.QueryPolicy.ps1
    SqlUtility.Database.ps1
    SqlUtility.Excel.ps1
  ```

- **Running the Application:** Instruct users to extract/copy the folder, ensure it is writable if persistence is desired, and launch `StartSqlUtility.cmd`. Explain `-NoLogo -NoProfile -STA -ExecutionPolicy Bypass` as a process-only invocation and warn not to run from inside a ZIP.
- **User Workflow:** Explain blank initial Server/Database fields, Test Connection, Connect, saved-pair selection/deletion, query execution, paging, export, settings, and confirmed Change Connection state loss.
- **Architecture and Implementation:** Include a Mermaid flowchart with nodes for `.cmd launcher`, `SqlUtility.ps1`, the four modules, SQL Server, `SqlUtility.config.json`, and `.xlsx destination`. Show that UI calls Config/Policy/Database/Excel, Database talks to SQL Server, Config owns JSON, and Excel owns workbook generation. Follow it with a responsibility table for all five PowerShell files. Emphasize neutral `DataTable`/metadata/schema/row boundaries and deterministic SQL resource disposal.
- **Configuration and Persistence:** Show schema version 1 JSON with default settings and a sample saved pair. Document exact ranges/defaults, case-insensitive pair uniqueness, safe temp/backup replacement, corruption-only reset confirmation, environmental read-error exit, and post-commit backup-cleanup warning behavior. Explicitly list credentials, queries, results, active selection, export paths, and window state as never persisted.
- **Query Policy:** Show the accepted logical shape and at least three valid examples: basic projection/WHERE, aggregate/GROUP BY, and ordered query using `CAST` or `TRY_CAST`. List exclusions: JOIN/APPLY, subquery/CTE, set operators, multiple statements, TOP/OFFSET/FETCH, INTO, temp/table variables, stored procedures/dynamic SQL, external table sources, and DML/DDL/transaction/permission/administrative commands. Mention aliases, CASE, aggregate/scalar expressions, nested comments, and dedicated CAST grammar. State the validator is an accidental-change safeguard, not a substitute for SQL Server least-privilege permissions.
- **Query Execution and Paging:** Explain ordered queries use parameterized app-added `OFFSET`/`FETCH`, 500 visible rows, and a 501st sentinel without `COUNT(*)` or total pages. Explain unordered queries retain at most configured limit plus one probe row, discard only the probe, display retained rows in local 500-row pages, and disable complete export if truncated. Explain stale-editor behavior and the stable unique ordering recommendation.
- **Excel Export:** Explain ordered export streams a fresh execution of the exact validated snapshot; complete unordered export uses all cached rows. Document bold filtered frozen header, neutral typed cells, null handling, formula-literal safety, `byte[]` uppercase `0x` hex, OOXML text escaping, early dates as ISO text, 32,767-character cells, 1,048,575 data-row limit, timeout, and atomic destination replacement with a non-null sibling backup. State incomplete/truncated unordered results cannot be exported as complete.
- **Security and Data Boundaries:** State Windows integrated authentication only, no credentials, connection strings via `SqlConnectionStringBuilder`, parameterized paging values, in-memory query/result state, app-local JSON writes, user-selected Excel writes, transient siblings, and SQL as the only automatic network protocol.
- **Error Handling and Runtime Model:** Cover synchronous busy state, temporary UI unresponsiveness, configured command/export timeout, fixed connection timeout, state clearing after query/page failure, destination preservation on export failure, and confirmations for delete/change/reset/overwrite.
- **Testing:** List the seven files under `tests/` by responsibility. Provide focused commands and the aggregate command:

  ```powershell
  powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
  powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
  powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
  ```

  State that no live database or Excel installation is required for automated tests.
- **Citrix Acceptance Checklist:** Use unchecked boxes for launch, real integrated connection and persistence, config reload, representative unordered complete/truncated and ordered queries, stable ordered paging, `.xlsx` opening/format confirmation, and filesystem-boundary confirmation. Label this external acceptance, not automated verification.
- **Version 1 Limitations:** List single named table, no joins/subqueries/CTEs/batches, one result set/query tab/connection, synchronous UI, no cancellation/progress/history/saved queries/total counts, and one worksheet only.
- **Future Extensions:** Identify explicit named-table `INNER JOIN`/`LEFT JOIN` grammar as the planned first policy extension and mention additional tabs/background execution only as possible later work, not commitments.
- **Design References:** Link relatively to the v1 design, v1 implementation plan, and documentation design.

- [ ] **Step 3: Validate README structure, paths, numeric contracts, and placeholders**

Run:

```powershell
$readme = Get-Content -LiteralPath .\README.md -Raw
$requiredHeadings = @(
    '# SQL Utility',
    '## Overview',
    '## Version 1 Features',
    '## Requirements and Portability',
    '## Portable Distribution',
    '## Running the Application',
    '## User Workflow',
    '## Architecture and Implementation',
    '## Configuration and Persistence',
    '## Query Policy',
    '## Query Execution and Paging',
    '## Excel Export',
    '## Security and Data Boundaries',
    '## Error Handling and Runtime Model',
    '## Testing',
    '## Citrix Acceptance Checklist',
    '## Version 1 Limitations',
    '## Future Extensions',
    '## Design References'
)
foreach ($heading in $requiredHeadings) {
    if ($readme -notmatch [regex]::Escape($heading)) { throw "Missing README heading: $heading" }
}
foreach ($requiredText in @('100', '2000', '1000', '3600', '120', '500', '501', '1,048,575', '32,767', 'Windows integrated')) {
    if ($readme -notmatch [regex]::Escape($requiredText)) { throw "Missing README contract: $requiredText" }
}
foreach ($path in @(
    'StartSqlUtility.cmd',
    'SqlUtility.ps1',
    'modules\SqlUtility.Config.ps1',
    'modules\SqlUtility.QueryPolicy.ps1',
    'modules\SqlUtility.Database.ps1',
    'modules\SqlUtility.Excel.ps1',
    'docs\superpowers\specs\2026-08-02-sql-utility-v1-design.md',
    'docs\superpowers\plans\2026-08-02-sql-utility-v1.md',
    'docs\superpowers\specs\2026-08-05-project-documentation-design.md'
)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "README target does not exist: $path" }
}
if ($readme -match '\b(?:TBD|TODO|PLACEHOLDER)\b') { throw 'README contains an unfinished placeholder.' }
Write-Host 'README validation passed.'
```

Expected: output `README validation passed.`.

- [ ] **Step 4: Run the aggregate suite**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
```

Expected: exit `0`; final line `All SQL Utility tests passed.`.

- [ ] **Step 5: Commit the README**

```powershell
git add README.md
git diff --cached --check
git commit -m "docs: add SQL Utility project guide"
```

---

### Task 2: Create Repository Agent Instructions

**Files:**

- Create: `AGENTS.md`
- Reference: `README.md`
- Reference: `docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md`
- Reference: `docs/superpowers/plans/2026-08-02-sql-utility-v1.md`
- Reference: `docs/superpowers/specs/2026-08-05-project-documentation-design.md`

**Interfaces:**

- Consumes: the project contracts and commands documented in `README.md`.
- Produces: repository-wide binding instructions for coding agents, readable by human maintainers.

- [ ] **Step 1: Verify AGENTS.md is currently absent and README.md exists**

Run:

```powershell
if (Test-Path -LiteralPath .\AGENTS.md) {
    throw 'AGENTS.md already exists; inspect it before following this create-only task.'
}
if (-not (Test-Path -LiteralPath .\README.md -PathType Leaf)) {
    throw 'Task 1 README.md must exist before writing AGENTS.md.'
}
Write-Host 'RED: AGENTS.md is absent and README.md is available.'
```

Expected: output `RED: AGENTS.md is absent and README.md is available.`.

- [ ] **Step 2: Create concise repository-wide agent instructions**

Create `AGENTS.md` with these exact top-level sections:

```markdown
# AGENTS.md

## Project Context
## Instruction Scope
## Required Reading
## Non-Negotiable Runtime Constraints
## Architecture Boundaries
## Behavioral Invariants
## Change Guidelines
## Testing and Verification
## Git and Workspace Safety
## Documentation Maintenance
## External Acceptance Boundary
```

Populate them with direct instructions:

- **Project Context:** One short paragraph explaining that SQL Utility is a portable internal PowerShell 5.1 WinForms client for Windows-authenticated SQL Server read-only querying, paging, and `.xlsx` export. Link to `README.md` for full behavior.
- **Instruction Scope:** State that the file applies to the whole repository and that a future nested `AGENTS.md` may add narrower instructions without silently weakening root safety constraints.
- **Required Reading:** Require reading `README.md`, the v1 design, the v1 implementation plan when historical task context is needed, and the documentation design for ownership rules. Use relative Markdown links.
- **Non-Negotiable Runtime Constraints:** Require `.cmd` launch, PowerShell 5.1, built-in .NET only, no installer/compiled EXE/admin/third party/Office automation/registry/environment mutation, Windows integrated authentication only, and no credentials.
- **Architecture Boundaries:** State exact ownership for `SqlUtility.ps1` and each module. Prohibit WinForms dependencies in Database/Excel, SQL ownership in Excel, database calls in QueryPolicy, and persistence outside Config/UI orchestration. Require neutral result/schema/row boundaries and deterministic disposal.
- **Behavioral Invariants:** Protect blank startup fields, app-local JSON only, safe replacement, setting ranges/defaults, 500-row pages, ordered 501-row sentinel paging, bounded unordered results, stale-query snapshot behavior, complete-result-only export, timeout enforcement, Excel row/text/data-fidelity limits, and query-policy scope. Require explicit user approval and regression tests before widening query grammar, especially JOIN support.
- **Change Guidelines:** Require a minimal-change proposal before non-trivial behavior changes, explicit handling of ambiguity, preservation of unrelated user work, no speculative features, no unrelated refactors, and regression-first TDD. State that approved future scope should extend existing boundaries rather than collapse modules.
- **Testing and Verification:** Require the relevant focused script, the aggregate PowerShell 5.1 command, launcher test for runtime/distribution changes, `git diff --check`, and clean/status review. Require tests for safety boundaries, failure paths, resource cleanup, and destination preservation. Do not claim success without fresh output.
- **Git and Workspace Safety:** Require non-destructive Git, awareness of dirty worktrees, no reset/checkout-over-user-changes, explicit staging, and `safe.directory` when this Windows checkout requires it.
- **Documentation Maintenance:** Require synchronized updates to README/design/plan/AGENTS when their owned contracts change, and prohibit copying stale plan snippets as current behavior.
- **External Acceptance Boundary:** Require Citrix/live SQL/cloud-drive/Desktop Excel checks to be labeled pending unless actually performed; automated fake/in-memory tests cannot satisfy them.

- [ ] **Step 3: Validate AGENTS.md scope, references, contracts, and concision**

Run:

```powershell
$agents = Get-Content -LiteralPath .\AGENTS.md -Raw
foreach ($heading in @(
    '# AGENTS.md',
    '## Project Context',
    '## Instruction Scope',
    '## Required Reading',
    '## Non-Negotiable Runtime Constraints',
    '## Architecture Boundaries',
    '## Behavioral Invariants',
    '## Change Guidelines',
    '## Testing and Verification',
    '## Git and Workspace Safety',
    '## Documentation Maintenance',
    '## External Acceptance Boundary'
)) {
    if ($agents -notmatch [regex]::Escape($heading)) { throw "Missing AGENTS heading: $heading" }
}
foreach ($requiredText in @('README.md', 'Windows PowerShell 5.1', 'Windows integrated', '500', '501', '100', '2000', '5', '3600', 'git diff --check', 'Test-All.ps1')) {
    if ($agents -notmatch [regex]::Escape($requiredText)) { throw "Missing AGENTS contract: $requiredText" }
}
if ($agents -match '\b(?:TBD|TODO|PLACEHOLDER)\b') { throw 'AGENTS.md contains an unfinished placeholder.' }
if (($agents -split "`n").Count -gt 240) { throw 'AGENTS.md is too long for a concise instruction layer.' }
Write-Host 'AGENTS validation passed.'
```

Expected: output `AGENTS validation passed.`.

- [ ] **Step 4: Cross-check documentation and run full verification**

Run:

```powershell
$readme = Get-Content -LiteralPath .\README.md -Raw
$agents = Get-Content -LiteralPath .\AGENTS.md -Raw
foreach ($contract in @('Windows PowerShell 5.1', 'Windows integrated', '500', '1000', '120', 'Citrix')) {
    if ($readme -notmatch [regex]::Escape($contract)) { throw "README missing shared contract: $contract" }
    if ($agents -notmatch [regex]::Escape($contract)) { throw "AGENTS missing shared contract: $contract" }
}
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git diff --check
git status --short
```

Expected: documentation checks pass; test output ends with `All SQL Utility tests passed.`; `git diff --check` is silent; status lists only `AGENTS.md` before staging.

- [ ] **Step 5: Commit AGENTS.md**

```powershell
git add AGENTS.md
git diff --cached --check
git commit -m "docs: add repository agent guidance"
git status --short
```

Expected: commit succeeds and final status is clean.
