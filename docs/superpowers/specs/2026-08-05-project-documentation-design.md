# SQL Utility Project Documentation Design

## Purpose

Create two repository-root documents that explain the completed SQL Utility v1 implementation and guide future work:

- `README.md` is the primary project, usage, design, and implementation reference.
- `AGENTS.md` contains concise repository-wide instructions for AI coding agents while remaining readable and useful to human maintainers.

The documents must describe the current implementation at the feature-branch head. They must not repeat obsolete implementation snippets from the original plan or claim that external Citrix acceptance has been completed.

## Documentation Model

Use a layered model to reduce duplication and maintenance drift.

`README.md` owns explanatory material: what the application does, how to run it, how its components interact, why key boundaries exist, and how to test it. `AGENTS.md` owns binding repository instructions: what an agent must preserve, how it should make and verify changes, and which documents provide deeper context.

`AGENTS.md` links to `README.md` and the existing v1 design and implementation plan instead of restating their detailed explanations. When a product or architecture contract changes, the responsible change must update the affected documentation in the same work.

## README.md Design

### Audience

The README serves three audiences:

1. Internal users copying and launching the portable application in Citrix.
2. Maintainers who need to understand the architecture and constraints.
3. Coding agents that need a reliable entry point before reading deeper design material.

### Required Content

The README contains:

- A concise introduction and version 1 capability summary.
- Runtime requirements and portability guarantees.
- The exact portable distribution layout and launch instructions.
- The end-user workflow for connection testing, connecting, query execution, paging, export, settings, saved connections, and changing connections.
- The Data Explorer workflow for physical-table discovery, bounded previews, structured filters, preview export, and generated-SQL handoff.
- A detailed implementation overview with a compact Mermaid component/data-flow diagram.
- The responsibility and dependency boundary of `SqlUtility.ps1` and each file in `modules/`.
- The application-local schema 2 configuration, schema 1 in-memory migration, validation ranges, persistence location, safe replacement behavior, and explicit list of data that is never stored.
- The accepted single-table read-only query shape, representative accepted examples, rejected constructs, and the warning that client validation does not replace SQL Server permissions.
- The two paging models: ordered server-side 500-row pages with a 501-row sentinel, and bounded unordered retrieval with local 500-row display pages.
- Complete-result Excel export behavior, availability rules, streaming/cached sources, workbook formatting, data-fidelity safeguards, Excel limits, and safe destination replacement.
- Authentication, security, privacy, filesystem-write, and network-write boundaries.
- Synchronous execution, timeout, confirmation, and error-handling behavior.
- The dependency-free PowerShell test layout and focused/aggregate commands.
- A cloud/Citrix acceptance checklist that remains explicitly pending until performed in that environment.
- Known version 1 limitations and future extension points, including explicit JOIN grammar as a later policy extension.
- Links to the authoritative v1 design, v1 implementation plan, and documentation design.

### Accuracy Rules

- Describe final code behavior, including release-review fixes, rather than planned-but-superseded implementation details.
- Use the actual setting ranges: unordered row limit `100` through `2000`, default `1000`; Query/Export timeout `5` through `3600` seconds, default `120`; fixed connection timeout `10` seconds.
- State that display pages are fixed at 500 rows.
- State that server and database fields start blank on every launch even when saved pairs exist.
- State that `SqlUtility.config.json` is absent from the distribution and is created beside the application only when a successful connection or settings operation needs persistence.
- State that the application uses Windows integrated authentication only and never stores credentials.
- Do not claim that Microsoft Excel is required to generate `.xlsx`; desktop Excel is used only for external acceptance/opening the result.
- Do not claim ordered queries receive a count or total-page calculation.

## AGENTS.md Design

### Audience and Authority

`AGENTS.md` is primarily a binding instruction file for AI coding agents operating anywhere in this repository. Its language remains plain enough for human contributors to understand the rationale and expected checks.

The file applies repository-wide unless a future nested `AGENTS.md` introduces more specific instructions for a subtree.

### Required Instructions

Agents must:

- Read `README.md` and the relevant design/plan documents before modifying behavior.
- Preserve the portable `StartSqlUtility.cmd` to Windows PowerShell 5.1 launch model.
- Use only Windows and .NET Framework components normally available in the target Citrix environment.
- Avoid third-party dependencies, installation, compiled executables, administrator access, Office automation, registry writes, and environment-variable writes.
- Preserve Windows integrated authentication and never introduce, request, log, or persist credentials.
- Respect the established module boundaries among UI/workflow, configuration, query policy, database access, and Excel export.
- Preserve the UI-neutral Data Explorer module boundary: catalog-derived identifiers and typed builder/preview descriptors belong there, while SQL execution and WinForms remain outside it.
- Treat the query policy as a safety boundary and obtain explicit approval before widening the single-statement, single-table, read-only `SELECT` scope.
- Preserve ordered/unordered paging semantics, complete-export eligibility, timeout enforcement, Excel limits, data fidelity, resource disposal, and safe file replacement.
- Limit persistence to validated app-local configuration/transient files and user-selected Excel destinations.
- Keep changes minimal and scoped; avoid unrelated refactoring and speculative features.
- Use regression-first test-driven development for behavior changes.
- Run the focused test file during development and `tests\Test-All.ps1` before completion.
- Verify Windows PowerShell 5.1 compatibility, launcher portability where relevant, temporary-file cleanup, `git diff --check`, and worktree status.
- Preserve unrelated and user-owned files and avoid destructive Git commands unless explicitly authorized.
- Update the relevant documentation when runtime, architecture, security, persistence, query-policy, paging, export, or verification contracts change.
- Report live SQL Server, Citrix, cloud-drive, and desktop Excel checks as external acceptance unless they were actually performed in that environment.

### Reference Strategy

`AGENTS.md` links to:

- `README.md` for the current application overview, architecture, workflows, and commands.
- `docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md` for the approved v1 product design.
- `docs/superpowers/plans/2026-08-02-sql-utility-v1.md` for implementation history and task decomposition.
- `docs/superpowers/specs/2026-08-05-project-documentation-design.md` for the documentation ownership model.
- `docs/superpowers/specs/2026-08-14-sql-utility-data-explorer-design.md` for the Data Explorer, schema 2, and seven-file runtime extension.

References are relative Markdown links so they work in local repository viewers and hosted Git forges.

## Verification

Documentation implementation is complete only when:

- Both files exist at the repository root.
- All referenced paths and commands exist in the checkout.
- The README distribution tree exactly matches the seven runtime files.
- Data Explorer architecture, `previewRowLimit`, and `tests\Test-DataExplorer.ps1` ownership match the extension design, production code, and aggregate runner.
- Setting ranges, defaults, timeouts, paging limits, Excel limits, and persistence behavior match production code and tests.
- No unfinished placeholder markers remain.
- `AGENTS.md` contains direct, testable instructions and does not conflict with `README.md`.
- Markdown links resolve to repository files.
- The full PowerShell test suite remains green even though the change is documentation-only.
- `git diff --check` reports no whitespace errors.

## Scope Boundaries

This documentation work does not change application behavior, packaging, configuration schema, query grammar, user interface, or tests. It does not merge the feature branch or complete the pending Citrix acceptance tests.
