# AGENTS.md

## Project Context

SQL Utility is a portable internal Windows PowerShell 5.1 WinForms client for Windows-authenticated SQL Server connection testing, restricted read-only querying, result paging, and dependency-free `.xlsx` export. Read [README.md](README.md) for current user behavior, architecture, and operating instructions.

## Instruction Scope

These instructions apply to the entire repository. A future nested `AGENTS.md` may add more specific guidance for its subtree, but it must not silently weaken the root portability, authentication, query-safety, persistence, or verification constraints.

This file is written primarily for AI coding agents. Its requirements are also the expected engineering practices for human contributors.

## Required Reading

Before changing behavior, read:

1. [README.md](README.md) for the current implementation and commands.
2. [SQL Utility Version 1 Design](docs/superpowers/specs/2026-08-02-sql-utility-v1-design.md) for approved product and architecture boundaries.
3. [SQL Utility Version 1 Implementation Plan](docs/superpowers/plans/2026-08-02-sql-utility-v1.md) when historical task decomposition or design rationale is relevant. Treat final code and newer review fixes as authoritative over superseded plan snippets.
4. [Project Documentation Design](docs/superpowers/specs/2026-08-05-project-documentation-design.md) when changing documentation ownership or structure.

Inspect the relevant production code and tests before proposing a change. Do not infer current behavior from filenames or old plans alone.

## Non-Negotiable Runtime Constraints

- Preserve `StartSqlUtility.cmd` as the user launch entry point and Windows PowerShell 5.1 as the runtime.
- Keep the process invocation portable: `-NoLogo -NoProfile -STA -ExecutionPolicy Bypass`. Do not change persistent execution policy.
- Use only Windows and .NET Framework components normally available in the target Citrix environment.
- Do not add an installer, compiled executable, third-party module, NuGet dependency, Python runtime, `sqlcmd` dependency, Office automation, or administrator requirement.
- Do not write to the registry or modify user/machine/process environment variables.
- Preserve Windows integrated authentication only. Never introduce, request, log, serialize, or persist usernames, passwords, tokens, or SQL-authentication credentials.
- Keep the runtime distribution to `StartSqlUtility.cmd`, `SqlUtility.ps1`, and the four required files under `modules/` unless the user explicitly approves a packaging change.

## Architecture Boundaries

Keep each file focused on its established ownership:

- `StartSqlUtility.cmd`: resolve and start the sibling PowerShell application; no business logic.
- `SqlUtility.ps1`: WinForms construction, application state, workflow orchestration, service boundaries, result binding, paging controls, status, and user messages.
- `modules/SqlUtility.Config.ps1`: configuration defaults/schema validation, saved-pair operations, JSON reads, and safe app-local writes.
- `modules/SqlUtility.QueryPolicy.ps1`: SQL tokenization, single-table read-only grammar, normalization, table extraction, and top-level ordering detection.
- `modules/SqlUtility.Database.ps1`: connection strings, SQL connections/commands/readers, diagnostic tests, paging, neutral result conversion, ordered streaming, timeouts, cancellation where possible, and deterministic disposal.
- `modules/SqlUtility.Excel.ps1`: neutral schema/row input to `.xlsx`, workbook limits, OOXML data fidelity, timeout checks, and safe destination replacement.

Do not add WinForms dependencies to the Database or Excel modules. Do not let Excel own SQL connections, commands, or readers. Do not add database access to QueryPolicy. Do not let Database render controls. Pass data through neutral `DataTable`, page metadata, schema, and row-value boundaries.

Preserve deterministic cleanup of SQL connections, commands, readers, streams, ZIP packages, XML writers, temporary files, and backups. Cleanup errors must not mask the primary failure or misreport a committed write as failed.

## Behavioral Invariants

- Server and Database fields start blank on every launch, even when saved pairs exist.
- Only a successful connection test may add/persist a server/database pair. Pairs are unique case-insensitively.
- Persistent application state is limited to validated `SqlUtility.config.json` and its same-directory safe-write transients. Export files/transients use only the user-selected destination directory.
- Configuration safe writes preserve the prior file on pre-commit failure. Only actual parse/schema corruption may offer reset; environmental read failures exit unchanged.
- `unorderedRowLimit` is `100` through `2000`, default `1000`.
- `queryExportTimeoutSeconds` is `5` through `3600`, default `120`. Connection timeout stays fixed at `10` seconds.
- Display pages remain fixed at `500` rows.
- Ordered paging uses application-controlled parameterized `OFFSET`/`FETCH` and a `501`-row fetch sentinel. Do not add an implicit `COUNT(*)` or total-page query.
- Unordered execution retains at most the configured limit after probing one additional row. Local paging must not repeat the SQL query.
- Paging and export use the exact last successful normalized query snapshot. Editor changes make results stale.
- Export is allowed only when the complete result is available: ordered results stream a fresh complete execution; complete unordered results use the full cache; incomplete/truncated unordered results must be rejected before prompting.
- Excel export must preserve the existing destination until a complete workbook is ready. Preserve row/text limits, formula-literal safety, binary hex fidelity, OOXML escaping, early-date handling, and timeout cleanup.
- The query policy remains one statement, one named table, and read-only `SELECT`. It excludes joins, subqueries, CTEs, set operators, batches, stored/dynamic SQL, external table sources, and data-changing, DDL, transaction, permission, or administrative commands.
- Treat query validation as an accidental-change safety boundary, not a replacement for least-privileged SQL Server permissions.

Widening query grammar—especially adding `INNER JOIN` or `LEFT JOIN`—requires explicit user approval, a grammar design, bypass-focused regression tests, and preservation of all existing rejected cases.

## Change Guidelines

- Start non-trivial changes with a minimal-change proposal: affected behavior, files, invariants, tests, and any external acceptance needed.
- Clarify ambiguity before implementing when different answers would materially change behavior or scope.
- Change only what the approved request requires. Avoid speculative features, unrelated refactoring, and module-boundary collapse.
- Preserve unrelated and user-owned changes in a dirty worktree. Inspect status and diffs before editing or staging.
- Use regression-first test-driven development for behavior changes: reproduce the failure, make the smallest correction, run the focused suite, then run the aggregate suite.
- Extend existing injected service/executor boundaries for tests instead of adding production-only hooks when real filesystem or runtime behavior can be exercised safely.
- Prefer `rg`/`rg --files` for repository searches and `apply_patch` for text edits.
- Keep Windows PowerShell 5.1 syntax and .NET Framework behavior in mind; do not rely on `pwsh`-only language/runtime features.

## Testing and Verification

Run the focused test that owns the changed contract during development, for example:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Config.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Database.ps1
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Excel.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-SqlUtilityUi.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Launcher.ps1
```

Before completion, run the aggregate suite from the repository root:

```powershell
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
```

Also run `git diff --check` and inspect `git status --short`. For launcher or distribution changes, test from a working directory outside the application folder. For persistence/export changes, cover success, pre-commit failure, post-commit cleanup behavior, existing-destination preservation, and transient cleanup.

Do not claim that work is complete or passing without fresh command output from the tree being handed off.

## Git and Workspace Safety

- Use a feature branch/worktree for substantial changes; do not implement directly on `main` without explicit approval.
- This checkout may require `git -c safe.directory='<absolute-worktree-path>' ...` because of Windows ownership metadata.
- Do not use destructive commands such as `git reset --hard`, force checkout over changes, branch force-deletion, or force-push unless the user explicitly authorizes the exact action.
- Stage only intentional files. Review the staged diff and `git diff --cached --check` before committing.
- Do not delete, overwrite, move, or commit unrelated untracked files.
- Keep commits scoped and descriptive. Do not merge, push, or remove a worktree without the user's integration choice.

## Documentation Maintenance

Keep documentation synchronized with the contract it owns:

- Update `README.md` when user workflow, runtime files, settings, architecture, query policy, paging, export, security, testing, or limitations change.
- Update the approved design before implementing a material product/architecture scope change.
- Update implementation plans when future work is decomposed, but do not present stale plan snippets as current behavior.
- Update `AGENTS.md` when repository-wide engineering or verification rules change.
- Link to the owning document instead of copying long explanations that can drift.

## External Acceptance Boundary

Local fake/in-memory/ZIP/XML tests do not prove cloud-environment behavior. Label Citrix launch, live SQL Server connectivity, cloud-drive write authorization, configuration reload, and opening exported workbooks in desktop Excel as pending external acceptance unless those checks were actually performed in that environment.
