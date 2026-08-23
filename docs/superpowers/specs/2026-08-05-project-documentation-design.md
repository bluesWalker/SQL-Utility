# SQL Utility Project Documentation Design

## Purpose

Use three repository-root documents with distinct audiences:

- `README.md` is the end-user manual displayed on the repository home page.
- `TECHNICAL_REFERENCE.md` is the detailed project, implementation, architecture, verification, and maintainer reference. It was created from the previous combined README so that content was preserved when the user manual became the primary README.
- `AGENTS.md` contains binding repository-wide instructions for coding agents while remaining readable by human maintainers.

The documents describe current behavior. They must not repeat obsolete implementation snippets or claim that external Citrix acceptance has been completed.

## Documentation model

Use a layered model to reduce duplication and maintenance drift:

- The user manual owns installation, connection, saved-connection management, every application tab, paging, counting, export, settings, user-visible limits, and troubleshooting.
- The technical reference owns the implementation overview, module boundaries, configuration details, query-policy grammar, security and persistence boundaries, test commands, external acceptance checks, and future engineering work.
- `AGENTS.md` owns rules contributors must preserve and links to the user manual, technical reference, and deeper design documents.

The README must link prominently to `TECHNICAL_REFERENCE.md`. Both files remain at the repository root so users can reach the technical reference from the GitHub repository home page and see it in the root file list.

## README.md requirements

### Audience

The README is written for people who install and operate SQL Utility, including internal users running it in Citrix. It must not require knowledge of the codebase.

### Required content

The README contains only information needed to use the application:

- What SQL Utility does and its read-only purpose.
- The exact seven-file portable layout, runtime requirement, and launch steps.
- How app-folder write access affects saved connections and settings.
- Windows-authenticated connection testing, connecting, saved pairs, deletion, and changing connections.
- Complete Data Explorer instructions: table discovery/filtering/refresh, output-column interactions, structured filters, previews, preview export, and Query handoff.
- Complete Query instructions: supported user-facing statement shape, execution, stale results, ordered and unordered paging, explicit counts, and complete-result export.
- Complete Settings instructions with current ranges, defaults, and effects.
- User-facing limits, performance cautions, persistence expectations, and troubleshooting.
- A prominent link to `TECHNICAL_REFERENCE.md` for non-user material.

Do not add architecture diagrams, module ownership, implementation details, developer test commands, internal safe-write algorithms, development plans, or contribution instructions to the README.

### Accuracy rules

- Describe final code behavior rather than plans.
- Keep button names and tab names identical to the UI.
- State that server and database fields start blank on every launch.
- State that authentication uses the signed-in Windows identity and that credentials are not requested or stored.
- State that Data Explorer previews are bounded, unordered snapshots and that Query export requires a complete result.
- State that display pages are fixed at 500 rows and counts are explicit.
- Use the actual setting ranges and defaults.
- Do not claim that Microsoft Excel is required to generate `.xlsx` files.

## TECHNICAL_REFERENCE.md requirements

The technical reference preserves and maintains the detailed material previously owned by README:

- Capability and runtime summaries.
- Architecture, module ownership, and data flow.
- Configuration schema, migration, and safe persistence.
- Data Explorer and query-policy contracts.
- Paging, count, and export implementation behavior.
- Authentication, security, filesystem, network, and resource-lifecycle boundaries.
- Error handling and synchronous runtime behavior.
- Automated verification commands and the pending Citrix acceptance checklist.
- Version limitations, future extensions, and design references.

When user-visible behavior changes, update both the README instructions and the corresponding technical contract where applicable. Prefer links over duplicating long internal explanations in README.

## AGENTS.md requirements

`AGENTS.md` must direct contributors to read README for user behavior and `TECHNICAL_REFERENCE.md` for implementation details. It must preserve the runtime, architecture, authentication, query-safety, persistence, verification, Git, and external-acceptance rules of the repository.

Its Documentation Maintenance section must keep ownership explicit: user-facing operating content belongs in README, while implementation, architecture, security, and verification content belongs in the technical reference.

## Verification

Documentation changes are complete only when:

- `README.md`, `TECHNICAL_REFERENCE.md`, and `AGENTS.md` exist at the repository root.
- README links directly to the technical reference.
- README contains all current screens and tabs but no maintainer-only sections.
- Technical details removed from README remain available in the technical reference.
- All repository-relative Markdown links resolve.
- Runtime filenames, labels, ranges, defaults, timeouts, paging limits, and Excel limits match production code and tests.
- No unfinished placeholder markers remain.
- The full PowerShell test suite remains green.
- `git diff --check` reports no whitespace errors.

## Scope boundaries

This documentation model does not change application behavior, packaging, configuration schema, query grammar, user interface, or tests. It does not merge a branch or complete pending Citrix acceptance tests.
