# SQL Utility JOIN Support Design

**Date:** 2026-08-12

**Status:** Approved for implementation planning

## Purpose

Extend SQL Utility's restricted read-only query policy to support ordinary reporting queries that combine multiple named tables. The feature adds a deliberately narrow grammar for chained `INNER JOIN` and `LEFT JOIN` clauses without changing the application's authentication, execution, paging, counting, export, persistence, portability, or user-interface boundaries.

This is a read-only scope expansion. `UPDATE` and every other data-modification statement remain explicitly out of scope. Guarded data modification requires a separate safety framework and design.

## Approved Scope

An accepted query continues to be exactly one read-only `SELECT` statement. Its table-source portion may contain:

- One primary `[schema.]table` with an optional alias.
- Zero or more named-table joins.
- Bare `JOIN`, treated as `INNER JOIN`.
- Explicit `INNER JOIN`.
- `LEFT JOIN` and `LEFT OUTER JOIN`.
- A required, nonempty `ON` predicate for every joined source.
- The existing predicate language inside each `ON` clause, including qualified columns, comparisons, arithmetic, functions, literals, `AND`/`OR`, and parentheses.

Multiple join types may be mixed in one chain. No configured join-count limit is added.

The following example is within scope:

```sql
SELECT o.Id, c.Name, r.RegionName
FROM dbo.Orders AS o
LEFT OUTER JOIN dbo.Customers AS c
    ON c.Id = o.CustomerId AND c.Enabled = 1
INNER JOIN dbo.Regions AS r
    ON r.Id = c.RegionId
WHERE o.CreatedAt >= '2026-01-01'
ORDER BY o.Id;
```

## Explicitly Excluded

The validator continues to reject:

- `RIGHT JOIN`, `RIGHT OUTER JOIN`, `FULL JOIN`, `FULL OUTER JOIN`, and `CROSS JOIN`.
- `CROSS APPLY`, `OUTER APPLY`, and comma-separated table sources.
- Derived tables and parenthesized table sources.
- Subqueries and nested `SELECT` tokens at any parenthesis depth, including inside `ON` predicates.
- Common table expressions.
- Table-valued functions and external or remote rowset functions.
- Temporary tables, table variables, table hints, and query hints.
- Three-part or four-part table names and other cross-database or linked-server sources.
- `UNION`, `INTERSECT`, `EXCEPT`, `TOP`, user-supplied `OFFSET`/`FETCH`, and `INTO`.
- Multiple statements and executable content appended with or without a semicolon.
- `INSERT`, `UPDATE`, `DELETE`, `MERGE`, and all existing denied DDL, transaction, permission, administrative, stored, and dynamic SQL forms.

The query policy remains an accidental-change safeguard rather than a SQL security sandbox. SQL Server permissions assigned to the signed-in Windows identity remain authoritative, and deployments should continue using least-privileged identities.

## Architecture

### Production ownership

`modules/SqlUtility.QueryPolicy.ps1` is the only production file that needs behavioral changes. It continues to own tokenization, grammar validation, normalization, primary-table extraction, top-level ordering detection, and count-source generation.

The other runtime boundaries remain unchanged:

- `SqlUtility.ps1` continues to consume the existing validation-result shape and orchestrate execution, paging, counting, and export.
- `modules/SqlUtility.Database.ps1` continues to execute only validated normalized SQL and return neutral results. It does not parse joins.
- `modules/SqlUtility.Excel.ps1` continues to consume neutral schema and rows without database ownership.
- `modules/SqlUtility.Config.ps1` and `StartSqlUtility.cmd` are unaffected.

### Reusable named-source parser

The current primary-table parsing logic will be extracted into a reusable named-source helper. Given a token index, the helper will validate:

```text
[schema.]table [AS] alias
```

The alias remains optional. Identifiers may use the existing unquoted, bracketed, or double-quoted forms. The helper returns enough information for its caller to continue parsing, including the next unconsumed token index and the exact source-identifier range. It does not query SQL Server metadata or resolve whether the object is a physical table or view.

The primary source's exact identifier continues to populate the existing `TableIdentifier` validation-result property. Joined identifiers are validated but are not added to the public result contract because no current consumer needs them.

### Join-chain parser

After the primary source, QueryPolicy repeatedly recognizes one of these depth-zero sequences:

```text
JOIN
INNER JOIN
LEFT JOIN
LEFT OUTER JOIN
```

For every recognized join prefix, the parser:

1. Parses one named source with the reusable source helper.
2. Requires the depth-zero keyword `ON`.
3. Locates the end of the `ON` token range.
4. Requires at least one predicate token.
5. Validates that range with the existing expression-list validator, with top-level comma and ordering syntax disabled.
6. Continues at the next join or trailing query clause.

An `ON` predicate ends at the first subsequent depth-zero join-prefix sequence or at the first depth-zero `WHERE`, `GROUP BY`, `HAVING`, or `ORDER BY` boundary. Parenthesized content cannot create a false boundary. Quoted identifiers, strings, and comments remain non-executable token forms.

Keyword recognition must use complete sequences. For example, `LEFT(value, 2)` remains an ordinary function expression because `LEFT` is not followed by the depth-zero join sequence. A quoted `[LEFT]` or the word `JOIN` inside a string or comment also cannot start a join.

After the join chain, existing trailing-clause parsing remains responsible for clause ordering and expression validation.

### Validation-result compatibility

Valid joined queries return the existing properties without additions:

- `IsValid`
- `ErrorMessage`
- `NormalizedSql`
- `TableIdentifier`
- `HasOrderBy`
- `CountSourceSql`

`NormalizedSql` contains the complete unpaged joined query with only the allowed trailing statement semicolon removed according to current behavior. `HasOrderBy` continues to represent only a depth-zero display `ORDER BY`. `CountSourceSql` continues to remove only that top-level display ordering and otherwise preserves the full joined query.

Invalid joined queries return the same safe blank executable fields as every other invalid query and never reach the database module.

## Execution and Result Behavior

No new execution mode is introduced.

### Ordered queries

For a joined query with top-level `ORDER BY`, the database module appends the existing application-controlled parameterized `OFFSET` and `FETCH`. The visible page remains 500 rows and the fetch remains 501 rows so the final row can act as the next-page sentinel. There is no implicit count.

### Unordered queries

For a joined query without top-level `ORDER BY`, execution retains at most `unorderedRowLimit` rows after reading one additional truncation-probe row. Local display paging reuses the cache and does not rerun the join.

These limits bound transferred and retained rows. They cannot prevent SQL Server from scanning large sources, producing large intermediate join results, sorting, grouping, blocking, or timing out.

### Explicit count

The existing count wrapper receives the joined `CountSourceSql`. Its generated derived-table column alias list continues to protect wildcard, unnamed, and duplicate output-column cases. Count remains explicit, potentially expensive, and point-in-time.

### Export

Ordered joined results are freshly re-executed and streamed through the existing neutral schema/row boundary. Complete unordered joined results export from the existing cache. Truncated unordered results remain ineligible for export.

Joined projections commonly return duplicate headings such as multiple `Id` columns. `New-SqlUtilityResultTable` already makes duplicate and blank names unique in reader ordinal order, so the database and Excel contracts require no JOIN-specific change.

## Error Handling

Structural JOIN failures are reported as query-validation warnings before any database call. Messages should identify the relevant missing or unsupported component, such as an unsupported join type, malformed named source, missing `ON`, or empty/invalid `ON` predicate.

QueryPolicy does not resolve tables, views, aliases, columns, collations, conversions, or data types. SQL Server semantic failures, including unknown or ambiguous columns, continue through the existing Query Error path.

Execution failures preserve current state behavior:

- The failed execution does not produce a successful query snapshot.
- Current result, count, paging, and export state are cleared.
- Busy-state restoration occurs in the existing `finally` path.
- Editing a successful query continues to make the result stale. Restoring the prior editor text does not clear stale state; the user must execute again, according to current behavior.

No preflight execution, automatic row count, metadata lookup, persistent query history, or operational log is added.

## Testing Strategy

Development follows regression-first test-driven development in `tests/Test-QueryPolicy.ps1`. A representative accepted JOIN test must fail under the current policy before the production grammar is changed.

### Accepted-query coverage

Tests will cover:

- Bare `JOIN`, explicit `INNER JOIN`, `LEFT JOIN`, and `LEFT OUTER JOIN`.
- Multiple chained joins and mixed inner/left chains.
- Sources with and without `AS` aliases.
- Unquoted, bracketed, and double-quoted identifiers.
- Compound `ON` conditions using `AND`, `OR`, comparisons, ranges, functions, literals, and parentheses.
- Non-equality joins.
- `LEFT(...)` as a function without a false join boundary.
- Joined queries followed by `WHERE`, `GROUP BY`, `HAVING`, and `ORDER BY` in the approved order.
- Top-level ordering detection and order-free count-source generation.
- Trailing semicolon normalization with trailing line and block comments.
- Join-related keywords inside strings, comments, and quoted identifiers.

Every accepted result must retain the current validation-result types and exact normalized SQL. `TableIdentifier` must remain the primary source identifier.

### Rejected-query and bypass coverage

Tests will cover:

- Every excluded join and apply form.
- Comma-separated sources.
- Missing or malformed source, alias, `JOIN`, `ON`, or predicate components.
- Empty `ON` predicates before another join or trailing clause.
- Invalid join prefixes and invalid join tokens in other query locations.
- Malformed chained boundaries.
- Parenthesized and derived sources, table functions, hints, temporary tables, table variables, and three-part sources.
- Nested `SELECT` in an `ON` predicate and CTE-based join attempts.
- Additional statements appended after a syntactically valid join, with and without semicolons.
- DML, DDL, transaction, permission, administrative, stored, dynamic, and other existing denied tokens placed around join boundaries.
- Unclosed comments, quotes, identifiers, and parentheses around joined source text.

All existing accepted and rejected query-policy cases remain regression requirements. The aggregate suite must stay green even though no non-policy production module is expected to change.

## Documentation Ownership

After implementation:

- `README.md` will describe the current accepted JOIN grammar, examples, limitations, query-cost boundary, architecture, test coverage, and Citrix acceptance checklist.
- `AGENTS.md` will replace its repository-wide single-table invariant with the exact approved named-source and JOIN boundary. It will continue requiring explicit design approval and bypass-focused tests for any future grammar widening.
- The original SQL Utility Version 1 design remains historical evidence of the original single-table scope and is not rewritten.
- A separate implementation plan will own task decomposition and exact regression-first edit order.

## Verification and Acceptance

Local completion requires fresh output from:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

Local tests do not prove real SQL Server or Citrix behavior. External acceptance remains pending until the portable distribution is exercised in the intended environment with representative queries that cover:

- Inner and left join execution.
- Mixed chained joins.
- Ordered paging across more than one page.
- Unordered truncation behavior.
- Explicit count.
- Complete Excel export and opening the workbook in desktop Excel.
- A rejected unsupported join and a server-side semantic error.

## Deferred Work

The following require separate designs and are not implied by this feature:

- `RIGHT`, `FULL`, or `CROSS JOIN`.
- `APPLY`, derived tables, subqueries, CTEs, set operators, or table functions.
- Cross-database or remote sources.
- Background execution or cancellation improvements.
- Guarded `UPDATE` or any other data-modification framework.
