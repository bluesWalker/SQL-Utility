# SQL Utility JOIN Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict read-only support for unlimited chained named-table bare/`INNER` and `LEFT` joins while preserving all existing query-safety, execution, paging, count, export, portability, and authentication behavior.

**Architecture:** Extend only `modules/SqlUtility.QueryPolicy.ps1` in production. Extract reusable named-source parsing, recognize complete allowed join-prefix sequences at depth zero, and validate each required `ON` range with the existing expression validator before handing the unchanged normalized SELECT contract to the UI and database layers. Add regression-first accepted and bypass-focused rejected cases, then synchronize the README and repository-wide agent rules.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework, existing token-based SQL policy, dependency-free PowerShell regression scripts, Git.

## Global Constraints

- Preserve `StartSqlUtility.cmd` as the launch entry point, Windows PowerShell 5.1 as the runtime, and `-NoLogo -NoProfile -STA -ExecutionPolicy Bypass` as the portable process invocation.
- Add no installer, executable, third-party module, NuGet dependency, Python runtime, `sqlcmd`, Office automation, administrator requirement, registry write, environment-variable change, SQL authentication, credential storage, or packaging change.
- Continue accepting exactly one read-only `SELECT` statement. `INSERT`, `UPDATE`, `DELETE`, `MERGE`, and every existing denied DDL, transaction, permission, administrative, stored, and dynamic SQL form remain rejected.
- Permit zero or more bare `JOIN`, explicit `INNER JOIN`, `LEFT JOIN`, and `LEFT OUTER JOIN` units. Every joined source is a one- or two-part named source with an optional alias and a required nonempty `ON` predicate.
- Continue rejecting `RIGHT`, `FULL`, and `CROSS JOIN`; both forms of `APPLY`; comma joins; subqueries; CTEs; derived sources; table functions; hints; temporary/table-variable sources; three- or four-part names; remote sources; and additional statements.
- Reuse the existing expression grammar for `ON`; do not add SQL Server metadata lookups, alias/column resolution, automatic counts, preflight execution, or new public validation properties.
- Preserve the 500-row display size, ordered parameterized `OFFSET`/`FETCH` with the 501-row sentinel, bounded unordered cache, exact-query snapshot, explicit count behavior, and complete-result export rules.
- Preserve deterministic SQL/stream/file cleanup and existing module ownership. No WinForms changes and no database or Excel parsing responsibilities are allowed.
- Use regression-first TDD: add a focused failing assertion, run the focused suite and confirm the intended failure, implement the smallest correction, and rerun the focused suite.
- Preserve unrelated and user-owned changes. Review status and diffs before staging, and stage only the files named by the current task.
- Treat Citrix launch, real SQL Server JOIN execution, paging, counting, export, and opening the workbook in desktop Excel as pending external acceptance unless performed there.

---

## Interface and File Map

| File | Planned contract change |
| --- | --- |
| `modules/SqlUtility.QueryPolicy.ps1` | Add internal named-source and join-prefix helpers; allow and validate the approved join chain; keep the public validation-result shape unchanged. |
| `tests/Test-QueryPolicy.ps1` | Add accepted JOIN coverage, exact normalization/count-source assertions, unsupported-form rejection, malformed-boundary cases, and bypass-focused regressions. |
| `README.md` | Describe current JOIN grammar, examples, cost/safety limits, architecture, tests, and external acceptance. |
| `AGENTS.md` | Replace the single-table invariant with the exact approved JOIN boundary and retain an approval/design/test gate for future grammar widening. |

No changes are planned for `SqlUtility.ps1`, `modules/SqlUtility.Database.ps1`, `modules/SqlUtility.Excel.ps1`, `modules/SqlUtility.Config.ps1`, the launcher, or their tests. The existing validation contract remains:

```powershell
[pscustomobject][ordered]@{
    IsValid = $true
    ErrorMessage = ''
    NormalizedSql = $normalizedSql
    TableIdentifier = $primaryTableIdentifier
    HasOrderBy = [bool] $hasOrderBy
    CountSourceSql = $countSourceSql
}
```

The new helpers are internal QueryPolicy functions:

```powershell
function Read-SqlUtilityNamedSource {
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start
    )
    # Returns IsValid, ErrorMessage, NextIndex, IdentifierStart, IdentifierEnd.
}

function Get-SqlUtilityJoinPrefix {
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start
    )
    # Returns Kind ('None', 'Allowed', or 'Unsupported'), Length, and ErrorMessage.
}
```

`Read-SqlUtilityNamedSource` parses `[schema.]table [AS] alias`, consuming an implicit alias only when `Test-SqlUtilityAliasToken` accepts the next token. `Get-SqlUtilityJoinPrefix` recognizes complete depth-zero sequences only; it must not classify `LEFT(value, 2)`, quoted identifiers, strings, comments, or nested tokens as join boundaries.

---

### Task 1: Accept the approved named-table JOIN grammar

**Files:**
- Modify: `tests/Test-QueryPolicy.ps1:8-75`
- Modify: `modules/SqlUtility.QueryPolicy.ps1:276-307,793-1029`

**Interfaces:**
- Consumes: `Get-SqlUtilitySqlTokens`, `Test-SqlUtilityIdentifierToken`, `Test-SqlUtilityAliasToken`, and `Test-SqlUtilityExpressionList` from the existing QueryPolicy module.
- Produces: internal `Read-SqlUtilityNamedSource -Tokens -Start`, internal `Get-SqlUtilityJoinPrefix -Tokens -Start`, and the unchanged `Test-SqlUtilityQuery -Sql` result contract.

- [ ] **Step 1: Add representative accepted JOIN cases to the `$accepted` table**

Insert these cases before the closing `)` of `$accepted`:

```powershell
    @{ Name = 'bare inner join'; Sql = 'SELECT i.Id, c.Name FROM dbo.Items i JOIN dbo.Categories c ON c.Id = i.CategoryId'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id, c.Name FROM dbo.Items i JOIN dbo.Categories c ON c.Id = i.CategoryId' }
    @{ Name = 'explicit inner join'; Sql = 'SELECT i.Id FROM dbo.Items AS i INNER JOIN dbo.Other AS o ON o.Id = i.Id'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id FROM dbo.Items AS i INNER JOIN dbo.Other AS o ON o.Id = i.Id' }
    @{ Name = 'left join'; Sql = 'SELECT i.Id, c.Name FROM dbo.Items i LEFT JOIN dbo.Categories c ON c.Id = i.CategoryId'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id, c.Name FROM dbo.Items i LEFT JOIN dbo.Categories c ON c.Id = i.CategoryId' }
    @{ Name = 'left outer join'; Sql = 'SELECT i.Id, c.Name FROM dbo.Items i LEFT OUTER JOIN dbo.Categories c ON c.Id = i.CategoryId'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id, c.Name FROM dbo.Items i LEFT OUTER JOIN dbo.Categories c ON c.Id = i.CategoryId' }
    @{ Name = 'mixed chained joins'; Sql = "SELECT o.Id, c.Name, r.RegionName FROM dbo.Orders o LEFT OUTER JOIN dbo.Customers c ON c.Id = o.CustomerId AND c.Enabled = 1 INNER JOIN dbo.Regions r ON r.Id = c.RegionId WHERE o.CreatedAt >= '2026-01-01' ORDER BY o.Id"; Table = 'dbo.Orders'; Ordered = $true; Normalized = "SELECT o.Id, c.Name, r.RegionName FROM dbo.Orders o LEFT OUTER JOIN dbo.Customers c ON c.Id = o.CustomerId AND c.Enabled = 1 INNER JOIN dbo.Regions r ON r.Id = c.RegionId WHERE o.CreatedAt >= '2026-01-01' ORDER BY o.Id" }
    @{ Name = 'quoted joined sources and aliases'; Sql = 'SELECT "i"."Id" FROM [dbo].[Items] AS [i] INNER JOIN "dbo"."Other" AS "o" ON "o"."Id" = [i].[Id]'; Table = '[dbo].[Items]'; Ordered = $false; Normalized = 'SELECT "i"."Id" FROM [dbo].[Items] AS [i] INNER JOIN "dbo"."Other" AS "o" ON "o"."Id" = [i].[Id]' }
    @{ Name = 'compound range and function join'; Sql = "SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON (o.Minimum <= i.Score AND i.Score < o.Maximum) OR LEFT(o.Code, 2) = LEFT(i.Code, 2)"; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON (o.Minimum <= i.Score AND i.Score < o.Maximum) OR LEFT(o.Code, 2) = LEFT(i.Code, 2)" }
    @{ Name = 'grouped joined query'; Sql = 'SELECT c.Id, COUNT(*) AS ItemCount FROM dbo.Items i LEFT JOIN dbo.Categories c ON c.Id = i.CategoryId GROUP BY c.Id HAVING COUNT(*) > 1 ORDER BY c.Id;'; Table = 'dbo.Items'; Ordered = $true; Normalized = 'SELECT c.Id, COUNT(*) AS ItemCount FROM dbo.Items i LEFT JOIN dbo.Categories c ON c.Id = i.CategoryId GROUP BY c.Id HAVING COUNT(*) > 1 ORDER BY c.Id' }
    @{ Name = 'joined final semicolon before line comment'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id; -- trailing JOIN comment'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id -- trailing JOIN comment' }
    @{ Name = 'joined final semicolon before block comment'; Sql = "SELECT i.Id FROM dbo.Items i LEFT JOIN dbo.Other o ON o.Id = i.Id ;`r`n/* trailing RIGHT JOIN comment ; */  "; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT i.Id FROM dbo.Items i LEFT JOIN dbo.Other o ON o.Id = i.Id `r`n/* trailing RIGHT JOIN comment ; */  " }
```

The existing table-driven assertions already verify validity, Boolean types, exact normalized SQL, primary `TableIdentifier`, and top-level ordering.

- [ ] **Step 2: Add exact joined count-source assertions**

Add after the existing `$orderInExpression` assertion:

```powershell
$orderedJoin = Test-SqlUtilityQuery -Sql @'
SELECT o.Id, c.Name
FROM dbo.Orders o
LEFT JOIN dbo.Customers c ON c.Id = o.CustomerId
ORDER BY o.Id;
'@
Assert-Equal @'
SELECT o.Id, c.Name
FROM dbo.Orders o
LEFT JOIN dbo.Customers c ON c.Id = o.CustomerId
'@ $orderedJoin.CountSourceSql 'Joined count source removes only top-level ORDER BY'

$unorderedJoin = Test-SqlUtilityQuery -Sql 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id;'
Assert-Equal 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id' `
    $unorderedJoin.CountSourceSql 'Unordered joined count source keeps the complete normalized query'
```

- [ ] **Step 3: Run the focused suite and verify the new cases fail for the intended reason**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
```

Expected: FAIL on the first new JOIN acceptance assertion because the current denied-token list reports `JOIN` as not allowed. Record that failure before editing production code.

- [ ] **Step 4: Add the reusable named-source helper**

Place `Read-SqlUtilityNamedSource` after `Test-SqlUtilityAliasToken`. Implement the complete result shape and parse sequence below using PowerShell 5.1-compatible syntax:

```powershell
function Read-SqlUtilityNamedSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start
    )

    $invalid = {
        param([string] $Message)
        [pscustomobject][ordered]@{
            IsValid = $false
            ErrorMessage = $Message
            NextIndex = $Start
            IdentifierStart = -1
            IdentifierEnd = -1
        }
    }

    if ($Start -ge $Tokens.Count -or $Tokens[$Start].Depth -ne 0 -or
        -not (Test-SqlUtilityIdentifierToken -Token $Tokens[$Start])) {
        return & $invalid 'A table source must begin with a one- or two-part table identifier.'
    }

    $index = $Start
    $identifierStart = $Tokens[$index].Start
    $identifierEnd = $Tokens[$index].End
    $index++

    if ($index -lt $Tokens.Count -and $Tokens[$index].Depth -eq 0 -and
        $Tokens[$index].Kind -eq 'Symbol' -and $Tokens[$index].Text -eq '.') {
        $index++
        if ($index -ge $Tokens.Count -or $Tokens[$index].Depth -ne 0 -or
            -not (Test-SqlUtilityIdentifierToken -Token $Tokens[$index])) {
            return & $invalid 'The table identifier is malformed.'
        }
        $identifierEnd = $Tokens[$index].End
        $index++
    }

    if ($index -lt $Tokens.Count -and $Tokens[$index].Depth -eq 0 -and
        $Tokens[$index].Kind -eq 'Word' -and $Tokens[$index].Upper -eq 'AS') {
        $index++
        if ($index -ge $Tokens.Count -or $Tokens[$index].Depth -ne 0 -or
            -not (Test-SqlUtilityAliasToken -Token $Tokens[$index])) {
            return & $invalid 'AS must be followed by a table alias.'
        }
        $index++
    }
    elseif ($index -lt $Tokens.Count -and $Tokens[$index].Depth -eq 0 -and
        (Test-SqlUtilityAliasToken -Token $Tokens[$index])) {
        $index++
    }

    return [pscustomobject][ordered]@{
        IsValid = $true
        ErrorMessage = ''
        NextIndex = $index
        IdentifierStart = $identifierStart
        IdentifierEnd = $identifierEnd
    }
}
```

Do not allow a third identifier part: the helper stops after the second part, leaving another `.` for the caller to reject.

- [ ] **Step 5: Add complete-sequence join-prefix recognition**

Place `Get-SqlUtilityJoinPrefix` after the named-source helper. Use a small local top-level word test and return `None`, `Allowed`, or `Unsupported`:

```powershell
function Get-SqlUtilityJoinPrefix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start
    )

    $wordAt = {
        param([int] $Index, [string] $Word)
        return $Index -lt $Tokens.Count -and $Tokens[$Index].Depth -eq 0 -and
            $Tokens[$Index].Kind -eq 'Word' -and $Tokens[$Index].Upper -eq $Word
    }

    if (& $wordAt $Start 'JOIN') {
        return [pscustomobject]@{ Kind = 'Allowed'; Length = 1; ErrorMessage = '' }
    }
    if ((& $wordAt $Start 'INNER') -and (& $wordAt ($Start + 1) 'JOIN')) {
        return [pscustomobject]@{ Kind = 'Allowed'; Length = 2; ErrorMessage = '' }
    }
    if (& $wordAt $Start 'LEFT') {
        if (& $wordAt ($Start + 1) 'JOIN') {
            return [pscustomobject]@{ Kind = 'Allowed'; Length = 2; ErrorMessage = '' }
        }
        if ((& $wordAt ($Start + 1) 'OUTER') -and (& $wordAt ($Start + 2) 'JOIN')) {
            return [pscustomobject]@{ Kind = 'Allowed'; Length = 3; ErrorMessage = '' }
        }
    }

    foreach ($unsupported in @('RIGHT', 'FULL')) {
        if (& $wordAt $Start $unsupported) {
            $joinOffset = if (& $wordAt ($Start + 1) 'OUTER') { 2 } else { 1 }
            if (& $wordAt ($Start + $joinOffset) 'JOIN') {
                return [pscustomobject]@{
                    Kind = 'Unsupported'
                    Length = $joinOffset + 1
                    ErrorMessage = "$unsupported JOIN is not allowed. Use INNER JOIN or LEFT JOIN."
                }
            }
        }
    }
    if ((& $wordAt $Start 'CROSS') -and (& $wordAt ($Start + 1) 'JOIN')) {
        return [pscustomobject]@{
            Kind = 'Unsupported'
            Length = 2
            ErrorMessage = 'CROSS JOIN is not allowed. Use INNER JOIN or LEFT JOIN with ON.'
        }
    }

    return [pscustomobject]@{ Kind = 'None'; Length = 0; ErrorMessage = '' }
}
```

Do not treat incomplete sequences such as `LEFT(value, 2)`, `INNER value`, or `LEFT OUTER value` as join prefixes.

- [ ] **Step 6: Replace the single-source block with primary-source plus join-chain parsing**

In `Test-SqlUtilityQuery`:

1. Remove only `'JOIN'` from `$denied`; keep `'APPLY'` and every other denied keyword.
2. Replace the current source-parsing block from `$sourceIndex = $fromIndexes[0] + 1` through the old single-source boundary check with the following control flow.
3. Keep `$clauses` parsing and all result construction after the new block.

```powershell
$sourceIndex = $fromIndexes[0] + 1
$primarySource = Read-SqlUtilityNamedSource -Tokens $tokens -Start $sourceIndex
if (-not $primarySource.IsValid) {
    return New-SqlUtilityInvalidQueryResult -Message $primarySource.ErrorMessage
}

$tableStart = $primarySource.IdentifierStart
$tableEnd = $primarySource.IdentifierEnd
$sourceIndex = $primarySource.NextIndex
$clauseWords = @('WHERE', 'GROUP', 'HAVING', 'ORDER')

while ($sourceIndex -lt $tokens.Count) {
    $joinPrefix = Get-SqlUtilityJoinPrefix -Tokens $tokens -Start $sourceIndex
    if ($joinPrefix.Kind -eq 'Unsupported') {
        return New-SqlUtilityInvalidQueryResult -Message $joinPrefix.ErrorMessage
    }
    if ($joinPrefix.Kind -ne 'Allowed') {
        break
    }

    $joinedSourceStart = $sourceIndex + $joinPrefix.Length
    $joinedSource = Read-SqlUtilityNamedSource -Tokens $tokens -Start $joinedSourceStart
    if (-not $joinedSource.IsValid) {
        return New-SqlUtilityInvalidQueryResult -Message $joinedSource.ErrorMessage
    }

    $onIndex = $joinedSource.NextIndex
    if ($onIndex -ge $tokens.Count -or $tokens[$onIndex].Depth -ne 0 -or
        $tokens[$onIndex].Kind -ne 'Word' -or $tokens[$onIndex].Upper -ne 'ON') {
        return New-SqlUtilityInvalidQueryResult -Message 'Each JOIN must be followed by an ON predicate.'
    }

    $predicateStart = $onIndex + 1
    $predicateEnd = $predicateStart
    while ($predicateEnd -lt $tokens.Count) {
        $boundaryToken = $tokens[$predicateEnd]
        if ($boundaryToken.Depth -eq 0) {
            $nextJoin = Get-SqlUtilityJoinPrefix -Tokens $tokens -Start $predicateEnd
            if ($nextJoin.Kind -ne 'None') {
                break
            }
            if ($boundaryToken.Kind -eq 'Word' -and $clauseWords -contains $boundaryToken.Upper) {
                break
            }
        }
        $predicateEnd++
    }

    if ($predicateStart -ge $predicateEnd -or
        -not (Test-SqlUtilityExpressionList -Tokens $tokens -Start $predicateStart -End $predicateEnd)) {
        return New-SqlUtilityInvalidQueryResult -Message 'The JOIN ON clause contains an invalid or empty predicate.'
    }

    $sourceIndex = $predicateEnd
}

$atClauseBoundary = $sourceIndex -ge $tokens.Count -or (
    $tokens[$sourceIndex].Depth -eq 0 -and
    $tokens[$sourceIndex].Kind -eq 'Word' -and
    $clauseWords -contains $tokens[$sourceIndex].Upper
)
if (-not $atClauseBoundary) {
    return New-SqlUtilityInvalidQueryResult -Message 'Unexpected tokens follow the table source or JOIN chain.'
}
```

Before finalizing, confirm that the actual `Test-SqlUtilityExpressionList` defaults keep commas and ordering disabled. Pass explicit `-AllowComma $false -AllowOrdering $false` only if PowerShell parameter binding requires it in the current signature.

- [ ] **Step 7: Run the focused suite and make the accepted grammar green**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
```

Expected: all accepted JOIN cases and all pre-existing tests pass except the old rejected case named `JOIN`, which must now be removed from `$rejected`. Remove only that obsolete rejection and rerun until the focused suite reports zero failures.

- [ ] **Step 8: Review and commit the accepted grammar slice**

Run:

```powershell
git diff --check
git diff -- modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git status --short
```

Stage only the policy and its focused test, inspect the staged diff, and commit:

```powershell
git add -- modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git diff --cached --check
git diff --cached --stat
git commit -m "feat: support named-table inner and left joins"
```

Expected: one scoped commit containing accepted JOIN parsing and regression tests; documentation remains unchanged.

---

### Task 2: Harden JOIN boundaries against unsupported and bypass forms

**Files:**
- Modify: `tests/Test-QueryPolicy.ps1:96-245`
- Modify: `modules/SqlUtility.QueryPolicy.ps1` only if a new rejection demonstrates a real parser bypass or unclear structural result

**Interfaces:**
- Consumes: the `Read-SqlUtilityNamedSource`, `Get-SqlUtilityJoinPrefix`, and join-chain behavior completed in Task 1.
- Produces: the final approved JOIN grammar, with invalid results always returning blank `NormalizedSql`, blank `TableIdentifier`, `HasOrderBy = $false`, and blank `CountSourceSql`.

- [ ] **Step 1: Add unsupported JOIN and APPLY cases**

Add these entries to `$rejected`:

```powershell
    @{ Name = 'RIGHT JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i RIGHT JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'RIGHT OUTER JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i RIGHT OUTER JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'FULL JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i FULL JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'FULL OUTER JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i FULL OUTER JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'CROSS JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i CROSS JOIN dbo.Other o' },
    @{ Name = 'CROSS APPLY'; Sql = 'SELECT i.Id FROM dbo.Items i CROSS APPLY dbo.Func(i.Id) f' },
    @{ Name = 'OUTER APPLY'; Sql = 'SELECT i.Id FROM dbo.Items i OUTER APPLY dbo.Func(i.Id) f' },
    @{ Name = 'comma after join'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id, dbo.Third t' }
```

- [ ] **Step 2: Add malformed join-unit cases**

Add:

```powershell
    @{ Name = 'JOIN missing source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN ON i.Id = 1' },
    @{ Name = 'JOIN missing ON'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o WHERE i.Id = 1' },
    @{ Name = 'JOIN empty ON at end'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON' },
    @{ Name = 'JOIN empty ON before WHERE'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON WHERE i.Id = 1' },
    @{ Name = 'JOIN empty ON before chained join'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON LEFT JOIN dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'joined AS missing alias'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other AS ON i.Id = 1' },
    @{ Name = 'joined malformed two-part name'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo. o ON o.Id = i.Id' },
    @{ Name = 'joined extra source token'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o extra ON o.Id = i.Id' },
    @{ Name = 'INNER missing JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i INNER dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'LEFT OUTER missing JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i LEFT OUTER dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'malformed chained INNER boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id INNER dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'malformed chained LEFT OUTER boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id LEFT OUTER dbo.Third t ON t.Id = i.Id' }
```

- [ ] **Step 3: Add excluded-source cases at joined positions**

Add:

```powershell
    @{ Name = 'joined derived source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN (dbo.Other) o ON o.Id = i.Id' },
    @{ Name = 'joined subquery source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN (SELECT Id FROM dbo.Other) o ON o.Id = i.Id' },
    @{ Name = 'subquery inside ON'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id IN (SELECT Id FROM dbo.Third)' },
    @{ Name = 'joined table function'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.GetOther() o ON o.Id = i.Id' },
    @{ Name = 'joined table hint'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o WITH (NOLOCK) ON o.Id = i.Id' },
    @{ Name = 'joined temporary table'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN #Other o ON o.Id = i.Id' },
    @{ Name = 'joined table variable'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN @Other o ON o.Id = i.Id' },
    @{ Name = 'joined three-part source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN UtilityDb.dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'joined four-part source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN ServerA.UtilityDb.dbo.Other o ON o.Id = i.Id' }
```

- [ ] **Step 4: Add statement-boundary and token-confusion cases**

Add:

```powershell
    @{ Name = 'second statement after join'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id; DELETE FROM dbo.Items' },
    @{ Name = 'semicolonless batch after join'; Sql = "SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id`r`nWAITFOR DELAY '00:00:01'" },
    @{ Name = 'JOIN token in SELECT list'; Sql = 'SELECT JOIN FROM dbo.Items' },
    @{ Name = 'JOIN token after WHERE'; Sql = 'SELECT i.Id FROM dbo.Items i WHERE i.Id = 1 JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'unclosed ON parenthesis'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON (o.Id = i.Id' },
    @{ Name = 'unclosed comment at join boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN /* open dbo.Other o ON o.Id = i.Id' }
```

Also add accepted non-executable-token cases to `$accepted` so hardening cannot become overbroad:

```powershell
    @{ Name = 'join keywords in joined query text'; Sql = "SELECT 'RIGHT JOIN FULL JOIN CROSS JOIN' AS Label FROM dbo.Items i JOIN dbo.Other o ON o.Note = 'LEFT JOIN' /* OUTER APPLY */"; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT 'RIGHT JOIN FULL JOIN CROSS JOIN' AS Label FROM dbo.Items i JOIN dbo.Other o ON o.Note = 'LEFT JOIN' /* OUTER APPLY */" }
    @{ Name = 'delimited join keywords as identifiers'; Sql = 'SELECT [JOIN].[LEFT] FROM dbo.Items [JOIN] INNER JOIN dbo.Other [RIGHT] ON [RIGHT].[Id] = [JOIN].[Id]'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT [JOIN].[LEFT] FROM dbo.Items [JOIN] INNER JOIN dbo.Other [RIGHT] ON [RIGHT].[Id] = [JOIN].[Id]' }
```

- [ ] **Step 5: Run the focused suite and classify every failure before editing production code**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
```

Expected: most rejected cases already pass. Any accepted bypass is a production defect; any rejected safe case is an overbroad boundary defect. Do not weaken an existing denied-token or single-statement check to make a case pass.

- [ ] **Step 6: Apply only evidence-driven parser corrections**

For each failing case:

- Keep unsupported complete join prefixes returning `New-SqlUtilityInvalidQueryResult` before SQL execution.
- Keep `APPLY` in the global denied list.
- Keep nested/additional `SELECT` and all statement starters denied wherever they occur as executable word tokens.
- Keep the named-source helper limited to two identifier parts.
- Make join-boundary recognition conditional on complete depth-zero word sequences.
- Make an empty `ON` range fail before expression validation.
- Do not special-case quoted identifiers, strings, or comments; rely on their token kinds.

If no new production correction is needed, commit the test hardening alone. If a correction is needed, add the smallest branch or boundary check and rerun the focused suite after each change.

- [ ] **Step 7: Verify the invalid-result safety contract explicitly**

The existing loop at the end of `Test-QueryPolicy.ps1` must continue asserting for every rejected case:

```powershell
Assert-False $result.IsValid "$($case.Name) is rejected"
Assert-Equal '' $result.NormalizedSql "$($case.Name) has no executable SQL"
Assert-Equal '' $result.TableIdentifier "$($case.Name) has no executable table"
Assert-Equal $false $result.HasOrderBy "$($case.Name) cannot be treated as ordered"
Assert-Equal '' $result.CountSourceSql "$($case.Name) has no count source"
```

Do not add a second rejection loop with weaker assertions.

- [ ] **Step 8: Run focused and aggregate suites**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
```

Expected: both commands exit `0` with all suites passing. A failure outside QueryPolicy must be investigated as a contract regression; do not update unrelated expectations without evidence.

- [ ] **Step 9: Review and commit the bypass-hardening slice**

Run:

```powershell
git diff --check
git diff -- modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git status --short
```

Stage only files changed by this task, inspect the staged diff, and commit:

```powershell
git add -- modules/SqlUtility.QueryPolicy.ps1 tests/Test-QueryPolicy.ps1
git diff --cached --check
git diff --cached --stat
git commit -m "test: harden join query policy boundaries"
```

If the production module did not change in this task, stage and commit only `tests/Test-QueryPolicy.ps1`.

---

### Task 3: Synchronize user and engineering documentation

**Files:**
- Modify: `README.md:3-18,76,103-106,157-206,295-350`
- Modify: `AGENTS.md:3-66`

**Interfaces:**
- Consumes: the final verified grammar and behavior from Tasks 1 and 2.
- Produces: current user-facing operating documentation and repository-wide engineering invariants that match the implemented feature.

- [ ] **Step 1: Update README overview, features, architecture, and workflow wording**

Make these exact contract changes without rewriting unrelated sections:

- Replace "one named table"/"single-table" wording with a read-only `SELECT` over one primary named source plus optional chained named-table `INNER`/`LEFT` joins.
- Keep the grid read-only; "read-only" describes query behavior, not the number of sources.
- Update the QueryPolicy architecture row to say it validates named sources, approved join chains, normalization, primary-table extraction, and top-level ordering.
- Do not change connection, paging, count, export, or packaging instructions.

Use this concise feature bullet:

```markdown
- One read-only `SELECT` statement over a primary named table with optional chained named-table `INNER JOIN` and `LEFT JOIN` clauses.
```

- [ ] **Step 2: Replace the README Query Policy grammar and examples**

Document the logical shape as:

```sql
SELECT [DISTINCT] expressions
FROM [schema.]table [alias]
{
    [INNER] JOIN [schema.]table [alias] ON predicate
  | LEFT [OUTER] JOIN [schema.]table [alias] ON predicate
} [...]
[WHERE predicate]
[GROUP BY expressions]
[HAVING predicate]
[ORDER BY expressions]
```

Add one simple inner-join example and one mixed chained-join example from the approved design. State explicitly:

- Bare `JOIN` means `INNER JOIN`.
- Every join requires `ON`.
- `ON` accepts the existing expression/predicate grammar.
- `TableIdentifier` remains an internal primary-source result and is not a user-facing multi-source catalog.
- Result limits do not bound SQL Server join work or intermediate results.

Replace the blanket `JOIN` rejection bullet with the exact still-rejected forms from the design. Keep the least-privilege warning intact.

- [ ] **Step 3: Update README testing, acceptance, limitations, and future-extension sections**

- Update the `Test-QueryPolicy.ps1` coverage row to include accepted JOIN chains and bypass-focused source/join rejection.
- Add representative inner/left/mixed joins to the Citrix acceptance checklist, including ordered paging, explicit count, complete export, one unsupported-join rejection, and one server-side semantic error.
- State that JOIN execution in real Citrix/SQL Server remains pending until actually performed.
- Move `INNER JOIN`/`LEFT JOIN` out of Future Extensions. Keep `RIGHT`/`FULL`/`CROSS JOIN`, `APPLY`, derived sources, subqueries, CTEs, set operators, table functions, cross-database sources, background execution, and guarded UPDATE deferred.

- [ ] **Step 4: Update AGENTS.md architecture and behavioral invariants**

Replace the QueryPolicy ownership line with:

```markdown
- `modules/SqlUtility.QueryPolicy.ps1`: SQL tokenization, named-source read-only grammar, approved `INNER JOIN`/`LEFT JOIN` chain validation, normalization, primary-table extraction, and top-level ordering detection.
```

Replace the single-table invariant with:

```markdown
- The query policy remains one statement and read-only `SELECT`. It permits one primary one- or two-part named source plus zero or more chained bare/`INNER JOIN`, `LEFT JOIN`, or `LEFT OUTER JOIN` units, each with one named source and a required valid `ON` predicate. It continues to exclude all other join/apply/source forms, subqueries, CTEs, set operators, batches, stored/dynamic SQL, external sources, and data-changing, DDL, transaction, permission, or administrative commands.
```

Replace the now-completed JOIN-specific widening sentence with a future-facing rule:

```markdown
Any further query-grammar widening requires explicit user approval, an updated grammar design, bypass-focused regression tests, and preservation of every unaffected rejected case.
```

- [ ] **Step 5: Run documentation consistency searches**

Run:

```powershell
rg -n "one named table|single-table|rejects:|JOIN|Future Extensions|query policy remains" README.md AGENTS.md
```

Expected: historical single-table statements remain only in historical spec/plan files, not as current behavior in README or AGENTS. Every current JOIN statement must match the approved types and named-source restrictions.

- [ ] **Step 6: Run fresh full verification after documentation changes**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-QueryPolicy.ps1
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1
git diff --check
git status --short
```

Expected: focused and aggregate tests exit `0`; diff check emits no errors; status lists only the intentional README and AGENTS changes for this task.

- [ ] **Step 7: Review and commit documentation**

Run:

```powershell
git diff -- README.md AGENTS.md
git status --short
```

Stage only the documentation, inspect the staged diff, and commit:

```powershell
git add -- README.md AGENTS.md
git diff --cached --check
git diff --cached --stat
git commit -m "docs: document read-only join support"
```

Expected: one documentation-only commit with no runtime distribution changes.

---

## Final Review and Handoff

- [ ] Confirm `git status --short` is clean.
- [ ] Inspect `git log --oneline main..HEAD` and `git diff --stat main...HEAD` for only the approved design, plan, policy, tests, README, and AGENTS changes.
- [ ] Rerun `powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-All.ps1` immediately before claiming completion.
- [ ] Rerun `git diff --check` and record the output.
- [ ] Report the exact focused and aggregate pass totals from fresh output.
- [ ] Report Citrix/live SQL Server/desktop Excel JOIN acceptance as pending unless it was actually performed.
- [ ] Do not merge, push, delete the branch, or remove any worktree. Ask the user for the integration choice after review.

## External Acceptance Checklist

Use a Windows identity with least-privileged access and non-sensitive test data. In Citrix:

1. Launch the unchanged six-file portable distribution through `StartSqlUtility.cmd`.
2. Execute a named-table `INNER JOIN` and verify returned columns/rows.
3. Execute a mixed chained `LEFT OUTER JOIN` plus `INNER JOIN`.
4. Page an ordered joined result past the first 500 rows and back.
5. Run explicit Count for an ordered joined result.
6. Trigger the unordered row limit with a joined result and verify truncation/export behavior.
7. Export a complete joined result and open the workbook in desktop Excel, including duplicate source column names.
8. Verify an unsupported `RIGHT JOIN` is rejected before a database call.
9. Verify a valid-shape query with an unknown or ambiguous column reaches SQL Server and reports the existing Query Error path.
10. Restart and confirm no query text, results, or export destination was persisted.
