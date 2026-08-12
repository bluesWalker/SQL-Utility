$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.QueryPolicy.ps1')

$accepted = @(
    @{ Name = 'simple select'; Sql = 'SELECT * FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT * FROM dbo.Items' },
    @{ Name = 'alias and ordered select'; Sql = 'SELECT i.Id FROM dbo.Items AS i WHERE i.Enabled = 1 ORDER BY i.Id'; Table = 'dbo.Items'; Ordered = $true; Normalized = 'SELECT i.Id FROM dbo.Items AS i WHERE i.Enabled = 1 ORDER BY i.Id' },
    @{ Name = 'grouped bracketed select'; Sql = 'SELECT DISTINCT [Type], COUNT(*) AS [Count] FROM [dbo].[Items] GROUP BY [Type] HAVING COUNT(*) > 1 ORDER BY [Type];'; Table = '[dbo].[Items]'; Ordered = $true; Normalized = 'SELECT DISTINCT [Type], COUNT(*) AS [Count] FROM [dbo].[Items] GROUP BY [Type] HAVING COUNT(*) > 1 ORDER BY [Type]' },
    @{ Name = 'keywords in comments and strings'; Sql = "-- SELECT FROM fake`r`nSELECT CASE WHEN Name = 'ORDER BY' THEN 1 ELSE 0 END AS Flag FROM dbo.Items WHERE Note = 'JOIN'"; Table = 'dbo.Items'; Ordered = $false; Normalized = "-- SELECT FROM fake`r`nSELECT CASE WHEN Name = 'ORDER BY' THEN 1 ELSE 0 END AS Flag FROM dbo.Items WHERE Note = 'JOIN'" },
    @{ Name = 'keywords in quoted text'; Sql = 'SELECT "ORDER BY" AS [Value] FROM dbo.Items /* JOIN x */ ORDER BY Id'; Table = 'dbo.Items'; Ordered = $true; Normalized = 'SELECT "ORDER BY" AS [Value] FROM dbo.Items /* JOIN x */ ORDER BY Id' },
    @{ Name = 'alias without AS'; Sql = 'SELECT i.Id FROM dbo.Items i WHERE i.Enabled = 1'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT i.Id FROM dbo.Items i WHERE i.Enabled = 1' },
    @{ Name = 'quoted alias without AS'; Sql = 'SELECT [item alias].Id FROM dbo.Items [item alias]'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT [item alias].Id FROM dbo.Items [item alias]' },
    @{ Name = 'quoted table and alias'; Sql = 'SELECT i.Id FROM "dbo"."Items" AS "i"'; Table = '"dbo"."Items"'; Ordered = $false; Normalized = 'SELECT i.Id FROM "dbo"."Items" AS "i"' },
    @{ Name = 'reserved words in quoted identifiers'; Sql = 'SELECT [JOIN], "TOP" FROM [dbo].[UNION]'; Table = '[dbo].[UNION]'; Ordered = $false; Normalized = 'SELECT [JOIN], "TOP" FROM [dbo].[UNION]' },
    @{ Name = 'parenthesized scalar expressions'; Sql = 'SELECT COALESCE((Id + 1), 0) AS NextId FROM dbo.Items WHERE (Enabled = 1)'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT COALESCE((Id + 1), 0) AS NextId FROM dbo.Items WHERE (Enabled = 1)' },
    @{ Name = 'nested order by does not order result'; Sql = 'SELECT ROW_NUMBER() OVER (ORDER BY Id) AS RowNumber FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT ROW_NUMBER() OVER (ORDER BY Id) AS RowNumber FROM dbo.Items' },
    @{ Name = 'compound scalar expression'; Sql = "SELECT CASE WHEN i.Score >= 10 AND i.Name IS NOT NULL THEN COALESCE(i.Score, 0) ELSE -1 END AS Score FROM dbo.Items AS i WHERE i.Id <> 0 AND (i.Enabled = 1 OR i.Name LIKE 'A%') ORDER BY i.Id DESC"; Table = 'dbo.Items'; Ordered = $true; Normalized = "SELECT CASE WHEN i.Score >= 10 AND i.Name IS NOT NULL THEN COALESCE(i.Score, 0) ELSE -1 END AS Score FROM dbo.Items AS i WHERE i.Id <> 0 AND (i.Enabled = 1 OR i.Name LIKE 'A%') ORDER BY i.Id DESC" },
    @{ Name = 'qualified wildcard decimal and negated operators'; Sql = "SELECT i.*, i.Price * 1.25 AS AdjustedPrice FROM dbo.Items AS i WHERE i.Id NOT IN (1, 2) AND i.Name NOT LIKE 'X%' ORDER BY i.Price DESC"; Table = 'dbo.Items'; Ordered = $true; Normalized = "SELECT i.*, i.Price * 1.25 AS AdjustedPrice FROM dbo.Items AS i WHERE i.Id NOT IN (1, 2) AND i.Name NOT LIKE 'X%' ORDER BY i.Price DESC" },
    @{ Name = 'Unicode string literal'; Sql = "SELECT N'unicode' AS [Value] FROM dbo.Items"; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT N'unicode' AS [Value] FROM dbo.Items" },
    @{ Name = 'leading decimal literal'; Sql = 'SELECT .5 AS Fraction FROM dbo.Items WHERE Ratio >= .5'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT .5 AS Fraction FROM dbo.Items WHERE Ratio >= .5' },
    @{ Name = 'statement starters in non-executable text'; Sql = "SELECT 'WAITFOR KILL SHUTDOWN RECONFIGURE CHECKPOINT PRINT RAISERROR THROW RETURN GOTO' AS [PRINT], [KILL] FROM dbo.Items /* BREAK CONTINUE IF WHILE OPEN CLOSE DEALLOCATE SAVE REVERT */"; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT 'WAITFOR KILL SHUTDOWN RECONFIGURE CHECKPOINT PRINT RAISERROR THROW RETURN GOTO' AS [PRINT], [KILL] FROM dbo.Items /* BREAK CONTINUE IF WHILE OPEN CLOSE DEALLOCATE SAVE REVERT */" },
    @{ Name = 'final semicolon before line comment'; Sql = 'SELECT Id FROM dbo.Items ORDER BY Id; -- trailing comment'; Table = 'dbo.Items'; Ordered = $true; Normalized = 'SELECT Id FROM dbo.Items ORDER BY Id -- trailing comment' },
    @{ Name = 'final semicolon before block comment'; Sql = "SELECT Id FROM dbo.Items ;`r`n/* trailing comment ; */  "; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT Id FROM dbo.Items `r`n/* trailing comment ; */  " },
    @{ Name = 'escaped delimiters'; Sql = 'SELECT ''it''''s'', [a]]b], "a""b" FROM [dbo].[Items]'; Table = '[dbo].[Items]'; Ordered = $false; Normalized = 'SELECT ''it''''s'', [a]]b], "a""b" FROM [dbo].[Items]' }
    @{ Name = 'balanced nested block comments'; Sql = 'SELECT /* outer /* inner */ outer */ Id FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT /* outer /* inner */ outer */ Id FROM dbo.Items' }
    @{ Name = 'ordinary implicit select alias'; Sql = 'SELECT Id ItemId FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT Id ItemId FROM dbo.Items' }
    @{ Name = 'delimited implicit select alias'; Sql = 'SELECT Id [SHUTDOWN] FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT Id [SHUTDOWN] FROM dbo.Items' }
    @{ Name = 'delimited reserved table alias'; Sql = 'SELECT [SHUTDOWN].Id FROM dbo.Items [SHUTDOWN]'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT [SHUTDOWN].Id FROM dbo.Items [SHUTDOWN]' }
    @{ Name = 'CAST scalar type'; Sql = 'SELECT CAST(Id AS bigint) AS ConvertedId FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT CAST(Id AS bigint) AS ConvertedId FROM dbo.Items' }
    @{ Name = 'CAST length type and implicit alias'; Sql = 'SELECT CAST(Name AS nvarchar(50)) ConvertedName FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT CAST(Name AS nvarchar(50)) ConvertedName FROM dbo.Items' }
    @{ Name = 'CAST precision and scale type'; Sql = 'SELECT CAST(Amount AS decimal(18, 2)) AS ConvertedAmount FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT CAST(Amount AS decimal(18, 2)) AS ConvertedAmount FROM dbo.Items' }
    @{ Name = 'TRY_CAST maximum length type'; Sql = 'SELECT TRY_CAST(Payload AS varbinary(max)) AS ConvertedPayload FROM dbo.Items'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT TRY_CAST(Payload AS varbinary(max)) AS ConvertedPayload FROM dbo.Items' }
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
    @{ Name = 'join keywords in joined query text'; Sql = "SELECT 'RIGHT JOIN FULL JOIN CROSS JOIN' AS Label FROM dbo.Items i JOIN dbo.Other o ON o.Note = 'LEFT JOIN' /* OUTER APPLY */"; Table = 'dbo.Items'; Ordered = $false; Normalized = "SELECT 'RIGHT JOIN FULL JOIN CROSS JOIN' AS Label FROM dbo.Items i JOIN dbo.Other o ON o.Note = 'LEFT JOIN' /* OUTER APPLY */" }
    @{ Name = 'delimited join keywords as identifiers'; Sql = 'SELECT [JOIN].[LEFT] FROM dbo.Items [JOIN] INNER JOIN dbo.Other [RIGHT] ON [RIGHT].[Id] = [JOIN].[Id]'; Table = 'dbo.Items'; Ordered = $false; Normalized = 'SELECT [JOIN].[LEFT] FROM dbo.Items [JOIN] INNER JOIN dbo.Other [RIGHT] ON [RIGHT].[Id] = [JOIN].[Id]' }
)

foreach ($case in $accepted) {
    $result = Test-SqlUtilityQuery -Sql $case.Sql
    Assert-True $result.IsValid "$($case.Name) is accepted"
    Assert-True ($result.IsValid -is [bool]) "$($case.Name) returns Boolean IsValid"
    Assert-Equal '' $result.ErrorMessage "$($case.Name) has no error"
    Assert-Equal $case.Normalized $result.NormalizedSql "$($case.Name) preserves and normalizes source"
    Assert-Equal $case.Table $result.TableIdentifier "$($case.Name) returns the exact table identifier"
    Assert-Equal $case.Ordered $result.HasOrderBy "$($case.Name) detects only depth-zero ORDER BY"
    Assert-True ($result.HasOrderBy -is [bool]) "$($case.Name) returns Boolean HasOrderBy"
    Assert-True ($null -ne $result.PSObject.Properties['CountSourceSql']) "$($case.Name) returns a count source"
}

$unordered = Test-SqlUtilityQuery -Sql 'SELECT * FROM dbo.Items;'
Assert-Equal 'SELECT * FROM dbo.Items' $unordered.CountSourceSql `
    'Unordered count source keeps the complete normalized query'

$ordered = Test-SqlUtilityQuery -Sql @'
SELECT CategoryId, COUNT(*)
FROM dbo.Items
GROUP BY CategoryId
HAVING COUNT(*) > 5
ORDER BY CategoryId;
'@
Assert-Equal @'
SELECT CategoryId, COUNT(*)
FROM dbo.Items
GROUP BY CategoryId
HAVING COUNT(*) > 5
'@ $ordered.CountSourceSql 'Ordered count source removes only top-level ORDER BY'

$orderInText = Test-SqlUtilityQuery -Sql "SELECT 'ORDER BY' AS Label FROM dbo.Items ORDER BY Id"
Assert-Equal "SELECT 'ORDER BY' AS Label FROM dbo.Items" $orderInText.CountSourceSql `
    'ORDER BY text inside a string literal is not the truncation point'

$orderInExpression = Test-SqlUtilityQuery -Sql 'SELECT ROW_NUMBER() OVER (ORDER BY Id) AS RowNumber FROM dbo.Items ORDER BY RowNumber'
Assert-Equal 'SELECT ROW_NUMBER() OVER (ORDER BY Id) AS RowNumber FROM dbo.Items' $orderInExpression.CountSourceSql `
    'ORDER BY inside an expression is not the truncation point'

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

$countSql = New-SqlUtilityCountSql `
    -CountSourceSql 'SELECT *, Price * 1.25 FROM dbo.Items' `
    -OutputColumnCount 4

$expectedCountSql = @'
SELECT COUNT_BIG(*)
FROM (
SELECT *, Price * 1.25 FROM dbo.Items
) AS [SqlUtilityCountSource] ([SqlUtilityCountColumn1], [SqlUtilityCountColumn2], [SqlUtilityCountColumn3], [SqlUtilityCountColumn4]);
'@ -replace "`r?`n", "`r`n"
Assert-Equal $expectedCountSql $countSql 'Count wrapper supplies one alias per actual result column'

Assert-Throws {
    New-SqlUtilityCountSql -CountSourceSql 'SELECT Id FROM dbo.Items' -OutputColumnCount 0
} 'System.ArgumentOutOfRangeException' 'Count builder rejects a zero-column schema'

Assert-Throws {
    New-SqlUtilityCountSql -CountSourceSql '   ' -OutputColumnCount 1
} 'System.ArgumentException' 'Count builder rejects an empty source'

$rejected = @(
    @{ Name = 'blank SQL'; Sql = '   ' },
    @{ Name = 'comment-only SQL'; Sql = '-- nothing executable' },
    @{ Name = 'two statements'; Sql = 'SELECT * FROM dbo.Items; SELECT * FROM dbo.Other' },
    @{ Name = 'two final semicolons'; Sql = 'SELECT * FROM dbo.Items;;' },
    @{ Name = 'CTE'; Sql = 'WITH cte AS (SELECT * FROM dbo.Items) SELECT * FROM cte' },
    @{ Name = 'nested SELECT'; Sql = 'SELECT Id FROM dbo.Items WHERE Id IN (SELECT Id FROM dbo.Other)' },
    @{ Name = 'APPLY'; Sql = 'SELECT i.Id FROM dbo.Items i CROSS APPLY dbo.Func(i.Id) f' },
    @{ Name = 'comma table source'; Sql = 'SELECT * FROM dbo.Items, dbo.Other' },
    @{ Name = 'UNION'; Sql = 'SELECT Id FROM dbo.Items UNION SELECT Id FROM dbo.Other' },
    @{ Name = 'INTERSECT'; Sql = 'SELECT Id FROM dbo.Items INTERSECT SELECT Id FROM dbo.Other' },
    @{ Name = 'EXCEPT'; Sql = 'SELECT Id FROM dbo.Items EXCEPT SELECT Id FROM dbo.Other' },
    @{ Name = 'INTO'; Sql = 'SELECT * INTO dbo.Copy FROM dbo.Items' },
    @{ Name = 'TOP'; Sql = 'SELECT TOP 10 * FROM dbo.Items' },
    @{ Name = 'OFFSET'; Sql = 'SELECT * FROM dbo.Items ORDER BY Id OFFSET 10 ROWS' },
    @{ Name = 'FETCH'; Sql = 'SELECT * FROM dbo.Items ORDER BY Id FETCH NEXT 10 ROWS ONLY' },
    @{ Name = 'EXEC'; Sql = 'EXEC dbo.DoWork' },
    @{ Name = 'EXECUTE'; Sql = 'EXECUTE dbo.DoWork' },
    @{ Name = 'INSERT'; Sql = 'INSERT INTO dbo.Items(Id) VALUES (1)' },
    @{ Name = 'UPDATE'; Sql = 'UPDATE dbo.Items SET Enabled = 0' },
    @{ Name = 'DELETE'; Sql = 'DELETE FROM dbo.Items' },
    @{ Name = 'MERGE'; Sql = 'MERGE dbo.Items AS target USING dbo.Other AS source ON target.Id = source.Id WHEN MATCHED THEN UPDATE SET target.Id = source.Id;' },
    @{ Name = 'DROP'; Sql = 'DROP TABLE dbo.Items' },
    @{ Name = 'ALTER'; Sql = 'ALTER TABLE dbo.Items ADD Flag bit' },
    @{ Name = 'CREATE'; Sql = 'CREATE TABLE dbo.Items(Id int)' },
    @{ Name = 'TRUNCATE'; Sql = 'TRUNCATE TABLE dbo.Items' },
    @{ Name = 'BEGIN transaction'; Sql = 'BEGIN TRANSACTION' },
    @{ Name = 'COMMIT'; Sql = 'COMMIT TRANSACTION' },
    @{ Name = 'ROLLBACK'; Sql = 'ROLLBACK TRANSACTION' },
    @{ Name = 'GRANT'; Sql = 'GRANT SELECT ON dbo.Items TO Public' },
    @{ Name = 'REVOKE'; Sql = 'REVOKE SELECT ON dbo.Items FROM Public' },
    @{ Name = 'DENY'; Sql = 'DENY SELECT ON dbo.Items TO Public' },
    @{ Name = 'DECLARE'; Sql = 'DECLARE @Id int' },
    @{ Name = 'SET'; Sql = 'SET NOCOUNT ON' },
    @{ Name = 'USE'; Sql = 'USE master' },
    @{ Name = 'BACKUP'; Sql = 'BACKUP DATABASE UtilityDb TO DISK = ''x.bak''' },
    @{ Name = 'RESTORE'; Sql = 'RESTORE DATABASE UtilityDb FROM DISK = ''x.bak''' },
    @{ Name = 'DBCC'; Sql = 'DBCC CHECKDB' },
    @{ Name = 'BULK'; Sql = 'BULK INSERT dbo.Items FROM ''items.csv''' },
    @{ Name = 'temporary table'; Sql = 'SELECT * FROM #Items' },
    @{ Name = 'table variable'; Sql = 'SELECT * FROM @Items' },
    @{ Name = 'three-part source'; Sql = 'SELECT * FROM UtilityDb.dbo.Items' },
    @{ Name = 'four-part source'; Sql = 'SELECT * FROM ServerA.UtilityDb.dbo.Items' },
    @{ Name = 'OPENQUERY'; Sql = 'SELECT * FROM OPENQUERY(ServerA, ''SELECT Id FROM dbo.Items'')' },
    @{ Name = 'OPENROWSET'; Sql = 'SELECT * FROM OPENROWSET(''SQLNCLI'', ''Server=ServerA;Trusted_Connection=yes;'', ''SELECT Id FROM dbo.Items'')' },
    @{ Name = 'OPENDATASOURCE'; Sql = 'SELECT * FROM OPENDATASOURCE(''SQLNCLI'', ''Data Source=ServerA;Integrated Security=SSPI'').UtilityDb.dbo.Items' },
    @{ Name = 'unclosed string'; Sql = 'SELECT ''unterminated FROM dbo.Items' },
    @{ Name = 'unclosed quoted identifier'; Sql = 'SELECT "unterminated FROM dbo.Items' },
    @{ Name = 'unclosed bracketed identifier'; Sql = 'SELECT [unterminated FROM dbo.Items' },
    @{ Name = 'unclosed block comment'; Sql = 'SELECT * FROM dbo.Items /* unterminated' },
    @{ Name = 'unclosed parenthesis'; Sql = 'SELECT (Id + 1 FROM dbo.Items' },
    @{ Name = 'unexpected closing parenthesis'; Sql = 'SELECT Id) FROM dbo.Items' },
    @{ Name = 'missing FROM'; Sql = 'SELECT 1' },
    @{ Name = 'multiple depth-zero FROM clauses'; Sql = 'SELECT FROM FROM dbo.Items' },
    @{ Name = 'parenthesized table source'; Sql = 'SELECT * FROM (dbo.Items)' },
    @{ Name = 'table alias missing after AS'; Sql = 'SELECT * FROM dbo.Items AS' },
    @{ Name = 'clause keyword is not an alias after AS'; Sql = 'SELECT * FROM dbo.Items AS WHERE' },
    @{ Name = 'extra table source token'; Sql = 'SELECT * FROM dbo.Items i extra' },
    @{ Name = 'GROUP without BY'; Sql = 'SELECT Id FROM dbo.Items GROUP Id' },
    @{ Name = 'ORDER without BY'; Sql = 'SELECT Id FROM dbo.Items ORDER Id' }
)

$rejected += @(
    @{ Name = 'semicolonless WAITFOR batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nWAITFOR DELAY '00:00:01'" },
    @{ Name = 'semicolonless KILL batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nKILL 51" },
    @{ Name = 'semicolonless SHUTDOWN batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nSHUTDOWN" },
    @{ Name = 'semicolonless RECONFIGURE batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nRECONFIGURE" },
    @{ Name = 'semicolonless CHECKPOINT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nCHECKPOINT" },
    @{ Name = 'semicolonless PRINT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nPRINT 'message'" },
    @{ Name = 'semicolonless RAISERROR batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nRAISERROR ('message', 16, 1)" },
    @{ Name = 'semicolonless THROW batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nTHROW 50000, 'message', 1" },
    @{ Name = 'semicolonless RETURN batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nRETURN" },
    @{ Name = 'semicolonless GOTO batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nGOTO target_label" },
    @{ Name = 'semicolonless BREAK batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nBREAK" },
    @{ Name = 'semicolonless CONTINUE batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nCONTINUE" },
    @{ Name = 'semicolonless IF batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nIF 1 = 1 PRINT 'message'" },
    @{ Name = 'semicolonless WHILE batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nWHILE 1 = 0 PRINT 'message'" },
    @{ Name = 'semicolonless CLOSE cursor batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nCLOSE item_cursor" },
    @{ Name = 'semicolonless DEALLOCATE cursor batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nDEALLOCATE item_cursor" },
    @{ Name = 'semicolonless OPEN cursor batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nOPEN item_cursor" },
    @{ Name = 'semicolonless SAVE transaction batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nSAVE TRANSACTION SavePoint1" },
    @{ Name = 'semicolonless REVERT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nREVERT" },
    @{ Name = 'semicolonless READTEXT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nREADTEXT dbo.Items.Payload @pointer 0 1" },
    @{ Name = 'semicolonless UPDATETEXT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nUPDATETEXT dbo.Items.Payload @pointer 0 1 'x'" },
    @{ Name = 'semicolonless WRITETEXT batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nWRITETEXT dbo.Items.Payload @pointer 'x'" },
    @{ Name = 'semicolonless RECEIVE batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nRECEIVE * FROM dbo.Queue" },
    @{ Name = 'semicolonless SEND batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nSEND ON CONVERSATION @handle MESSAGE TYPE [type] ('x')" },
    @{ Name = 'semicolonless SETUSER batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nSETUSER 'dbo'" },
    @{ Name = 'semicolonless DUMP batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nDUMP DATABASE UtilityDb TO DISK = 'utility.bak'" },
    @{ Name = 'semicolonless LOAD batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nLOAD DATABASE UtilityDb FROM DISK = 'utility.bak'" },
    @{ Name = 'semicolonless DISK batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nDISK INIT NAME = 'utility', PHYSNAME = 'utility.dat', VDEVNO = 1, SIZE = 2" },
    @{ Name = 'semicolonless ADD SIGNATURE batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nADD SIGNATURE TO dbo.ProcedureName BY CERTIFICATE SigningCertificate" },
    @{ Name = 'semicolonless DISABLE TRIGGER batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nDISABLE TRIGGER dbo.TriggerName ON dbo.Items" },
    @{ Name = 'semicolonless ENABLE TRIGGER batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nENABLE TRIGGER dbo.TriggerName ON dbo.Items" },
    @{ Name = 'semicolonless RENAME batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nRENAME OBJECT dbo.Items TO ArchivedItems" },
    @{ Name = 'semicolonless END CONVERSATION batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nEND CONVERSATION @handle" },
    @{ Name = 'semicolonless MOVE CONVERSATION batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nMOVE CONVERSATION @handle TO @group" },
    @{ Name = 'semicolonless GET CONVERSATION GROUP batch'; Sql = "SELECT * FROM dbo.Items WHERE 1 = 1`r`nGET CONVERSATION GROUP @group FROM dbo.Queue" }
)

$rejected += @(
    @{ Name = 'arbitrary trailing words after WHERE'; Sql = "SELECT * FROM dbo.Items WHERE Id = 1`r`nFROBNICATE target" },
    @{ Name = 'arbitrary trailing number after WHERE'; Sql = "SELECT * FROM dbo.Items WHERE Id = 1`r`n987654" },
    @{ Name = 'arbitrary trailing words after GROUP BY'; Sql = "SELECT [Type] FROM dbo.Items GROUP BY [Type]`r`nMYSTERY target" },
    @{ Name = 'arbitrary trailing words after HAVING'; Sql = "SELECT [Type] FROM dbo.Items GROUP BY [Type] HAVING COUNT(*) > 1`r`nXYZZY target" },
    @{ Name = 'arbitrary trailing words after ORDER BY'; Sql = "SELECT Id FROM dbo.Items ORDER BY Id`r`nWARP target" },
    @{ Name = 'out-of-order WHERE clause'; Sql = 'SELECT Id FROM dbo.Items ORDER BY Id WHERE Enabled = 1' },
    @{ Name = 'repeated WHERE clause'; Sql = 'SELECT Id FROM dbo.Items WHERE Enabled = 1 WHERE Id > 0' },
    @{ Name = 'empty WHERE expression'; Sql = 'SELECT Id FROM dbo.Items WHERE' },
    @{ Name = 'empty GROUP BY expression'; Sql = 'SELECT Id FROM dbo.Items GROUP BY' },
    @{ Name = 'empty HAVING expression'; Sql = 'SELECT Id FROM dbo.Items GROUP BY Id HAVING' },
    @{ Name = 'empty ORDER BY expression'; Sql = 'SELECT Id FROM dbo.Items ORDER BY' }
)

$rejected += @(
    @{ Name = 'parenthesis cannot follow completed operand'; Sql = 'SELECT * FROM dbo.Items WHERE Id = 1 (FROBNICATE target)' },
    @{ Name = 'nested group content is validated'; Sql = 'SELECT * FROM dbo.Items WHERE (Id = 1 FROBNICATE target)' },
    @{ Name = 'duplicate equals operators'; Sql = 'SELECT * FROM dbo.Items WHERE Id = = 1' },
    @{ Name = 'spaced symbols do not form compound operator'; Sql = 'SELECT * FROM dbo.Items WHERE Id > = 1' }
)

$rejected += @(
    @{ Name = 'RIGHT JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i RIGHT JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'RIGHT OUTER JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i RIGHT OUTER JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'FULL JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i FULL JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'FULL OUTER JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i FULL OUTER JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'CROSS JOIN'; Sql = 'SELECT i.Id FROM dbo.Items i CROSS JOIN dbo.Other o' },
    @{ Name = 'LOOP JOIN hint after initial source'; Sql = 'SELECT Items.Id FROM dbo.Items LOOP JOIN dbo.Other o ON o.Id = Items.Id' },
    @{ Name = 'HASH JOIN hint after initial source'; Sql = 'SELECT Items.Id FROM dbo.Items HASH JOIN dbo.Other o ON o.Id = Items.Id' },
    @{ Name = 'MERGE JOIN hint after initial source'; Sql = 'SELECT Items.Id FROM dbo.Items MERGE JOIN dbo.Other o ON o.Id = Items.Id' },
    @{ Name = 'REMOTE JOIN hint after initial source'; Sql = 'SELECT Items.Id FROM dbo.Items REMOTE JOIN dbo.Other o ON o.Id = Items.Id' },
    @{ Name = 'LOOP JOIN hint at chained boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id LOOP JOIN dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'HASH JOIN hint at chained boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id HASH JOIN dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'MERGE JOIN hint at chained boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id MERGE JOIN dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'REMOTE JOIN hint at chained boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id REMOTE JOIN dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'CROSS APPLY'; Sql = 'SELECT i.Id FROM dbo.Items i CROSS APPLY dbo.Func(i.Id) f' },
    @{ Name = 'OUTER APPLY'; Sql = 'SELECT i.Id FROM dbo.Items i OUTER APPLY dbo.Func(i.Id) f' },
    @{ Name = 'comma after join'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id, dbo.Third t' },
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
    @{ Name = 'malformed chained LEFT OUTER boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id LEFT OUTER dbo.Third t ON t.Id = i.Id' },
    @{ Name = 'joined derived source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN (dbo.Other) o ON o.Id = i.Id' },
    @{ Name = 'joined subquery source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN (SELECT Id FROM dbo.Other) o ON o.Id = i.Id' },
    @{ Name = 'subquery inside ON'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id IN (SELECT Id FROM dbo.Third)' },
    @{ Name = 'joined table function'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.GetOther() o ON o.Id = i.Id' },
    @{ Name = 'joined table hint'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o WITH (NOLOCK) ON o.Id = i.Id' },
    @{ Name = 'joined temporary table'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN #Other o ON o.Id = i.Id' },
    @{ Name = 'joined table variable'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN @Other o ON o.Id = i.Id' },
    @{ Name = 'joined three-part source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN UtilityDb.dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'joined four-part source'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN ServerA.UtilityDb.dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'second statement after join'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id; DELETE FROM dbo.Items' },
    @{ Name = 'semicolonless batch after join'; Sql = "SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON o.Id = i.Id`r`nWAITFOR DELAY '00:00:01'" },
    @{ Name = 'JOIN token in SELECT list'; Sql = 'SELECT JOIN FROM dbo.Items' },
    @{ Name = 'JOIN token after WHERE'; Sql = 'SELECT i.Id FROM dbo.Items i WHERE i.Id = 1 JOIN dbo.Other o ON o.Id = i.Id' },
    @{ Name = 'unclosed ON parenthesis'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN dbo.Other o ON (o.Id = i.Id' },
    @{ Name = 'unclosed comment at join boundary'; Sql = 'SELECT i.Id FROM dbo.Items i JOIN /* open dbo.Other o ON o.Id = i.Id' }
)

$rejected += @(
    @{ Name = 'SHUTDOWN cannot be an implicit table alias'; Sql = "SELECT * FROM dbo.Items`r`nSHUTDOWN" },
    @{ Name = 'CHECKPOINT cannot be an implicit table alias after a block comment'; Sql = 'SELECT * FROM dbo.Items /* boundary */ CHECKPOINT;' },
    @{ Name = 'RECONFIGURE cannot be an implicit table alias after a line comment'; Sql = "SELECT * FROM dbo.Items -- boundary`r`nRECONFIGURE;" },
    @{ Name = 'statement starter cannot be an implicit select alias'; Sql = 'SELECT Id SHUTDOWN FROM dbo.Items' },
    @{ Name = 'reserved word cannot be an implicit select alias'; Sql = 'SELECT Id DELETE FROM dbo.Items' },
    @{ Name = 'unclosed nested block comment'; Sql = 'SELECT /* outer /* inner */ Id FROM dbo.Items */' },
    @{ Name = 'CAST requires a type'; Sql = 'SELECT CAST(Id AS) FROM dbo.Items' },
    @{ Name = 'CAST requires a source expression'; Sql = 'SELECT CAST(AS int) FROM dbo.Items' },
    @{ Name = 'CAST rejects an empty type argument'; Sql = 'SELECT CAST(Id AS decimal()) FROM dbo.Items' },
    @{ Name = 'CAST rejects a missing scale'; Sql = 'SELECT CAST(Id AS decimal(18,)) FROM dbo.Items' },
    @{ Name = 'CAST rejects excess type arguments'; Sql = 'SELECT CAST(Id AS decimal(18,2,1)) FROM dbo.Items' },
    @{ Name = 'TRY_CAST rejects malformed precision and scale'; Sql = 'SELECT TRY_CAST(Id AS decimal(18 2)) FROM dbo.Items' },
    @{ Name = 'generic function cannot use AS'; Sql = 'SELECT COALESCE(Id AS bigint) FROM dbo.Items' },
    @{ Name = 'CAST AS cannot bypass expression boundary'; Sql = 'SELECT CAST(Id FROBNICATE AS bigint) FROM dbo.Items' },
    @{ Name = 'CAST rejects a second AS'; Sql = 'SELECT CAST(Id AS bigint AS ConvertedId) FROM dbo.Items' }
)

foreach ($case in $rejected) {
    $result = Test-SqlUtilityQuery -Sql $case.Sql
    Assert-True (-not $result.IsValid) "$($case.Name) is rejected"
    Assert-True ($result.IsValid -is [bool]) "$($case.Name) returns Boolean IsValid"
    Assert-True (-not [string]::IsNullOrWhiteSpace($result.ErrorMessage)) "$($case.Name) has a user-facing error"
    Assert-Equal '' $result.NormalizedSql "$($case.Name) has no executable SQL"
    Assert-Equal '' $result.TableIdentifier "$($case.Name) has no executable table"
    Assert-Equal $false $result.HasOrderBy "$($case.Name) cannot be treated as ordered"
    Assert-Equal '' $result.CountSourceSql "$($case.Name) has no count source"
}

Complete-TestFile 'All query policy tests passed.'
