function New-SqlUtilityInvalidQueryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Message
    )

    return [pscustomobject][ordered]@{
        IsValid = $false
        ErrorMessage = $Message
        NormalizedSql = ''
        TableIdentifier = ''
        HasOrderBy = $false
    }
}

function Get-SqlUtilitySqlTokens {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $Sql
    )

    $tokens = New-Object System.Collections.Generic.List[object]
    $length = $Sql.Length
    $index = 0
    $depth = 0

    while ($index -lt $length) {
        $character = $Sql[$index]

        if ([char]::IsWhiteSpace($character)) {
            $index++
            continue
        }

        if ($character -eq '-' -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq '-') {
            $index += 2
            while ($index -lt $length -and $Sql[$index] -ne "`r" -and $Sql[$index] -ne "`n") {
                $index++
            }
            continue
        }

        if ($character -eq '/' -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq '*') {
            $index += 2
            $closed = $false
            while ($index -lt $length) {
                if ($Sql[$index] -eq '*' -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq '/') {
                    $index += 2
                    $closed = $true
                    break
                }
                $index++
            }
            if (-not $closed) {
                throw [System.ArgumentException]::new('The SQL contains an unclosed block comment.')
            }
            continue
        }

        if ($character -eq "'") {
            $start = $index
            $index++
            $closed = $false
            while ($index -lt $length) {
                if ($Sql[$index] -eq "'") {
                    if (($index + 1) -lt $length -and $Sql[$index + 1] -eq "'") {
                        $index += 2
                        continue
                    }
                    $index++
                    $closed = $true
                    break
                }
                $index++
            }
            if (-not $closed) {
                throw [System.ArgumentException]::new('The SQL contains an unclosed string literal.')
            }

            $text = $Sql.Substring($start, $index - $start)
            [void] $tokens.Add([pscustomobject]@{
                Text = $text
                Upper = $text.ToUpperInvariant()
                Kind = 'Symbol'
                Depth = $depth
                Start = $start
                End = $index
            })
            continue
        }

        if ($character -eq '"') {
            $start = $index
            $index++
            $closed = $false
            while ($index -lt $length) {
                if ($Sql[$index] -eq '"') {
                    if (($index + 1) -lt $length -and $Sql[$index + 1] -eq '"') {
                        $index += 2
                        continue
                    }
                    $index++
                    $closed = $true
                    break
                }
                $index++
            }
            if (-not $closed) {
                throw [System.ArgumentException]::new('The SQL contains an unclosed quoted identifier.')
            }

            $text = $Sql.Substring($start, $index - $start)
            [void] $tokens.Add([pscustomobject]@{
                Text = $text
                Upper = $text.ToUpperInvariant()
                Kind = 'Identifier'
                Depth = $depth
                Start = $start
                End = $index
            })
            continue
        }

        if ($character -eq '[') {
            $start = $index
            $index++
            $closed = $false
            while ($index -lt $length) {
                if ($Sql[$index] -eq ']') {
                    if (($index + 1) -lt $length -and $Sql[$index + 1] -eq ']') {
                        $index += 2
                        continue
                    }
                    $index++
                    $closed = $true
                    break
                }
                $index++
            }
            if (-not $closed) {
                throw [System.ArgumentException]::new('The SQL contains an unclosed bracketed identifier.')
            }

            $text = $Sql.Substring($start, $index - $start)
            [void] $tokens.Add([pscustomobject]@{
                Text = $text
                Upper = $text.ToUpperInvariant()
                Kind = 'Identifier'
                Depth = $depth
                Start = $start
                End = $index
            })
            continue
        }

        if ([char]::IsLetterOrDigit($character) -or $character -eq '_' -or $character -eq '$') {
            $start = $index
            while ($index -lt $length) {
                $wordCharacter = $Sql[$index]
                if (-not ([char]::IsLetterOrDigit($wordCharacter) -or $wordCharacter -eq '_' -or $wordCharacter -eq '$')) {
                    break
                }
                $index++
            }

            $text = $Sql.Substring($start, $index - $start)
            [void] $tokens.Add([pscustomobject]@{
                Text = $text
                Upper = $text.ToUpperInvariant()
                Kind = 'Word'
                Depth = $depth
                Start = $start
                End = $index
            })
            continue
        }

        $start = $index
        $kind = 'Symbol'
        if ($character -eq ';') {
            $kind = 'Semicolon'
            $index++
        }
        elseif ($character -eq '(') {
            $index++
            $tokenDepth = $depth
            $depth++
        }
        elseif ($character -eq ')') {
            if ($depth -eq 0) {
                throw [System.ArgumentException]::new('The SQL contains an unexpected closing parenthesis.')
            }
            $depth--
            $tokenDepth = $depth
            $index++
        }
        else {
            $index++
        }

        if ($null -eq $tokenDepth) {
            $tokenDepth = $depth
        }
        $text = $Sql.Substring($start, $index - $start)
        [void] $tokens.Add([pscustomobject]@{
            Text = $text
            Upper = $text.ToUpperInvariant()
            Kind = $kind
            Depth = $tokenDepth
            Start = $start
            End = $index
        })
        Remove-Variable tokenDepth -ErrorAction SilentlyContinue
    }

    if ($depth -ne 0) {
        throw [System.ArgumentException]::new('The SQL contains an unclosed parenthesis.')
    }

    return $tokens.ToArray()
}

function Test-SqlUtilityIdentifierToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Token
    )

    if ($Token.Kind -eq 'Identifier') {
        return $true
    }
    if ($Token.Kind -ne 'Word') {
        return $false
    }

    return $Token.Text -match '^[\p{L}_][\p{L}\p{Nd}_@$]*$'
}

function Test-SqlUtilityScalarExpression {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [bool] $AllowAlias = $false,
        [bool] $AllowOrdering = $false
    )

    if ($Start -ge $End) {
        return $false
    }

    $expectOperand = $true
    $operatorContinuation = $false
    $caseDepth = 0
    $index = $Start
    $operatorSymbols = @('+', '-', '*', '/', '%', '=', '<', '>', '!', '~', '&', '|', '^')
    $binaryWordOperators = @('AND', 'OR', 'LIKE', 'IN', 'IS', 'BETWEEN', 'COLLATE', 'OVER')

    while ($index -lt $End) {
        $token = $Tokens[$index]

        if ($token.Depth -gt 0) {
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text -eq '(') {
            $closeIndex = $index + 1
            while ($closeIndex -lt $End) {
                if ($Tokens[$closeIndex].Depth -eq 0 -and $Tokens[$closeIndex].Kind -eq 'Symbol' -and $Tokens[$closeIndex].Text -eq ')') {
                    break
                }
                $closeIndex++
            }
            if ($closeIndex -ge $End) {
                return $false
            }
            if ($expectOperand -and $closeIndex -eq ($index + 1)) {
                return $false
            }

            $expectOperand = $false
            $operatorContinuation = $false
            $index = $closeIndex + 1
            continue
        }

        if ($token.Kind -eq 'Identifier' -or $token.Kind -eq 'Word') {
            if ($token.Kind -eq 'Word') {
                if ($token.Upper -eq 'CASE') {
                    if (-not $expectOperand) {
                        return $false
                    }
                    $caseDepth++
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'WHEN') {
                    if ($caseDepth -eq 0) {
                        return $false
                    }
                    $expectOperand = $true
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'THEN' -or $token.Upper -eq 'ELSE') {
                    if ($caseDepth -eq 0 -or $expectOperand) {
                        return $false
                    }
                    $expectOperand = $true
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'END') {
                    if ($caseDepth -eq 0 -or $expectOperand) {
                        return $false
                    }
                    $caseDepth--
                    $expectOperand = $false
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'NOT') {
                    if (-not $expectOperand) {
                        if (
                            ($index + 1) -ge $End -or
                            $Tokens[$index + 1].Kind -ne 'Word' -or
                            @('LIKE', 'IN', 'BETWEEN') -notcontains $Tokens[$index + 1].Upper
                        ) {
                            return $false
                        }
                        $expectOperand = $true
                        $operatorContinuation = $false
                        $index += 2
                        continue
                    }
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($binaryWordOperators -contains $token.Upper) {
                    if ($expectOperand) {
                        return $false
                    }
                    $expectOperand = $true
                    $operatorContinuation = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'ASC' -or $token.Upper -eq 'DESC') {
                    if (-not $AllowOrdering -or $expectOperand -or $index -ne ($End - 1)) {
                        return $false
                    }
                    $index++
                    continue
                }
                if ($token.Upper -eq 'AS') {
                    if (-not $AllowAlias -or $expectOperand -or $caseDepth -ne 0 -or ($index + 2) -ne $End) {
                        return $false
                    }
                    if (-not (Test-SqlUtilityIdentifierToken -Token $Tokens[$index + 1])) {
                        return $false
                    }
                    return $true
                }
            }

            if (-not $expectOperand) {
                return $false
            }
            $expectOperand = $false
            $operatorContinuation = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text.StartsWith("'")) {
            if (-not $expectOperand) {
                return $false
            }
            $expectOperand = $false
            $operatorContinuation = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text -eq '.') {
            if ($expectOperand -or ($index + 1) -ge $End) {
                return $false
            }

            $previousToken = $Tokens[$index - 1]
            $nextToken = $Tokens[$index + 1]
            if (
                $previousToken.Kind -eq 'Word' -and $previousToken.Text -match '^\d+$' -and
                $nextToken.Kind -eq 'Word' -and $nextToken.Text -match '^\d+$'
            ) {
                $expectOperand = $false
                $operatorContinuation = $false
                $index += 2
                continue
            }
            if (
                -not (Test-SqlUtilityIdentifierToken -Token $nextToken) -and
                -not ($nextToken.Kind -eq 'Symbol' -and $nextToken.Text -eq '*')
            ) {
                return $false
            }
            $expectOperand = $true
            $operatorContinuation = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $operatorSymbols -contains $token.Text) {
            if ($expectOperand) {
                if ($token.Text -eq '*') {
                    $expectOperand = $false
                    $operatorContinuation = $false
                }
                elseif ($token.Text -eq '+' -or $token.Text -eq '-' -or $token.Text -eq '~' -or $operatorContinuation) {
                    $operatorContinuation = $true
                }
                else {
                    return $false
                }
            }
            else {
                $expectOperand = $true
                $operatorContinuation = $true
            }
            $index++
            continue
        }

        return $false
    }

    return (-not $expectOperand) -and $caseDepth -eq 0
}

function Test-SqlUtilityExpressionList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [bool] $AllowComma = $false,
        [bool] $AllowAlias = $false,
        [bool] $AllowOrdering = $false
    )

    if ($Start -ge $End) {
        return $false
    }

    $itemStart = $Start
    for ($index = $Start; $index -lt $End; $index++) {
        $token = $Tokens[$index]
        if ($token.Depth -eq 0 -and $token.Kind -eq 'Symbol' -and $token.Text -eq ',') {
            if (-not $AllowComma) {
                return $false
            }
            if (-not (Test-SqlUtilityScalarExpression -Tokens $Tokens -Start $itemStart -End $index -AllowAlias $AllowAlias -AllowOrdering $AllowOrdering)) {
                return $false
            }
            $itemStart = $index + 1
        }
    }

    return Test-SqlUtilityScalarExpression -Tokens $Tokens -Start $itemStart -End $End -AllowAlias $AllowAlias -AllowOrdering $AllowOrdering
}

function Test-SqlUtilityQuery {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string] $Sql
    )

    if ([string]::IsNullOrWhiteSpace($Sql)) {
        return New-SqlUtilityInvalidQueryResult -Message 'Enter a SELECT query.'
    }

    try {
        $tokens = @(Get-SqlUtilitySqlTokens -Sql $Sql)
    }
    catch [System.ArgumentException] {
        return New-SqlUtilityInvalidQueryResult -Message $_.Exception.Message
    }

    if ($tokens.Count -eq 0) {
        return New-SqlUtilityInvalidQueryResult -Message 'Enter a SELECT query.'
    }

    $semicolonIndexes = @()
    for ($tokenIndex = 0; $tokenIndex -lt $tokens.Count; $tokenIndex++) {
        if ($tokens[$tokenIndex].Kind -eq 'Semicolon') {
            $semicolonIndexes += $tokenIndex
        }
    }

    $normalizedSql = $Sql
    if ($semicolonIndexes.Count -gt 0) {
        if ($semicolonIndexes.Count -ne 1 -or $semicolonIndexes[0] -ne ($tokens.Count - 1)) {
            return New-SqlUtilityInvalidQueryResult -Message 'Only one SELECT statement is allowed.'
        }

        $semicolon = $tokens[$semicolonIndexes[0]]
        $normalizedSql = $Sql.Remove($semicolon.Start, $semicolon.End - $semicolon.Start)
        $tokens = @($tokens | Select-Object -First ($tokens.Count - 1))
    }

    if ($tokens.Count -eq 0 -or $tokens[0].Kind -ne 'Word' -or $tokens[0].Upper -ne 'SELECT' -or $tokens[0].Depth -ne 0) {
        return New-SqlUtilityInvalidQueryResult -Message 'The query must begin with SELECT.'
    }

    $denied = @(
        'INSERT','UPDATE','DELETE','MERGE','DROP','ALTER','CREATE','TRUNCATE',
        'EXEC','EXECUTE','DECLARE','SET','USE','GRANT','REVOKE','DENY',
        'BEGIN','COMMIT','ROLLBACK','BACKUP','RESTORE','DBCC','BULK',
        'JOIN','APPLY','UNION','INTERSECT','EXCEPT','INTO','TOP','OFFSET','FETCH',
        'OPENQUERY','OPENROWSET','OPENDATASOURCE'
    )

    for ($tokenIndex = 0; $tokenIndex -lt $tokens.Count; $tokenIndex++) {
        $token = $tokens[$tokenIndex]
        if ($token.Kind -ne 'Word') {
            continue
        }
        if ($denied -contains $token.Upper) {
            return New-SqlUtilityInvalidQueryResult -Message "'$($token.Text)' is not allowed in a read-only query."
        }
        if ($token.Upper -eq 'SELECT' -and $tokenIndex -ne 0) {
            return New-SqlUtilityInvalidQueryResult -Message 'Nested or additional SELECT statements are not allowed.'
        }
    }

    $fromIndexes = @()
    for ($tokenIndex = 0; $tokenIndex -lt $tokens.Count; $tokenIndex++) {
        $token = $tokens[$tokenIndex]
        if ($token.Kind -eq 'Word' -and $token.Upper -eq 'FROM' -and $token.Depth -eq 0) {
            $fromIndexes += $tokenIndex
        }
    }
    if ($fromIndexes.Count -ne 1) {
        return New-SqlUtilityInvalidQueryResult -Message 'The query must contain exactly one FROM clause.'
    }

    $selectExpressionStart = 1
    if (
        $selectExpressionStart -lt $fromIndexes[0] -and
        $tokens[$selectExpressionStart].Kind -eq 'Word' -and
        $tokens[$selectExpressionStart].Depth -eq 0 -and
        ($tokens[$selectExpressionStart].Upper -eq 'DISTINCT' -or $tokens[$selectExpressionStart].Upper -eq 'ALL')
    ) {
        $selectExpressionStart++
    }
    if (-not (Test-SqlUtilityExpressionList -Tokens $tokens -Start $selectExpressionStart -End $fromIndexes[0] -AllowComma $true -AllowAlias $true)) {
        return New-SqlUtilityInvalidQueryResult -Message 'The SELECT list contains an invalid expression.'
    }

    $hasOrderBy = $false
    for ($tokenIndex = 0; $tokenIndex -lt $tokens.Count; $tokenIndex++) {
        $token = $tokens[$tokenIndex]
        if ($token.Kind -ne 'Word' -or $token.Depth -ne 0) {
            continue
        }
        if ($token.Upper -eq 'GROUP' -or $token.Upper -eq 'ORDER') {
            if (($tokenIndex + 1) -ge $tokens.Count -or $tokens[$tokenIndex + 1].Kind -ne 'Word' -or $tokens[$tokenIndex + 1].Upper -ne 'BY' -or $tokens[$tokenIndex + 1].Depth -ne 0) {
                return New-SqlUtilityInvalidQueryResult -Message "'$($token.Text)' must be followed by BY."
            }
            if ($token.Upper -eq 'ORDER') {
                $hasOrderBy = $true
            }
        }
    }

    $sourceIndex = $fromIndexes[0] + 1
    if ($sourceIndex -ge $tokens.Count -or $tokens[$sourceIndex].Depth -ne 0 -or -not (Test-SqlUtilityIdentifierToken -Token $tokens[$sourceIndex])) {
        return New-SqlUtilityInvalidQueryResult -Message 'FROM must be followed by a table identifier.'
    }

    $tableStart = $tokens[$sourceIndex].Start
    $tableEnd = $tokens[$sourceIndex].End
    $sourceIndex++

    if ($sourceIndex -lt $tokens.Count -and $tokens[$sourceIndex].Depth -eq 0 -and $tokens[$sourceIndex].Kind -eq 'Symbol' -and $tokens[$sourceIndex].Text -eq '.') {
        $sourceIndex++
        if ($sourceIndex -ge $tokens.Count -or $tokens[$sourceIndex].Depth -ne 0 -or -not (Test-SqlUtilityIdentifierToken -Token $tokens[$sourceIndex])) {
            return New-SqlUtilityInvalidQueryResult -Message 'The table identifier is malformed.'
        }
        $tableEnd = $tokens[$sourceIndex].End
        $sourceIndex++
    }

    $clauseWords = @('WHERE', 'GROUP', 'HAVING', 'ORDER')
    $atClauseBoundary = $sourceIndex -ge $tokens.Count -or (
        $tokens[$sourceIndex].Depth -eq 0 -and
        $tokens[$sourceIndex].Kind -eq 'Word' -and
        $clauseWords -contains $tokens[$sourceIndex].Upper
    )

    if (-not $atClauseBoundary) {
        if ($tokens[$sourceIndex].Depth -ne 0 -or -not (Test-SqlUtilityIdentifierToken -Token $tokens[$sourceIndex])) {
            return New-SqlUtilityInvalidQueryResult -Message 'The table source must be a single table with an optional alias.'
        }

        if ($tokens[$sourceIndex].Kind -eq 'Word' -and $tokens[$sourceIndex].Upper -eq 'AS') {
            $sourceIndex++
            if (
                $sourceIndex -ge $tokens.Count -or
                $tokens[$sourceIndex].Depth -ne 0 -or
                -not (Test-SqlUtilityIdentifierToken -Token $tokens[$sourceIndex]) -or
                ($tokens[$sourceIndex].Kind -eq 'Word' -and $clauseWords -contains $tokens[$sourceIndex].Upper)
            ) {
                return New-SqlUtilityInvalidQueryResult -Message 'AS must be followed by a table alias.'
            }
            $sourceIndex++
        }
        else {
            if (-not (Test-SqlUtilityIdentifierToken -Token $tokens[$sourceIndex])) {
                return New-SqlUtilityInvalidQueryResult -Message 'The table alias is malformed.'
            }
            $sourceIndex++
        }

        $atClauseBoundary = $sourceIndex -ge $tokens.Count -or (
            $tokens[$sourceIndex].Depth -eq 0 -and
            $tokens[$sourceIndex].Kind -eq 'Word' -and
            $clauseWords -contains $tokens[$sourceIndex].Upper
        )
        if (-not $atClauseBoundary) {
            return New-SqlUtilityInvalidQueryResult -Message 'Only one table and one optional alias are allowed after FROM.'
        }
    }

    $clauses = New-Object System.Collections.Generic.List[object]
    for ($tokenIndex = $sourceIndex; $tokenIndex -lt $tokens.Count; $tokenIndex++) {
        $token = $tokens[$tokenIndex]
        if ($token.Depth -ne 0 -or $token.Kind -ne 'Word') {
            continue
        }

        $clauseName = $null
        $rank = 0
        $contentStart = $tokenIndex + 1
        if ($token.Upper -eq 'WHERE') {
            $clauseName = 'WHERE'
            $rank = 1
        }
        elseif ($token.Upper -eq 'GROUP') {
            $clauseName = 'GROUP BY'
            $rank = 2
            $contentStart++
        }
        elseif ($token.Upper -eq 'HAVING') {
            $clauseName = 'HAVING'
            $rank = 3
        }
        elseif ($token.Upper -eq 'ORDER') {
            $clauseName = 'ORDER BY'
            $rank = 4
            $contentStart++
        }

        if ($null -ne $clauseName) {
            [void] $clauses.Add([pscustomobject]@{
                Name = $clauseName
                Rank = $rank
                Start = $tokenIndex
                ContentStart = $contentStart
            })
        }
    }

    if ($sourceIndex -lt $tokens.Count -and ($clauses.Count -eq 0 -or $clauses[0].Start -ne $sourceIndex)) {
        return New-SqlUtilityInvalidQueryResult -Message 'Unexpected tokens follow the table source.'
    }

    $previousRank = 0
    for ($clauseIndex = 0; $clauseIndex -lt $clauses.Count; $clauseIndex++) {
        $clause = $clauses[$clauseIndex]
        if ($clause.Rank -le $previousRank) {
            return New-SqlUtilityInvalidQueryResult -Message 'Query clauses must appear once and in WHERE, GROUP BY, HAVING, ORDER BY order.'
        }
        $previousRank = $clause.Rank

        $contentEnd = if (($clauseIndex + 1) -lt $clauses.Count) { $clauses[$clauseIndex + 1].Start } else { $tokens.Count }
        $allowComma = $clause.Name -eq 'GROUP BY' -or $clause.Name -eq 'ORDER BY'
        $allowOrdering = $clause.Name -eq 'ORDER BY'
        if (-not (Test-SqlUtilityExpressionList -Tokens $tokens -Start $clause.ContentStart -End $contentEnd -AllowComma $allowComma -AllowOrdering $allowOrdering)) {
            return New-SqlUtilityInvalidQueryResult -Message "The $($clause.Name) clause contains an invalid expression or trailing statement."
        }
    }

    return [pscustomobject][ordered]@{
        IsValid = $true
        ErrorMessage = ''
        NormalizedSql = $normalizedSql
        TableIdentifier = $Sql.Substring($tableStart, $tableEnd - $tableStart)
        HasOrderBy = [bool] $hasOrderBy
    }
}
