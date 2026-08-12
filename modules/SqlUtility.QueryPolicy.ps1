$script:SqlUtilityReservedAliasWords = @{}
foreach ($reservedWord in @(
    'ADD','ALL','ALTER','AND','ANY','AS','ASC','AUTHORIZATION','BACKUP','BEGIN','BETWEEN','BREAK','BROWSE','BULK','BY',
    'CASCADE','CASE','CAST','CHECK','CHECKPOINT','CLOSE','CLUSTERED','COALESCE','COLLATE','COLUMN','COMMIT','COMPUTE',
    'CONSTRAINT','CONTAINS','CONTAINSTABLE','CONTINUE','CONVERT','CREATE','CROSS','CURRENT','CURRENT_DATE','CURRENT_TIME',
    'CURRENT_TIMESTAMP','CURRENT_USER','CURSOR','DATABASE','DBCC','DEALLOCATE','DECLARE','DEFAULT','DELETE','DENY','DESC',
    'DISABLE','DISK','DISTINCT','DISTRIBUTED','DOUBLE','DROP','DUMP','ELSE','ENABLE','END','ERRLVL','ESCAPE','EXCEPT',
    'EXEC','EXECUTE','EXISTS','EXIT','EXTERNAL','FETCH','FILE','FILLFACTOR','FOR','FOREIGN','FREETEXT','FREETEXTTABLE',
    'FROM','FULL','FUNCTION','GET','GOTO','GRANT','GROUP','HAVING','HOLDLOCK','IDENTITY','IDENTITY_INSERT','IDENTITYCOL',
    'IF','IN','INDEX','INNER','INSERT','INTERSECT','INTO','IS','JOIN','KEY','KILL','LEFT','LIKE','LINENO','LOAD','MERGE',
    'MOVE','NATIONAL','NOCHECK','NONCLUSTERED','NOT','NULL','NULLIF','OF','OFF','OFFSETS','ON','OPEN','OPENDATASOURCE',
    'OPENQUERY','OPENROWSET','OPENXML','OPTION','OR','ORDER','OUTER','OVER','PERCENT','PIVOT','PLAN','PRECISION','PRIMARY',
    'PRINT','PROC','PROCEDURE','PUBLIC','RAISERROR','READ','READTEXT','RECEIVE','RECONFIGURE','REFERENCES','RENAME',
    'REPLICATION','RESTORE','RESTRICT','RETURN','REVERT','REVOKE','RIGHT','ROLLBACK','ROWCOUNT','ROWGUIDCOL','RULE','SAVE',
    'SCHEMA','SECURITYAUDIT','SELECT','SEMANTICKEYPHRASETABLE','SEMANTICSIMILARITYDETAILSTABLE','SEMANTICSIMILARITYTABLE',
    'SEND','SESSION_USER','SET','SETUSER','SHUTDOWN','SOME','STATISTICS','SYSTEM_USER','TABLE','TABLESAMPLE','TEXTSIZE',
    'THEN','THROW','TO','TOP','TRAN','TRANSACTION','TRIGGER','TRUNCATE','TRY_CAST','TRY_CONVERT','TSEQUAL','UNION',
    'UNIQUE','UNPIVOT','UPDATE','UPDATETEXT','USE','USER','VALUES','VARYING','VIEW','WAITFOR','WHEN','WHERE','WHILE',
    'WITH','WRITETEXT'
)) {
    $script:SqlUtilityReservedAliasWords[$reservedWord] = $true
}

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
        CountSourceSql = ''
    }
}

function New-SqlUtilityCountSql {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string] $CountSourceSql,
        [int] $OutputColumnCount
    )

    if ([string]::IsNullOrWhiteSpace($CountSourceSql)) {
        throw [System.ArgumentException]::new('Count source SQL cannot be empty.', 'CountSourceSql')
    }
    if ($OutputColumnCount -lt 1) {
        throw [System.ArgumentOutOfRangeException]::new('OutputColumnCount', $OutputColumnCount, 'At least one output column is required.')
    }

    $columnAliases = @(
        for ($index = 1; $index -le $OutputColumnCount; $index++) {
            '[SqlUtilityCountColumn{0}]' -f $index
        }
    )

    return "SELECT COUNT_BIG(*)`r`nFROM (`r`n{0}`r`n) AS [SqlUtilityCountSource] ({1});" -f `
        $CountSourceSql.Trim(), ($columnAliases -join ', ')
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
            $commentDepth = 1
            while ($index -lt $length -and $commentDepth -gt 0) {
                if ($Sql[$index] -eq '/' -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq '*') {
                    $commentDepth++
                    $index += 2
                    continue
                }
                if ($Sql[$index] -eq '*' -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq '/') {
                    $commentDepth--
                    $index += 2
                    continue
                }
                $index++
            }
            if ($commentDepth -ne 0) {
                throw [System.ArgumentException]::new('The SQL contains an unclosed block comment.')
            }
            continue
        }

        $isUnicodeString = ($character -eq 'N' -or $character -eq 'n') -and ($index + 1) -lt $length -and $Sql[$index + 1] -eq "'"
        if ($character -eq "'" -or $isUnicodeString) {
            $start = $index
            if ($isUnicodeString) {
                $index++
            }
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
        $tokenDepth = $depth
        if ($character -eq ';') {
            $kind = 'Semicolon'
            $index++
        }
        elseif ($character -eq '(') {
            $index++
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

        $text = $Sql.Substring($start, $index - $start)
        [void] $tokens.Add([pscustomobject]@{
            Text = $text
            Upper = $text.ToUpperInvariant()
            Kind = $kind
            Depth = $tokenDepth
            Start = $start
            End = $index
        })
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

function Test-SqlUtilityAliasToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Token
    )

    if ($Token.Kind -eq 'Identifier') {
        return $true
    }
    if (-not (Test-SqlUtilityIdentifierToken -Token $Token)) {
        return $false
    }

    return -not $script:SqlUtilityReservedAliasWords.ContainsKey($Token.Upper)
}

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

function Test-SqlUtilityCastType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [Parameter(Mandatory = $true)][int] $BaseDepth
    )

    if (
        $Start -ge $End -or
        $Tokens[$Start].Depth -ne $BaseDepth -or
        -not (Test-SqlUtilityIdentifierToken -Token $Tokens[$Start])
    ) {
        return $false
    }

    $index = $Start + 1
    if (
        $index -lt $End -and
        $Tokens[$index].Depth -eq $BaseDepth -and
        $Tokens[$index].Kind -eq 'Symbol' -and
        $Tokens[$index].Text -eq '.'
    ) {
        $index++
        if (
            $index -ge $End -or
            $Tokens[$index].Depth -ne $BaseDepth -or
            -not (Test-SqlUtilityIdentifierToken -Token $Tokens[$index])
        ) {
            return $false
        }
        $index++
    }

    if ($index -eq $End) {
        return $true
    }
    if (
        $Tokens[$index].Depth -ne $BaseDepth -or
        $Tokens[$index].Kind -ne 'Symbol' -or
        $Tokens[$index].Text -ne '(' -or
        $Tokens[$End - 1].Depth -ne $BaseDepth -or
        $Tokens[$End - 1].Kind -ne 'Symbol' -or
        $Tokens[$End - 1].Text -ne ')'
    ) {
        return $false
    }

    $argumentStart = $index + 1
    $argumentCount = ($End - 1) - $argumentStart
    if ($argumentCount -eq 1) {
        $argument = $Tokens[$argumentStart]
        return $argument.Depth -eq ($BaseDepth + 1) -and
            $argument.Kind -eq 'Word' -and
            ($argument.Text -match '^\d+$' -or $argument.Upper -eq 'MAX')
    }
    if ($argumentCount -eq 3) {
        return $Tokens[$argumentStart].Depth -eq ($BaseDepth + 1) -and
            $Tokens[$argumentStart].Kind -eq 'Word' -and $Tokens[$argumentStart].Text -match '^\d+$' -and
            $Tokens[$argumentStart + 1].Depth -eq ($BaseDepth + 1) -and
            $Tokens[$argumentStart + 1].Kind -eq 'Symbol' -and $Tokens[$argumentStart + 1].Text -eq ',' -and
            $Tokens[$argumentStart + 2].Depth -eq ($BaseDepth + 1) -and
            $Tokens[$argumentStart + 2].Kind -eq 'Word' -and $Tokens[$argumentStart + 2].Text -match '^\d+$'
    }

    return $false
}

function Test-SqlUtilityCastExpression {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [Parameter(Mandatory = $true)][int] $BaseDepth
    )

    $asIndex = -1
    for ($index = $Start; $index -lt $End; $index++) {
        if (
            $Tokens[$index].Depth -eq $BaseDepth -and
            $Tokens[$index].Kind -eq 'Word' -and
            $Tokens[$index].Upper -eq 'AS'
        ) {
            if ($asIndex -ne -1) {
                return $false
            }
            $asIndex = $index
        }
    }

    if ($asIndex -le $Start -or $asIndex -ge ($End - 1)) {
        return $false
    }
    if (-not (Test-SqlUtilityScalarExpression -Tokens $Tokens -Start $Start -End $asIndex -BaseDepth $BaseDepth)) {
        return $false
    }

    return Test-SqlUtilityCastType -Tokens $Tokens -Start ($asIndex + 1) -End $End -BaseDepth $BaseDepth
}

function Test-SqlUtilityWindowExpression {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [Parameter(Mandatory = $true)][int] $BaseDepth
    )

    if ($Start -ge $End) {
        return $true
    }

    $partitionStart = $null
    $orderStart = $null
    if (
        ($Start + 1) -lt $End -and
        $Tokens[$Start].Depth -eq $BaseDepth -and $Tokens[$Start].Kind -eq 'Word' -and $Tokens[$Start].Upper -eq 'PARTITION' -and
        $Tokens[$Start + 1].Depth -eq $BaseDepth -and $Tokens[$Start + 1].Kind -eq 'Word' -and $Tokens[$Start + 1].Upper -eq 'BY'
    ) {
        $partitionStart = $Start + 2
    }
    elseif (
        ($Start + 1) -lt $End -and
        $Tokens[$Start].Depth -eq $BaseDepth -and $Tokens[$Start].Kind -eq 'Word' -and $Tokens[$Start].Upper -eq 'ORDER' -and
        $Tokens[$Start + 1].Depth -eq $BaseDepth -and $Tokens[$Start + 1].Kind -eq 'Word' -and $Tokens[$Start + 1].Upper -eq 'BY'
    ) {
        $orderStart = $Start + 2
    }
    else {
        return $false
    }

    if ($null -ne $partitionStart) {
        for ($index = $partitionStart; $index -lt ($End - 1); $index++) {
            if (
                $Tokens[$index].Depth -eq $BaseDepth -and $Tokens[$index].Kind -eq 'Word' -and $Tokens[$index].Upper -eq 'ORDER' -and
                $Tokens[$index + 1].Depth -eq $BaseDepth -and $Tokens[$index + 1].Kind -eq 'Word' -and $Tokens[$index + 1].Upper -eq 'BY'
            ) {
                $orderStart = $index + 2
                if (-not (Test-SqlUtilityExpressionList -Tokens $Tokens -Start $partitionStart -End $index -BaseDepth $BaseDepth -AllowComma $true)) {
                    return $false
                }
                break
            }
        }
        if ($null -eq $orderStart) {
            return Test-SqlUtilityExpressionList -Tokens $Tokens -Start $partitionStart -End $End -BaseDepth $BaseDepth -AllowComma $true
        }
    }

    return Test-SqlUtilityExpressionList -Tokens $Tokens -Start $orderStart -End $End -BaseDepth $BaseDepth -AllowComma $true -AllowOrdering $true
}

function Test-SqlUtilityScalarExpression {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [int] $BaseDepth = 0,
        [bool] $AllowAlias = $false,
        [bool] $AllowOrdering = $false
    )

    if ($Start -ge $End) {
        return $false
    }

    $expectOperand = $true
    $expectWindowGroup = $false
    $canCall = $false
    $callName = ''
    $caseDepth = 0
    $index = $Start
    $binarySymbols = @('+', '-', '*', '/', '%', '=', '<', '>', '&', '|', '^')
    $compoundSymbols = @('<>', '<=', '>=', '!=', '!<', '!>')
    $binaryWordOperators = @('AND', 'OR', 'LIKE', 'IN', 'IS', 'BETWEEN', 'COLLATE')

    while ($index -lt $End) {
        $token = $Tokens[$index]

        if ($token.Depth -gt $BaseDepth) {
            $index++
            continue
        }
        if ($token.Depth -lt $BaseDepth) {
            return $false
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text -eq '(') {
            $closeIndex = $index + 1
            while ($closeIndex -lt $End) {
                if ($Tokens[$closeIndex].Depth -eq $BaseDepth -and $Tokens[$closeIndex].Kind -eq 'Symbol' -and $Tokens[$closeIndex].Text -eq ')') {
                    break
                }
                $closeIndex++
            }
            if ($closeIndex -ge $End) {
                return $false
            }

            if ($expectWindowGroup) {
                if (-not (Test-SqlUtilityWindowExpression -Tokens $Tokens -Start ($index + 1) -End $closeIndex -BaseDepth ($BaseDepth + 1))) {
                    return $false
                }
            }
            elseif ($expectOperand) {
                if (-not (Test-SqlUtilityExpressionList -Tokens $Tokens -Start ($index + 1) -End $closeIndex -BaseDepth ($BaseDepth + 1) -AllowComma $true)) {
                    return $false
                }
            }
            else {
                if (-not $canCall) {
                    return $false
                }
                if ($callName -eq 'CAST' -or $callName -eq 'TRY_CAST') {
                    if (-not (Test-SqlUtilityCastExpression -Tokens $Tokens -Start ($index + 1) -End $closeIndex -BaseDepth ($BaseDepth + 1))) {
                        return $false
                    }
                }
                elseif (($index + 1) -lt $closeIndex) {
                    $argumentStart = $index + 1
                    if (
                        $Tokens[$argumentStart].Depth -eq ($BaseDepth + 1) -and
                        $Tokens[$argumentStart].Kind -eq 'Word' -and
                        ($Tokens[$argumentStart].Upper -eq 'DISTINCT' -or $Tokens[$argumentStart].Upper -eq 'ALL')
                    ) {
                        $argumentStart++
                    }
                    if (-not (Test-SqlUtilityExpressionList -Tokens $Tokens -Start $argumentStart -End $closeIndex -BaseDepth ($BaseDepth + 1) -AllowComma $true)) {
                        return $false
                    }
                }
            }

            $expectOperand = $false
            $expectWindowGroup = $false
            $canCall = $false
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
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'WHEN') {
                    if ($caseDepth -eq 0) {
                        return $false
                    }
                    $expectOperand = $true
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'THEN' -or $token.Upper -eq 'ELSE') {
                    if ($caseDepth -eq 0 -or $expectOperand) {
                        return $false
                    }
                    $expectOperand = $true
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'END') {
                    if ($caseDepth -eq 0 -or $expectOperand) {
                        return $false
                    }
                    $caseDepth--
                    $expectOperand = $false
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'NOT') {
                    if (-not $expectOperand) {
                        if (
                            ($index + 1) -ge $End -or
                            $Tokens[$index + 1].Depth -ne $BaseDepth -or
                            $Tokens[$index + 1].Kind -ne 'Word' -or
                            @('LIKE', 'IN', 'BETWEEN') -notcontains $Tokens[$index + 1].Upper
                        ) {
                            return $false
                        }
                        $expectOperand = $true
                        $canCall = $false
                        $index += 2
                        continue
                    }
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'OVER') {
                    if ($expectOperand) {
                        return $false
                    }
                    $expectOperand = $true
                    $expectWindowGroup = $true
                    $canCall = $false
                    $index++
                    continue
                }
                if ($binaryWordOperators -contains $token.Upper) {
                    if ($expectOperand) {
                        return $false
                    }
                    $expectOperand = $true
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'ASC' -or $token.Upper -eq 'DESC') {
                    if (-not $AllowOrdering -or $expectOperand -or $index -ne ($End - 1)) {
                        return $false
                    }
                    $canCall = $false
                    $index++
                    continue
                }
                if ($token.Upper -eq 'AS') {
                    if (-not $AllowAlias -or $expectOperand -or $caseDepth -ne 0 -or ($index + 2) -ne $End) {
                        return $false
                    }
                    if (-not (Test-SqlUtilityAliasToken -Token $Tokens[$index + 1])) {
                        return $false
                    }
                    return $true
                }
            }

            if (-not $expectOperand) {
                if ($AllowAlias -and $caseDepth -eq 0 -and $index -eq ($End - 1) -and (Test-SqlUtilityAliasToken -Token $token)) {
                    return $true
                }
                return $false
            }
            $expectOperand = $false
            $canCall = $token.Kind -eq 'Identifier' -or $token.Text -match '^[\p{L}_]'
            $callName = if ($token.Kind -eq 'Word') { $token.Upper } else { '' }
            $index++
            continue
        }

        if (
            $token.Kind -eq 'Symbol' -and
            ($token.Text.StartsWith("'") -or $token.Text -match "^[Nn]'")
        ) {
            if (-not $expectOperand) {
                return $false
            }
            $expectOperand = $false
            $canCall = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text -eq '.') {
            if (($index + 1) -ge $End) {
                return $false
            }

            $nextToken = $Tokens[$index + 1]
            if ($expectOperand) {
                if ($nextToken.Depth -ne $BaseDepth -or $nextToken.Kind -ne 'Word' -or $nextToken.Text -notmatch '^\d+$') {
                    return $false
                }
                $expectOperand = $false
                $canCall = $false
                $index += 2
                continue
            }

            $previousToken = $Tokens[$index - 1]
            if (
                $previousToken.Kind -eq 'Word' -and $previousToken.Text -match '^\d+$' -and
                $nextToken.Depth -eq $BaseDepth -and $nextToken.Kind -eq 'Word' -and $nextToken.Text -match '^\d+$'
            ) {
                $expectOperand = $false
                $canCall = $false
                $index += 2
                continue
            }
            if (
                $nextToken.Depth -ne $BaseDepth -or
                (-not (Test-SqlUtilityIdentifierToken -Token $nextToken) -and -not ($nextToken.Kind -eq 'Symbol' -and $nextToken.Text -eq '*'))
            ) {
                return $false
            }
            $expectOperand = $true
            $canCall = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and (@('+', '-', '~') -contains $token.Text) -and $expectOperand) {
            $canCall = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and $token.Text -eq '*' -and $expectOperand) {
            $expectOperand = $false
            $canCall = $false
            $index++
            continue
        }

        if ($token.Kind -eq 'Symbol' -and -not $expectOperand) {
            $operatorLength = 1
            $operatorText = $token.Text
            if (
                ($index + 1) -lt $End -and
                $Tokens[$index + 1].Depth -eq $BaseDepth -and
                $Tokens[$index + 1].Kind -eq 'Symbol' -and
                $token.End -eq $Tokens[$index + 1].Start
            ) {
                $compoundOperator = $token.Text + $Tokens[$index + 1].Text
                if ($compoundSymbols -contains $compoundOperator) {
                    $operatorText = $compoundOperator
                    $operatorLength = 2
                }
            }
            if ($operatorLength -eq 1 -and $binarySymbols -notcontains $operatorText) {
                return $false
            }
            if ($operatorText -eq '!') {
                return $false
            }

            $expectOperand = $true
            $canCall = $false
            $index += $operatorLength
            continue
        }

        return $false
    }

    return (-not $expectOperand) -and (-not $expectWindowGroup) -and $caseDepth -eq 0
}

function Test-SqlUtilityExpressionList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $Tokens,
        [Parameter(Mandatory = $true)][int] $Start,
        [Parameter(Mandatory = $true)][int] $End,
        [int] $BaseDepth = 0,
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
        if ($token.Depth -eq $BaseDepth -and $token.Kind -eq 'Symbol' -and $token.Text -eq ',') {
            if (-not $AllowComma) {
                return $false
            }
            if (-not (Test-SqlUtilityScalarExpression -Tokens $Tokens -Start $itemStart -End $index -BaseDepth $BaseDepth -AllowAlias $AllowAlias -AllowOrdering $AllowOrdering)) {
                return $false
            }
            $itemStart = $index + 1
        }
    }

    return Test-SqlUtilityScalarExpression -Tokens $Tokens -Start $itemStart -End $End -BaseDepth $BaseDepth -AllowAlias $AllowAlias -AllowOrdering $AllowOrdering
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
        'APPLY','UNION','INTERSECT','EXCEPT','INTO','TOP','OFFSET','FETCH',
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
    $topLevelOrderToken = $null
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
                $topLevelOrderToken = $token
            }
        }
    }

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

    $countSourceSql = $normalizedSql
    if ($null -ne $topLevelOrderToken) {
        $countSourceSql = $normalizedSql.Substring(0, $topLevelOrderToken.Start).TrimEnd()
    }

    return [pscustomobject][ordered]@{
        IsValid = $true
        ErrorMessage = ''
        NormalizedSql = $normalizedSql
        TableIdentifier = $Sql.Substring($tableStart, $tableEnd - $tableStart)
        HasOrderBy = [bool] $hasOrderBy
        CountSourceSql = $countSourceSql
    }
}
