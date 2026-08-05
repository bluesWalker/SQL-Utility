$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-Helpers.ps1')
. (Join-Path $projectRoot 'modules\SqlUtility.Excel.ps1')

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipXmlDocument($Archive, [string] $EntryName) {
    $entry = $Archive.GetEntry($EntryName)
    Assert-True ($null -ne $entry) "Workbook contains $EntryName"
    if ($null -eq $entry) { return $null }

    $stream = $entry.Open()
    try {
        $document = [System.Xml.XmlDocument]::new()
        $document.PreserveWhitespace = $true
        $document.Load($stream)
        return (, $document)
    }
    finally {
        $stream.Dispose()
    }
}

function New-SpreadsheetNamespaceManager([System.Xml.XmlDocument] $Document) {
    $manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
    $manager.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $manager.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    return (, $manager)
}

function Assert-NoTemporaryFiles([string] $Directory, [string] $Message) {
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $Directory -Filter '*.tmp' -File).Count $Message
}

function New-ColumnBoundaryTable([int] $ColumnCount) {
    $table = [System.Data.DataTable]::new('Boundary')
    for ($number = 1; $number -le $ColumnCount; $number++) {
        [void] $table.Columns.Add("Column$number", [string])
    }
    return (, $table)
}

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SqlUtilityExcelTests-' + [guid]::NewGuid().ToString('N'))
[void] [System.IO.Directory]::CreateDirectory($testRoot)

try {
    $sourceTable = [System.Data.DataTable]::new('Mixed')
    [void] $sourceTable.Columns.Add('Integer', [int])
    [void] $sourceTable.Columns.Add('Decimal', [decimal])
    [void] $sourceTable.Columns.Add('Enabled', [bool])
    [void] $sourceTable.Columns.Add('OccurredAt', [datetime])
    [void] $sourceTable.Columns.Add('Text', [string])
    [void] $sourceTable.Columns.Add('FormulaText', [string])
    [void] $sourceTable.Columns.Add('NullableText', [string])

    $date = [datetime]::new(2026, 8, 2, 13, 45, 30, [DateTimeKind]::Unspecified)
    [void] $sourceTable.Rows.Add(42, [decimal]::Parse('1234.50', [Globalization.CultureInfo]::InvariantCulture), $true, $date, ' normal text ', '=1+1', [DBNull]::Value)
    [void] $sourceTable.Rows.Add(-7, [decimal]::Parse('-0.25', [Globalization.CultureInfo]::InvariantCulture), $false, $date.AddDays(1), 'second', '+cmd', 'value')
    [void] $sourceTable.Rows.Add(0, [decimal]::Zero, $true, $date.AddDays(2), 'third', '-2+3', 'value')
    [void] $sourceTable.Rows.Add(9, [decimal]::One, $false, $date.AddDays(3), 'fourth', '@name', 'value')

    $sourceEvents = [System.Collections.Generic.List[string]]::new()
    $sourceSchema = [pscustomobject]@{ Value = $null }
    $sourceRows = [System.Collections.Generic.List[object]]::new()
    $continueCalls = [pscustomobject]@{ Count = 0 }
    $rowSource = New-SqlUtilityDataTableRowSource -DataTable $sourceTable
    & $rowSource `
        { param($Schema) $sourceSchema.Value = @($Schema); [void] $sourceEvents.Add('schema') } `
        { param($Values) [void] $sourceRows.Add([object[]]$Values); [void] $sourceEvents.Add("row:$($Values[0])") } `
        { $continueCalls.Count++; return $true }

    Assert-Equal 7 $sourceSchema.Value.Count 'Cached DataTable row source emits the complete neutral schema'
    Assert-Equal 'Integer' $sourceSchema.Value[0].Name 'Cached DataTable row source emits the first column name'
    Assert-Equal ([int]) $sourceSchema.Value[0].DataType 'Cached DataTable row source emits the CLR data type'
    Assert-Equal 0 $sourceSchema.Value[0].Ordinal 'Cached DataTable row source emits the column ordinal'
    Assert-Equal 'schema,row:42,row:-7,row:0,row:9' ($sourceEvents -join ',') 'Cached DataTable row source emits schema once before rows in result order'
    Assert-Equal 4 $continueCalls.Count 'Cached DataTable row source checks continuation before each row'
    Assert-Equal 4 $sourceRows.Count 'Cached DataTable row source emits every cached row'
    Assert-True ([DBNull]::Value.Equals($sourceRows[0][6])) 'Cached DataTable row source preserves database null values'

    $workbookPath = Join-Path $testRoot 'mixed.xlsx'
    Export-SqlUtilityXlsx -DestinationPath $workbookPath -RowSource $rowSource -TimeoutSeconds 30
    Assert-True (Test-Path -LiteralPath $workbookPath -PathType Leaf) 'Mixed workbook is created'

    $archive = [System.IO.Compression.ZipFile]::OpenRead($workbookPath)
    try {
        $requiredEntries = @(
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            'xl/styles.xml',
            'xl/worksheets/sheet1.xml'
        )
        foreach ($entryName in $requiredEntries) {
            Assert-True ($null -ne $archive.GetEntry($entryName)) "Workbook package contains $entryName"
        }

        $workbookXml = Get-ZipXmlDocument $archive 'xl/workbook.xml'
        $workbookNs = New-SpreadsheetNamespaceManager $workbookXml
        $sheet = $workbookXml.SelectSingleNode('/s:workbook/s:sheets/s:sheet', $workbookNs)
        Assert-Equal 'Results' $sheet.GetAttribute('name') 'Workbook sheet is named Results'
        Assert-Equal 'rId1' $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') 'Workbook sheet uses its worksheet relationship'

        $stylesXml = Get-ZipXmlDocument $archive 'xl/styles.xml'
        $stylesNs = New-SpreadsheetNamespaceManager $stylesXml
        Assert-True ($null -ne $stylesXml.SelectSingleNode('/s:styleSheet/s:fonts/s:font[2]/s:b', $stylesNs)) 'Header style uses a bold font'
        Assert-Equal '1' $stylesXml.SelectSingleNode('/s:styleSheet/s:cellXfs/s:xf[2]', $stylesNs).GetAttribute('fontId') 'Header style index 1 selects the bold font'
        Assert-Equal '22' $stylesXml.SelectSingleNode('/s:styleSheet/s:cellXfs/s:xf[3]', $stylesNs).GetAttribute('numFmtId') 'Date style index 2 uses a date/time number format'

        $sheetXml = Get-ZipXmlDocument $archive 'xl/worksheets/sheet1.xml'
        $sheetNs = New-SpreadsheetNamespaceManager $sheetXml
        $pane = $sheetXml.SelectSingleNode('/s:worksheet/s:sheetViews/s:sheetView/s:pane', $sheetNs)
        Assert-Equal '1' $pane.GetAttribute('ySplit') 'Worksheet freezes one header row'
        Assert-Equal 'A2' $pane.GetAttribute('topLeftCell') 'Worksheet frozen pane starts at A2'
        Assert-Equal 'frozen' $pane.GetAttribute('state') 'Worksheet pane is frozen'
        Assert-Equal 'A1:G5' $sheetXml.SelectSingleNode('/s:worksheet/s:autoFilter', $sheetNs).GetAttribute('ref') 'Auto-filter covers the header and all mixed data rows'

        $headerCells = @($sheetXml.SelectNodes('/s:worksheet/s:sheetData/s:row[@r="1"]/s:c', $sheetNs))
        Assert-Equal 7 $headerCells.Count 'Worksheet writes every header cell'
        foreach ($headerCell in $headerCells) {
            Assert-Equal '1' $headerCell.GetAttribute('s') "Header cell $($headerCell.GetAttribute('r')) uses bold style index 1"
            Assert-Equal 'inlineStr' $headerCell.GetAttribute('t') "Header cell $($headerCell.GetAttribute('r')) is inline text"
        }

        foreach ($cellAddress in @('A2', 'B2')) {
            $cell = $sheetXml.SelectSingleNode("/s:worksheet/s:sheetData/s:row/s:c[@r='$cellAddress']", $sheetNs)
            Assert-True ($cell.GetAttribute('t') -ne 'inlineStr') "$cellAddress is stored as a number, not inline text"
        }
        $booleanCell = $sheetXml.SelectSingleNode('/s:worksheet/s:sheetData/s:row/s:c[@r="C2"]', $sheetNs)
        Assert-Equal 'b' $booleanCell.GetAttribute('t') 'Boolean cell uses the Boolean type'
        Assert-Equal '1' $booleanCell.SelectSingleNode('s:v', $sheetNs).InnerText 'True Boolean is stored as 1'
        $dateCell = $sheetXml.SelectSingleNode('/s:worksheet/s:sheetData/s:row/s:c[@r="D2"]', $sheetNs)
        Assert-Equal '2' $dateCell.GetAttribute('s') 'Date cell uses date style index 2'
        Assert-Equal $date.ToOADate().ToString('R', [Globalization.CultureInfo]::InvariantCulture) $dateCell.SelectSingleNode('s:v', $sheetNs).InnerText 'Date cell uses an invariant OLE Automation number'

        $normalTextCell = $sheetXml.SelectSingleNode('/s:worksheet/s:sheetData/s:row/s:c[@r="E2"]', $sheetNs)
        Assert-Equal 'inlineStr' $normalTextCell.GetAttribute('t') 'Normal text is stored inline'
        Assert-Equal ' normal text ' $normalTextCell.SelectSingleNode('s:is/s:t', $sheetNs).InnerText 'Inline text preserves leading and trailing spaces'
        Assert-Equal 'preserve' $normalTextCell.SelectSingleNode('s:is/s:t', $sheetNs).GetAttribute('space', 'http://www.w3.org/XML/1998/namespace') 'Inline text requests whitespace preservation'

        $formulaExpectations = @{ F2 = '=1+1'; F3 = '+cmd'; F4 = '-2+3'; F5 = '@name' }
        foreach ($cellAddress in $formulaExpectations.Keys) {
            $cell = $sheetXml.SelectSingleNode("/s:worksheet/s:sheetData/s:row/s:c[@r='$cellAddress']", $sheetNs)
            Assert-Equal 'inlineStr' $cell.GetAttribute('t') "Formula-looking $cellAddress is stored as inline text"
            Assert-Equal $formulaExpectations[$cellAddress] $cell.SelectSingleNode('s:is/s:t', $sheetNs).InnerText "Formula-looking $cellAddress keeps its literal value"
        }
        Assert-Equal 0 @($sheetXml.SelectNodes('//s:f', $sheetNs)).Count 'Worksheet never emits formula elements'

        $nullCell = $sheetXml.SelectSingleNode('/s:worksheet/s:sheetData/s:row/s:c[@r="G2"]', $sheetNs)
        Assert-True ($null -ne $nullCell) 'Database null emits a cell element'
        Assert-Equal 0 $nullCell.ChildNodes.Count 'Database null emits an empty cell'
        Assert-Equal '' $nullCell.GetAttribute('t') 'Database null has no cell type'
    }
    finally {
        $archive.Dispose()
    }

    $emptyTable = [System.Data.DataTable]::new('Empty')
    [void] $emptyTable.Columns.Add('Id', [int])
    [void] $emptyTable.Columns.Add('Name', [string])
    $emptyPath = Join-Path $testRoot 'empty.xlsx'
    Export-SqlUtilityXlsx -DestinationPath $emptyPath -RowSource (New-SqlUtilityDataTableRowSource $emptyTable) -TimeoutSeconds 30
    $emptyArchive = [System.IO.Compression.ZipFile]::OpenRead($emptyPath)
    try {
        $emptyXml = Get-ZipXmlDocument $emptyArchive 'xl/worksheets/sheet1.xml'
        $emptyNs = New-SpreadsheetNamespaceManager $emptyXml
        Assert-Equal 1 @($emptyXml.SelectNodes('/s:worksheet/s:sheetData/s:row', $emptyNs)).Count 'Zero-row export contains only the header row'
        Assert-Equal 'A1:B1' $emptyXml.SelectSingleNode('/s:worksheet/s:autoFilter', $emptyNs).GetAttribute('ref') 'Zero-row auto-filter covers the headers'
    }
    finally {
        $emptyArchive.Dispose()
    }

    foreach ($validMaximum in @(0, 1048575)) {
        $maximumBoundaryPath = Join-Path $testRoot "maximum-$validMaximum.xlsx"
        Export-SqlUtilityXlsx -DestinationPath $maximumBoundaryPath `
            -RowSource (New-SqlUtilityDataTableRowSource $emptyTable) -TimeoutSeconds 30 `
            -MaximumDataRows $validMaximum
        Assert-True (Test-Path -LiteralPath $maximumBoundaryPath -PathType Leaf) `
            "MaximumDataRows accepts the Excel data-row boundary $validMaximum"
    }

    foreach ($invalidMaximum in @(-1, 1048576)) {
        $invalidMaximumPath = Join-Path $testRoot "invalid-maximum-$invalidMaximum.xlsx"
        Assert-Throws {
            Export-SqlUtilityXlsx -DestinationPath $invalidMaximumPath `
                -RowSource (New-SqlUtilityDataTableRowSource $emptyTable) -TimeoutSeconds 30 `
                -MaximumDataRows $invalidMaximum
        } 'System.Management.Automation.ParameterBindingValidationException' `
            "MaximumDataRows rejects the out-of-range value $invalidMaximum before package creation"
        Assert-True (-not (Test-Path -LiteralPath $invalidMaximumPath)) `
            "Rejected MaximumDataRows $invalidMaximum leaves no destination"
        Assert-NoTemporaryFiles $testRoot "Rejected MaximumDataRows $invalidMaximum leaves no temporary file"
    }

    Assert-Equal 'XFD' (ConvertTo-SqlUtilityExcelColumnName 16384) 'Excel column boundary 16384 converts to XFD'

    $schemaColumn = [pscustomobject]@{ Name = 'Value'; DataType = [string]; Ordinal = 0 }
    $tooManyColumns = [object[]]::new(16385)
    for ($columnIndex = 0; $columnIndex -lt $tooManyColumns.Count; $columnIndex++) {
        $tooManyColumns[$columnIndex] = $schemaColumn
    }
    $tooManyColumnsSource = {
        param($OnSchema, $OnRow, $ShouldContinue)
        $null = & $OnSchema $tooManyColumns
    }
    $tooManyColumnsPath = Join-Path $testRoot 'too-many-columns.xlsx'
    Assert-Throws {
        Export-SqlUtilityXlsx -DestinationPath $tooManyColumnsPath -RowSource $tooManyColumnsSource -TimeoutSeconds 30
    } 'System.InvalidOperationException' 'A 16385-column schema is rejected before an XFE header can be written'
    Assert-True (-not (Test-Path -LiteralPath $tooManyColumnsPath)) 'Rejected 16385-column schema leaves no destination'
    Assert-NoTemporaryFiles $testRoot 'Rejected 16385-column schema deletes its temporary package'

    foreach ($boundary in @(
        @{ Count = 26; Expected = 'A1:Z1' },
        @{ Count = 27; Expected = 'A1:AA1' }
    )) {
        $boundaryPath = Join-Path $testRoot "columns-$($boundary.Count).xlsx"
        Export-SqlUtilityXlsx -DestinationPath $boundaryPath `
            -RowSource (New-SqlUtilityDataTableRowSource (New-ColumnBoundaryTable $boundary.Count)) -TimeoutSeconds 30
        $boundaryArchive = [System.IO.Compression.ZipFile]::OpenRead($boundaryPath)
        try {
            $boundaryXml = Get-ZipXmlDocument $boundaryArchive 'xl/worksheets/sheet1.xml'
            $boundaryNs = New-SpreadsheetNamespaceManager $boundaryXml
            Assert-Equal $boundary.Expected $boundaryXml.SelectSingleNode('/s:worksheet/s:autoFilter', $boundaryNs).GetAttribute('ref') `
                "$($boundary.Count)-column export uses the correct final Excel column address"
        }
        finally {
            $boundaryArchive.Dispose()
        }
    }

    $priorPath = Join-Path $testRoot 'preserve.xlsx'
    $priorBytes = [byte[]]@(80, 82, 73, 79, 82)
    [System.IO.File]::WriteAllBytes($priorPath, $priorBytes)
    $timeoutSource = {
        param($OnSchema, $OnRow, $ShouldContinue)
        & $OnSchema (, @([pscustomobject]@{ Name = 'Id'; DataType = [int]; Ordinal = 0 }))
        Start-Sleep -Milliseconds 1100
        & $OnRow (, [object[]]@(1))
    }
    Assert-Throws {
        Export-SqlUtilityXlsx -DestinationPath $priorPath -RowSource $timeoutSource -TimeoutSeconds 1
    } 'System.TimeoutException' 'Delayed row source times out before its row is written'
    Assert-Equal ($priorBytes -join ',') ([System.IO.File]::ReadAllBytes($priorPath) -join ',') 'Timeout preserves the prior destination byte-for-byte'
    Assert-NoTemporaryFiles $testRoot 'Timeout deletes its temporary sibling'

    $overflowTable = [System.Data.DataTable]::new('Overflow')
    [void] $overflowTable.Columns.Add('Id', [int])
    [void] $overflowTable.Rows.Add(1)
    [void] $overflowTable.Rows.Add(2)
    [void] $overflowTable.Rows.Add(3)
    $overflowPath = Join-Path $testRoot 'overflow.xlsx'
    [System.IO.File]::WriteAllBytes($overflowPath, $priorBytes)
    Assert-Throws {
        Export-SqlUtilityXlsx -DestinationPath $overflowPath `
            -RowSource (New-SqlUtilityDataTableRowSource $overflowTable) -TimeoutSeconds 30 -MaximumDataRows 2
    } 'System.InvalidOperationException' 'MaximumDataRows rejects the first row beyond the configured limit'
    Assert-Equal ($priorBytes -join ',') ([System.IO.File]::ReadAllBytes($overflowPath) -join ',') 'Row overflow preserves the prior destination byte-for-byte'
    Assert-NoTemporaryFiles $testRoot 'Row overflow deletes its temporary sibling'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

Complete-TestFile 'All Excel export tests passed.'
