Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:SqlUtilitySpreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:SqlUtilityPackageRelationshipsNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:SqlUtilityOfficeRelationshipsNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:SqlUtilityXmlNamespace = 'http://www.w3.org/XML/1998/namespace'

function New-SqlUtilityDataTableRowSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.DataTable] $DataTable
    )

    $rowSource = {
        param($OnSchema, $OnRow, $ShouldContinue)

        $schema = @(
            foreach ($column in $DataTable.Columns) {
                [pscustomobject]@{
                    Name = $column.ColumnName
                    DataType = $column.DataType
                    Ordinal = $column.Ordinal
                }
            }
        )
        $null = & $OnSchema $schema

        foreach ($row in $DataTable.Rows) {
            if (-not (& $ShouldContinue)) { break }
            $values = [object[]] $row.ItemArray
            $null = & $OnRow $values
        }
    }

    return $rowSource.GetNewClosure()
}

function ConvertTo-SqlUtilityExcelColumnName {
    param([Parameter(Mandatory = $true)][int] $ColumnNumber)

    if ($ColumnNumber -lt 1) {
        throw [System.ArgumentOutOfRangeException]::new('ColumnNumber')
    }

    $name = ''
    $remaining = $ColumnNumber
    while ($remaining -gt 0) {
        $remaining--
        $name = [char](65 + ($remaining % 26)) + $name
        $remaining = [int] [Math]::Floor($remaining / 26)
    }
    return $name
}

function New-SqlUtilityArchiveXmlWriter {
    param(
        [Parameter(Mandatory = $true)][System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory = $true)][string] $EntryName
    )

    $entry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $settings = [System.Xml.XmlWriterSettings]::new()
        $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
        $settings.Indent = $false
        $settings.CloseOutput = $true
        return [System.Xml.XmlWriter]::Create($stream, $settings)
    }
    catch {
        $stream.Dispose()
        throw
    }
}

function Write-SqlUtilityContentTypes {
    param([System.IO.Compression.ZipArchive] $Archive)

    $writer = New-SqlUtilityArchiveXmlWriter $Archive '[Content_Types].xml'
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('Types', 'http://schemas.openxmlformats.org/package/2006/content-types')

        $writer.WriteStartElement('Default')
        $writer.WriteAttributeString('Extension', 'rels')
        $writer.WriteAttributeString('ContentType', 'application/vnd.openxmlformats-package.relationships+xml')
        $writer.WriteEndElement()

        $writer.WriteStartElement('Default')
        $writer.WriteAttributeString('Extension', 'xml')
        $writer.WriteAttributeString('ContentType', 'application/xml')
        $writer.WriteEndElement()

        foreach ($override in @(
            @('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
            @('/xl/worksheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            @('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
        )) {
            $writer.WriteStartElement('Override')
            $writer.WriteAttributeString('PartName', $override[0])
            $writer.WriteAttributeString('ContentType', $override[1])
            $writer.WriteEndElement()
        }

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SqlUtilityPackageRelationships {
    param([System.IO.Compression.ZipArchive] $Archive)

    $writer = New-SqlUtilityArchiveXmlWriter $Archive '_rels/.rels'
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('Relationships', $script:SqlUtilityPackageRelationshipsNamespace)
        $writer.WriteStartElement('Relationship')
        $writer.WriteAttributeString('Id', 'rId1')
        $writer.WriteAttributeString('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
        $writer.WriteAttributeString('Target', 'xl/workbook.xml')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SqlUtilityWorkbook {
    param([System.IO.Compression.ZipArchive] $Archive)

    $writer = New-SqlUtilityArchiveXmlWriter $Archive 'xl/workbook.xml'
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('workbook', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('xmlns', 'r', $null, $script:SqlUtilityOfficeRelationshipsNamespace)
        $writer.WriteStartElement('sheets', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteStartElement('sheet', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('name', 'Results')
        $writer.WriteAttributeString('sheetId', '1')
        $writer.WriteAttributeString('r', 'id', $script:SqlUtilityOfficeRelationshipsNamespace, 'rId1')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SqlUtilityWorkbookRelationships {
    param([System.IO.Compression.ZipArchive] $Archive)

    $writer = New-SqlUtilityArchiveXmlWriter $Archive 'xl/_rels/workbook.xml.rels'
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('Relationships', $script:SqlUtilityPackageRelationshipsNamespace)

        $writer.WriteStartElement('Relationship')
        $writer.WriteAttributeString('Id', 'rId1')
        $writer.WriteAttributeString('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        $writer.WriteAttributeString('Target', 'worksheets/sheet1.xml')
        $writer.WriteEndElement()

        $writer.WriteStartElement('Relationship')
        $writer.WriteAttributeString('Id', 'rId2')
        $writer.WriteAttributeString('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles')
        $writer.WriteAttributeString('Target', 'styles.xml')
        $writer.WriteEndElement()

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SqlUtilityStyles {
    param([System.IO.Compression.ZipArchive] $Archive)

    $writer = New-SqlUtilityArchiveXmlWriter $Archive 'xl/styles.xml'
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('styleSheet', $script:SqlUtilitySpreadsheetNamespace)

        $writer.WriteStartElement('fonts', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '2')
        foreach ($bold in @($false, $true)) {
            $writer.WriteStartElement('font', $script:SqlUtilitySpreadsheetNamespace)
            if ($bold) { $writer.WriteElementString('b', $script:SqlUtilitySpreadsheetNamespace, '') }
            $writer.WriteStartElement('sz', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('val', '11')
            $writer.WriteEndElement()
            $writer.WriteStartElement('name', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('val', 'Calibri')
            $writer.WriteEndElement()
            $writer.WriteStartElement('family', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('val', '2')
            $writer.WriteEndElement()
            $writer.WriteStartElement('scheme', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('val', 'minor')
            $writer.WriteEndElement()
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('fills', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '2')
        foreach ($pattern in @('none', 'gray125')) {
            $writer.WriteStartElement('fill', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteStartElement('patternFill', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('patternType', $pattern)
            $writer.WriteEndElement()
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('borders', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '1')
        $writer.WriteStartElement('border', $script:SqlUtilitySpreadsheetNamespace)
        foreach ($side in @('left', 'right', 'top', 'bottom', 'diagonal')) {
            $writer.WriteElementString($side, $script:SqlUtilitySpreadsheetNamespace, '')
        }
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        $writer.WriteStartElement('cellStyleXfs', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '1')
        $writer.WriteStartElement('xf', $script:SqlUtilitySpreadsheetNamespace)
        foreach ($attribute in @{ numFmtId = '0'; fontId = '0'; fillId = '0'; borderId = '0' }.GetEnumerator()) {
            $writer.WriteAttributeString($attribute.Key, $attribute.Value)
        }
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        $writer.WriteStartElement('cellXfs', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '3')
        foreach ($format in @(
            @{ numFmtId = '0'; fontId = '0'; applyFont = $null; applyNumberFormat = $null },
            @{ numFmtId = '0'; fontId = '1'; applyFont = '1'; applyNumberFormat = $null },
            @{ numFmtId = '22'; fontId = '0'; applyFont = $null; applyNumberFormat = '1' }
        )) {
            $writer.WriteStartElement('xf', $script:SqlUtilitySpreadsheetNamespace)
            $writer.WriteAttributeString('numFmtId', $format.numFmtId)
            $writer.WriteAttributeString('fontId', $format.fontId)
            $writer.WriteAttributeString('fillId', '0')
            $writer.WriteAttributeString('borderId', '0')
            $writer.WriteAttributeString('xfId', '0')
            if ($null -ne $format.applyFont) { $writer.WriteAttributeString('applyFont', $format.applyFont) }
            if ($null -ne $format.applyNumberFormat) { $writer.WriteAttributeString('applyNumberFormat', $format.applyNumberFormat) }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('cellStyles', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '1')
        $writer.WriteStartElement('cellStyle', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('name', 'Normal')
        $writer.WriteAttributeString('xfId', '0')
        $writer.WriteAttributeString('builtinId', '0')
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        $writer.WriteStartElement('dxfs', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '0')
        $writer.WriteEndElement()
        $writer.WriteStartElement('tableStyles', $script:SqlUtilitySpreadsheetNamespace)
        $writer.WriteAttributeString('count', '0')
        $writer.WriteAttributeString('defaultTableStyle', 'TableStyleMedium2')
        $writer.WriteAttributeString('defaultPivotStyle', 'PivotStyleLight16')
        $writer.WriteEndElement()

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SqlUtilityInlineStringCell {
    param(
        [System.Xml.XmlWriter] $Writer,
        [string] $CellReference,
        [string] $Value,
        [int] $StyleIndex = 0
    )

    $Writer.WriteStartElement('c', $script:SqlUtilitySpreadsheetNamespace)
    $Writer.WriteAttributeString('r', $CellReference)
    if ($StyleIndex -gt 0) { $Writer.WriteAttributeString('s', $StyleIndex.ToString([Globalization.CultureInfo]::InvariantCulture)) }
    $Writer.WriteAttributeString('t', 'inlineStr')
    $Writer.WriteStartElement('is', $script:SqlUtilitySpreadsheetNamespace)
    $Writer.WriteStartElement('t', $script:SqlUtilitySpreadsheetNamespace)
    $Writer.WriteAttributeString('xml', 'space', $script:SqlUtilityXmlNamespace, 'preserve')
    $Writer.WriteString($Value)
    $Writer.WriteEndElement()
    $Writer.WriteEndElement()
    $Writer.WriteEndElement()
}

function Write-SqlUtilityDataCell {
    param(
        [System.Xml.XmlWriter] $Writer,
        [string] $CellReference,
        $Value
    )

    if ($null -eq $Value -or [DBNull]::Value.Equals($Value)) {
        $Writer.WriteStartElement('c', $script:SqlUtilitySpreadsheetNamespace)
        $Writer.WriteAttributeString('r', $CellReference)
        $Writer.WriteEndElement()
        return
    }

    if ($Value -is [bool]) {
        $Writer.WriteStartElement('c', $script:SqlUtilitySpreadsheetNamespace)
        $Writer.WriteAttributeString('r', $CellReference)
        $Writer.WriteAttributeString('t', 'b')
        $Writer.WriteElementString('v', $script:SqlUtilitySpreadsheetNamespace, $(if ($Value) { '1' } else { '0' }))
        $Writer.WriteEndElement()
        return
    }

    if ($Value -is [datetime]) {
        $Writer.WriteStartElement('c', $script:SqlUtilitySpreadsheetNamespace)
        $Writer.WriteAttributeString('r', $CellReference)
        $Writer.WriteAttributeString('s', '2')
        $number = $Value.ToOADate().ToString('R', [Globalization.CultureInfo]::InvariantCulture)
        $Writer.WriteElementString('v', $script:SqlUtilitySpreadsheetNamespace, $number)
        $Writer.WriteEndElement()
        return
    }

    $typeCode = [System.Type]::GetTypeCode($Value.GetType())
    if ($typeCode -ge [System.TypeCode]::SByte -and $typeCode -le [System.TypeCode]::Decimal) {
        $Writer.WriteStartElement('c', $script:SqlUtilitySpreadsheetNamespace)
        $Writer.WriteAttributeString('r', $CellReference)
        $number = ([System.IFormattable] $Value).ToString($null, [Globalization.CultureInfo]::InvariantCulture)
        $Writer.WriteElementString('v', $script:SqlUtilitySpreadsheetNamespace, $number)
        $Writer.WriteEndElement()
        return
    }

    $text = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    Write-SqlUtilityInlineStringCell -Writer $Writer -CellReference $CellReference -Value $text
}

function Export-SqlUtilityXlsx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $DestinationPath,
        [Parameter(Mandatory = $true)][scriptblock] $RowSource,
        [Parameter(Mandatory = $true)][int] $TimeoutSeconds,
        [ValidateRange(0, 1048575)][int] $MaximumDataRows = 1048575
    )

    $destinationFullPath = [System.IO.Path]::GetFullPath($DestinationPath)
    $destinationDirectory = [System.IO.Path]::GetDirectoryName($destinationFullPath)
    $destinationFileName = [System.IO.Path]::GetFileName($destinationFullPath)
    $temporaryPath = Join-Path $destinationDirectory ($destinationFileName + '.' + [guid]::NewGuid().ToString('N') + '.tmp')

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $archive = $null
    $worksheetWriter = $null
    $primaryError = $null
    $cleanupError = $null
    $state = [pscustomobject]@{
        SchemaWritten = $false
        Columns = @()
        DataRowCount = 0
    }
    $convertColumnName = ${function:ConvertTo-SqlUtilityExcelColumnName}
    $writeInlineStringCell = ${function:Write-SqlUtilityInlineStringCell}
    $writeDataCell = ${function:Write-SqlUtilityDataCell}
    $spreadsheetNamespace = $script:SqlUtilitySpreadsheetNamespace
    $shouldContinue = {
        if ($stopwatch.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
            throw [System.TimeoutException]::new("Excel export exceeded $TimeoutSeconds seconds.")
        }
        return $true
    }.GetNewClosure()

    try {
        $archive = [System.IO.Compression.ZipFile]::Open($temporaryPath, [System.IO.Compression.ZipArchiveMode]::Create)
        Write-SqlUtilityContentTypes $archive
        Write-SqlUtilityPackageRelationships $archive
        Write-SqlUtilityWorkbook $archive
        Write-SqlUtilityWorkbookRelationships $archive
        Write-SqlUtilityStyles $archive

        $worksheetWriter = New-SqlUtilityArchiveXmlWriter $archive 'xl/worksheets/sheet1.xml'
        $worksheetWriter.WriteStartDocument()
        $worksheetWriter.WriteStartElement('worksheet', $script:SqlUtilitySpreadsheetNamespace)
        $worksheetWriter.WriteStartElement('sheetViews', $script:SqlUtilitySpreadsheetNamespace)
        $worksheetWriter.WriteStartElement('sheetView', $script:SqlUtilitySpreadsheetNamespace)
        $worksheetWriter.WriteAttributeString('workbookViewId', '0')
        $worksheetWriter.WriteStartElement('pane', $script:SqlUtilitySpreadsheetNamespace)
        $worksheetWriter.WriteAttributeString('ySplit', '1')
        $worksheetWriter.WriteAttributeString('topLeftCell', 'A2')
        $worksheetWriter.WriteAttributeString('activePane', 'bottomLeft')
        $worksheetWriter.WriteAttributeString('state', 'frozen')
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteStartElement('sheetFormatPr', $script:SqlUtilitySpreadsheetNamespace)
        $worksheetWriter.WriteAttributeString('defaultRowHeight', '15')
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteStartElement('sheetData', $script:SqlUtilitySpreadsheetNamespace)

        $onSchema = {
            param($Columns)
            if ($state.SchemaWritten) {
                throw [System.InvalidOperationException]::new('The row source invoked the schema callback more than once.')
            }

            $state.Columns = @($Columns)
            if ($state.Columns.Count -lt 1) {
                throw [System.InvalidOperationException]::new('The row source schema must contain at least one column.')
            }
            if ($state.Columns.Count -gt 16384) {
                throw [System.InvalidOperationException]::new('The row source schema exceeds the Excel limit of 16384 columns.')
            }
            $state.SchemaWritten = $true

            $worksheetWriter.WriteStartElement('row', $spreadsheetNamespace)
            $worksheetWriter.WriteAttributeString('r', '1')
            for ($columnIndex = 0; $columnIndex -lt $state.Columns.Count; $columnIndex++) {
                $null = & $shouldContinue
                $cellReference = (& $convertColumnName ($columnIndex + 1)) + '1'
                & $writeInlineStringCell -Writer $worksheetWriter -CellReference $cellReference `
                    -Value ([Convert]::ToString($state.Columns[$columnIndex].Name)) -StyleIndex 1
            }
            $worksheetWriter.WriteEndElement()
        }.GetNewClosure()

        $onRow = {
            param($Values)
            if (-not $state.SchemaWritten) {
                throw [System.InvalidOperationException]::new('The row source invoked a row callback before the schema callback.')
            }

            $null = & $shouldContinue
            $state.DataRowCount++
            if ($state.DataRowCount -gt $MaximumDataRows) {
                throw [System.InvalidOperationException]::new("Excel export exceeded the maximum of $MaximumDataRows data rows.")
            }

            $rowNumber = $state.DataRowCount + 1
            $worksheetWriter.WriteStartElement('row', $spreadsheetNamespace)
            $worksheetWriter.WriteAttributeString('r', $rowNumber.ToString([Globalization.CultureInfo]::InvariantCulture))
            $rowValues = @($Values)
            for ($columnIndex = 0; $columnIndex -lt $state.Columns.Count; $columnIndex++) {
                $cellReference = (& $convertColumnName ($columnIndex + 1)) + $rowNumber
                $value = if ($columnIndex -lt $rowValues.Count) { $rowValues[$columnIndex] } else { [DBNull]::Value }
                & $writeDataCell -Writer $worksheetWriter -CellReference $cellReference -Value $value
            }
            $worksheetWriter.WriteEndElement()
        }.GetNewClosure()

        $null = & $RowSource $onSchema $onRow $shouldContinue
        if (-not $state.SchemaWritten) {
            throw [System.InvalidOperationException]::new('The row source did not invoke the schema callback.')
        }

        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteStartElement('autoFilter', $script:SqlUtilitySpreadsheetNamespace)
        $lastColumn = ConvertTo-SqlUtilityExcelColumnName $state.Columns.Count
        $lastRow = $state.DataRowCount + 1
        $worksheetWriter.WriteAttributeString('ref', "A1:$lastColumn$lastRow")
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteEndElement()
        $worksheetWriter.WriteEndDocument()
        $worksheetWriter.Dispose()
        $worksheetWriter = $null

        $archive.Dispose()
        $archive = $null

        $null = & $shouldContinue
        if ([System.IO.File]::Exists($destinationFullPath)) {
            [System.IO.File]::Replace($temporaryPath, $destinationFullPath, $null)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $destinationFullPath)
        }
    }
    catch {
        $primaryError = $_
    }
    finally {
        try {
            if ($null -ne $worksheetWriter) { $worksheetWriter.Dispose() }
        }
        catch {
            if ($null -eq $cleanupError) { $cleanupError = $_ }
        }
        try {
            if ($null -ne $archive) { $archive.Dispose() }
        }
        catch {
            if ($null -eq $cleanupError) { $cleanupError = $_ }
        }
        try {
            if ([System.IO.File]::Exists($temporaryPath)) {
                [System.IO.File]::Delete($temporaryPath)
            }
        }
        catch {
            if ($null -eq $cleanupError) { $cleanupError = $_ }
        }
        try {
            $stopwatch.Stop()
        }
        catch {
            if ($null -eq $cleanupError) { $cleanupError = $_ }
        }
    }

    if ($null -ne $primaryError) {
        $PSCmdlet.ThrowTerminatingError($primaryError)
    }
    if ($null -ne $cleanupError) {
        $PSCmdlet.ThrowTerminatingError($cleanupError)
    }
}
