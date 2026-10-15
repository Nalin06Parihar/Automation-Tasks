[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$OutputFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$excel = $null
$workbook = $null
$release = @()

function Read-Num {
    param(
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )

    while ($true) {
        $raw = Read-Host $Prompt
        $value = 0

        if ([int]::TryParse([string]$raw, [ref]$value) -and $value -ge $Min -and $value -le $Max) {
            return $value
        }

        Write-Host "Enter $Min to $Max." -ForegroundColor Yellow
    }
}

function Read-YN {
    param(
        [string]$Prompt,
        [bool]$Default = $false
    )

    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }

    while ($true) {
        $answer = ([string](Read-Host "$Prompt $hint")).Trim().ToLower()

        if (-not $answer) {
            return $Default
        }

        if ($answer -in @('y', 'yes')) {
            return $true
        }

        if ($answer -in @('n', 'no')) {
            return $false
        }

        Write-Host 'Enter Y or N.' -ForegroundColor Yellow
    }
}

function Safe-Name {
    param(
        [string]$Name,
        [string]$Fallback
    )

    $safe = ([string]$Name).Trim() -replace '[^A-Za-z0-9_.-]', '_'

    if (-not $safe) {
        $safe = $Fallback
    }

    if ($safe -match '^[0-9.-]') {
        $safe = '_' + $safe
    }

    try {
        [System.Xml.XmlConvert]::VerifyName($safe) | Out-Null
        return $safe
    }
    catch {
        return $Fallback
    }
}

function Record-Name {
    param(
        [string]$Name
    )

    $base = Safe-Name -Name $Name -Fallback 'Record'

    if ($base.EndsWith('ies') -and $base.Length -gt 3) {
        return $base.Substring(0, $base.Length - 3) + 'y'
    }

    if ($base.EndsWith('s') -and -not $base.EndsWith('ss') -and $base.Length -gt 1) {
        return $base.Substring(0, $base.Length - 1)
    }

    return $base + 'Record'
}

function Inspect-Sheet {
    param(
        $Worksheet
    )

    $used = $Worksheet.UsedRange
    $script:release += , $used

    $firstRow = [int]$used.Row
    $lastRow = $firstRow + [int]$used.Rows.Count - 1
    $firstCol = [int]$used.Column
    $lastCol = $firstCol + [int]$used.Columns.Count - 1

    $headerRow = $firstRow
    $bestCount = -1
    $scanTo = [Math]::Min($lastRow, $firstRow + 9)

    for ($row = $firstRow; $row -le $scanTo; $row++) {
        $count = 0

        for ($col = $firstCol; $col -le $lastCol; $col++) {
            if (-not [string]::IsNullOrWhiteSpace([string]$Worksheet.Cells.Item($row, $col).Text)) {
                $count++
            }
        }

        if ($count -gt $bestCount) {
            $bestCount = $count
            $headerRow = $row
        }
    }

    $headers = @()
    $columns = @()

    for ($col = $firstCol; $col -le $lastCol; $col++) {
        $header = ([string]$Worksheet.Cells.Item($headerRow, $col).Text).Trim()

        if ($header) {
            $baseName = $header
            $suffix = 2

            while ($headers -contains $header) {
                $header = $baseName + '_' + $suffix
                $suffix++
            }

            $headers += [string]$header
            $columns += [int]$col
        }
    }

    if ($headers.Count -eq 0) {
        throw "No headers detected in '$($Worksheet.Name)'."
    }

    return [pscustomobject]@{
        Name = [string]$Worksheet.Name
        Worksheet = $Worksheet
        HeaderRow = $headerRow
        LastRow = $lastRow
        Headers = @($headers)
        Columns = @($columns)
    }
}

function Read-Rows {
    param(
        $Info
    )

    $rows = @()

    for ($row = $Info.HeaderRow + 1; $row -le $Info.LastRow; $row++) {
        $values = @{}
        $hasContent = $false

        for ($index = 0; $index -lt $Info.Headers.Count; $index++) {
            $header = [string]$Info.Headers[$index]
            $column = [int]$Info.Columns[$index]
            $value = [string]$Info.Worksheet.Cells.Item($row, $column).Text

            if ($value.Trim()) {
                $hasContent = $true
            }

            $values[$header] = $value
        }

        if ($hasContent) {
            $rows += , [pscustomobject]@{
                Row = $row
                Values = $values
            }
        }
    }

    return $rows
}

function Select-Column {
    param(
        $Info,
        [string]$Prompt
    )

    Write-Host "`nColumns in '$($Info.Name)':" -ForegroundColor Cyan

    for ($i = 0; $i -lt $Info.Headers.Count; $i++) {
        Write-Host "[$($i + 1)] $($Info.Headers[$i])"
    }

    $choice = Read-Num -Prompt $Prompt -Min 1 -Max $Info.Headers.Count
    return [string]$Info.Headers[$choice - 1]
}

function Add-Fields {
    param(
        [System.Xml.XmlDocument]$Document,
        [System.Xml.XmlElement]$Parent,
        $Values,
        $Headers,
        [string]$Exclude = ''
    )

    foreach ($item in @($Headers)) {
        $header = [string]$item

        if ($Exclude -and $header -eq $Exclude) {
            continue
        }

        $element = $Document.CreateElement((Safe-Name -Name $header -Fallback 'Field'))
        $element.InnerText = [string]$Values[$header]
        [void]$Parent.AppendChild($element)
    }
}

try {
    Write-Host 'Opening workbook...' -ForegroundColor Cyan

    $InputFile = (Resolve-Path -LiteralPath $InputFile).Path
    $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)

    if ([IO.Path]::GetExtension($InputFile).ToLower() -notin @('.xlsx', '.xlsm', '.xls')) {
        throw 'Input must be an Excel workbook.'
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($InputFile, 0, $true)

    $sheets = @()

    for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
        $sheet = $workbook.Worksheets.Item($index)
        $script:release += , $sheet

        try {
            $sheets += , (Inspect-Sheet -Worksheet $sheet)
        }
        catch {
            Write-Host "Skipping '$($sheet.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($sheets.Count -eq 0) {
        throw 'No usable worksheets found.'
    }

    Write-Host "`nWorkbook inspection" -ForegroundColor Green

    for ($i = 0; $i -lt $sheets.Count; $i++) {
        Write-Host "[$($i + 1)] $($sheets[$i].Name)"
        Write-Host "    Header row: $($sheets[$i].HeaderRow)"
        Write-Host "    Columns: $($sheets[$i].Headers -join ', ')"
    }

    $mainChoice = Read-Num -Prompt "`nSelect the main sheet" -Min 1 -Max $sheets.Count
    $main = $sheets[$mainChoice - 1]
    $mainRows = @(Read-Rows -Info $main)

    if ($mainRows.Count -eq 0) {
        throw "'$($main.Name)' has no data rows."
    }

    $mainKey = Select-Column -Info $main -Prompt 'Select the main key column'
    $defaultRoot = Safe-Name -Name $main.Name -Fallback 'Records'
    $defaultRecord = Record-Name -Name $main.Name

    $rootInput = [string](Read-Host "Root XML element [$defaultRoot]")
    $rootName = if ($rootInput) { Safe-Name -Name $rootInput -Fallback $defaultRoot } else { $defaultRoot }

    $recordInput = [string](Read-Host "Main record XML element [$defaultRecord]")
    $recordName = if ($recordInput) { Safe-Name -Name $recordInput -Fallback $defaultRecord } else { $defaultRecord }

    $seen = @{}
    $hasBadData = $false

    foreach ($row in $mainRows) {
        $key = [string]$row.Values[$mainKey]

        if (-not $key -or $seen.ContainsKey($key)) {
            $hasBadData = $true
        }
        else {
            $seen[$key] = $true
        }
    }

    if ($hasBadData -and -not (Read-YN -Prompt "'$mainKey' has blank or duplicate values. Continue?")) {
        throw 'Cancelled.'
    }

    $relations = @()

    for ($i = 0; $i -lt $sheets.Count; $i++) {
        if ($i -eq ($mainChoice - 1)) {
            continue
        }

        $child = $sheets[$i]

        if (-not (Read-YN -Prompt "Link '$($child.Name)' as child records of '$($main.Name)'?")) {
            continue
        }

        $parentColumn = $mainKey

        if (Read-YN -Prompt 'Use a different parent matching column?') {
            $parentColumn = Select-Column -Info $main -Prompt 'Select parent matching column'
        }

        $childColumn = Select-Column -Info $child -Prompt 'Select child matching column'

        $defaultContainer = Safe-Name -Name $child.Name -Fallback 'Items'
        $defaultChildRecord = Record-Name -Name $child.Name

        $containerInput = [string](Read-Host "Child container XML element [$defaultContainer]")
        $containerName = if ($containerInput) { Safe-Name -Name $containerInput -Fallback $defaultContainer } else { $defaultContainer }

        $childRecordInput = [string](Read-Host "Child record XML element [$defaultChildRecord]")
        $childRecordName = if ($childRecordInput) { Safe-Name -Name $childRecordInput -Fallback $defaultChildRecord } else { $defaultChildRecord }

        $index = @{}

        foreach ($childRow in @(Read-Rows -Info $child)) {
            $key = [string]$childRow.Values[$childColumn]

            if (-not $index.ContainsKey($key)) {
                $index[$key] = @()
            }

            $index[$key] = @($index[$key]) + @($childRow)
        }

        $relations += , [pscustomobject]@{
            Sheet = $child
            ParentColumn = [string]$parentColumn
            ChildColumn = [string]$childColumn
            Container = [string]$containerName
            Record = [string]$childRecordName
            Index = $index
        }
    }

    Write-Host "`nConversion plan" -ForegroundColor Green
    Write-Host "Main: $($main.Name).$mainKey -> <$rootName><$recordName>"

    foreach ($relation in $relations) {
        Write-Host "Child: $($main.Name).$($relation.ParentColumn) = $($relation.Sheet.Name).$($relation.ChildColumn)"
    }

    if (-not (Read-YN -Prompt 'Generate XML with this plan?' -Default $true)) {
        throw 'Cancelled.'
    }

    $document = New-Object System.Xml.XmlDocument
    [void]$document.AppendChild($document.CreateXmlDeclaration('1.0', 'UTF-8', $null))

    $root = $document.CreateElement($rootName)
    [void]$document.AppendChild($root)

    $childCount = 0

    foreach ($row in $mainRows) {
        $node = $document.CreateElement($recordName)
        Add-Fields -Document $document -Parent $node -Values $row.Values -Headers $main.Headers

        foreach ($relation in $relations) {
            $key = [string]$row.Values[$relation.ParentColumn]

            if (-not $relation.Index.ContainsKey($key)) {
                continue
            }

            $container = $document.CreateElement($relation.Container)

            foreach ($childRow in @($relation.Index[$key])) {
                $childNode = $document.CreateElement($relation.Record)
                Add-Fields -Document $document -Parent $childNode -Values $childRow.Values -Headers $relation.Sheet.Headers -Exclude $relation.ChildColumn
                [void]$container.AppendChild($childNode)
                $childCount++
            }

            if ($container.HasChildNodes) {
                [void]$node.AppendChild($container)
            }
        }

        [void]$root.AppendChild($node)
    }

    $directory = Split-Path -Parent $OutputFile
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $writer = [System.Xml.XmlWriter]::Create($OutputFile, $settings)

    try {
        $document.Save($writer)
    }
    finally {
        $writer.Dispose()
    }

    Write-Host "`nConversion completed successfully." -ForegroundColor Green
    Write-Host "Main records: $($mainRows.Count)"
    Write-Host "Nested child records: $childCount"
    Write-Host "Output: $OutputFile"
}
catch {
    Write-Error "Line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    exit 1
}
finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch { }
    }

    if ($excel) {
        try { $excel.Quit() } catch { }
    }

    foreach ($obj in @($release)) {
        if ($obj) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
        }
    }

    if ($workbook) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch { }
    }

    if ($excel) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

