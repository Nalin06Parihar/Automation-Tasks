<#
.SYNOPSIS
    Converts records from an Excel workbook into XML.

.EXAMPLE
    .\ExcelToXml.ps1 -InputFile ".\employees.xlsx" -OutputFile ".\employees.xml"

.EXAMPLE
    .\ExcelToXml.ps1 -InputFile ".\employees.xlsx" -OutputFile ".\employees.xml" `
        -SheetName "Employees" -RootElement "Employees" -RecordElement "Employee"

.NOTES
    Requires Microsoft Excel to be installed because the script uses Excel COM automation.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [string]$SheetName = "",

    [string]$RootElement = "Records",

    [string]$RecordElement = "Record"
)

$ErrorActionPreference = 'Stop'

$excel = $null
$workbook = $null
$worksheet = $null

try {
    $resolvedInputFile = (Resolve-Path -Path $InputFile -ErrorAction Stop).Path

    $resolvedOutputFile = [System.IO.Path]::GetFullPath($OutputFile)
    $outputDirectory = Split-Path -Path $resolvedOutputFile -Parent

    if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    Write-Host "Reading Excel file: $resolvedInputFile"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($resolvedInputFile)

    if ([string]::IsNullOrWhiteSpace($SheetName)) {
        $worksheet = $workbook.Worksheets.Item(1)
    }
    else {
        $worksheet = $workbook.Worksheets.Item($SheetName)
    }

    $usedRange = $worksheet.UsedRange

    if ($null -eq $usedRange -or $usedRange.Cells.Count -eq 0) {
        throw "The worksheet does not contain any data."
    }

    $rowCount = $usedRange.Rows.Count
    $columnCount = $usedRange.Columns.Count

    if ($rowCount -lt 2) {
        throw "The worksheet does not contain any data rows."
    }

    $headers = @()

    for ($column = 1; $column -le $columnCount; $column++) {
        $header = [string]$worksheet.Cells.Item(1, $column).Text

        if ([string]::IsNullOrWhiteSpace($header)) {
            throw "Column $column has an empty header."
        }

        $header = $header.Trim() -replace '[^a-zA-Z0-9_.-]', '_'

        if ($header -match '^[0-9]') {
            $header = "_$header"
        }

        $headers += $header
    }

    $duplicates = $headers |
        Group-Object |
        Where-Object { $_.Count -gt 1 }

    if ($duplicates) {
        $duplicateNames = ($duplicates.Name -join ', ')
        throw "Duplicate column names detected: $duplicateNames"
    }

    $xmlDocument = New-Object System.Xml.XmlDocument

    $declaration = $xmlDocument.CreateXmlDeclaration('1.0', 'UTF-8', $null)
    $xmlDocument.AppendChild($declaration) | Out-Null

    $root = $xmlDocument.CreateElement($RootElement)
    $xmlDocument.AppendChild($root) | Out-Null

    for ($row = 2; $row -le $rowCount; $row++) {
        $record = $xmlDocument.CreateElement($RecordElement)

        for ($column = 1; $column -le $columnCount; $column++) {
            $elementName = $headers[$column - 1]
            $value = [string]$worksheet.Cells.Item($row, $column).Text

            $element = $xmlDocument.CreateElement($elementName)
            $element.InnerText = $value

            $record.AppendChild($element) | Out-Null
        }

        $root.AppendChild($record) | Out-Null
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)

    $writer = [System.Xml.XmlWriter]::Create($resolvedOutputFile, $settings)

    try {
        $xmlDocument.Save($writer)
    }
    finally {
        $writer.Close()
    }

    Write-Host ""
    Write-Host "Conversion completed successfully." -ForegroundColor Green
    Write-Host "Records converted: $($rowCount - 1)"
    Write-Host "Output: $resolvedOutputFile"
}
catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
    }

    if ($excel -ne $null) {
        $excel.Quit()
    }

    if ($worksheet -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    }

    if ($workbook -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($excel -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}