# Excel to XML Automation

This folder contains two PowerShell workflows for converting Excel worksheet
data into XML:

- [ExcelToXml.ps1](./ExcelToXml.ps1) is a parameter-driven converter for one
  worksheet. It is suited to repeatable, simple tabular exports.
- [InteractiveExcelToXml_Corrected.ps1](./InteractiveExcelToXml_Corrected.ps1)
  inspects a workbook and interactively lets you select a main sheet, identify
  child sheets, and produce nested XML records.

The sample workbooks `EmployeeRecords.xlsx` and `CustomersAndOrders.xlsx` can be
used to explore the workflows.

## Requirements

- Windows and PowerShell.
- Microsoft Excel desktop installed. Both scripts use Excel COM automation to
  read workbook values; they do not parse `.xlsx` files independently.
- A workbook that Excel can open and that contains data in the expected
  tabular layout.

Close a workbook in Excel before converting it if it is locked or prompting
for edits. Both scripts open Excel invisibly and close the workbook without
saving changes.

## Workflow A: convert one sheet

`ExcelToXml.ps1` is the simpler, non-interactive workflow:

1. Resolves the input workbook and output path.
2. Starts an invisible Excel application and opens the workbook.
3. Selects the named worksheet, or the first worksheet if no name was given.
4. Treats row 1 as the column header row and each following used row as a
   record.
5. Validates that the sheet has data rows, that every header is non-empty, and
   that sanitized header names are unique.
6. Creates a root XML element and one record element for each data row.
7. Adds one child element per column, using the header as its element name and
   the displayed cell text as its value.
8. Writes indented UTF-8 XML without a byte-order mark, then closes Excel and
   releases its COM objects.

Example using the default first worksheet:

```powershell
.\ExcelToXml.ps1 `
  -InputFile ".\EmployeeRecords.xlsx" `
  -OutputFile ".\output\employees.xml"
```

Choose a worksheet and XML element names explicitly:

```powershell
.\ExcelToXml.ps1 `
  -InputFile ".\EmployeeRecords.xlsx" `
  -OutputFile ".\output\employees.xml" `
  -SheetName "Employees" `
  -RootElement "Employees" `
  -RecordElement "Employee"
```

Parameters:

| Parameter | Required | Meaning |
|---|---:|---|
| `InputFile` | Yes | Path to the Excel workbook. |
| `OutputFile` | Yes | Path for the generated XML; missing output directories are created. |
| `SheetName` | No | Worksheet name; defaults to the first worksheet. |
| `RootElement` | No | XML document root name; defaults to `Records`. |
| `RecordElement` | No | Element name for each data row; defaults to `Record`. |

Headers are trimmed and characters outside letters, digits, underscore, dot,
and hyphen are replaced with underscores. Headers beginning with a digit are
prefixed with an underscore. The first script reads displayed cell text, so
number/date formatting in Excel can affect the resulting XML text. It does not
infer data types or preserve formulas as formulas.
The `RootElement` and `RecordElement` values must also be valid XML element
names; the simple script passes those names directly to the XML writer rather
than sanitizing them.

Example shape of the output:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Employees>
  <Employee>
    <EmployeeID>101</EmployeeID>
    <Name>Ada Example</Name>
  </Employee>
</Employees>
```

The element names and values above are illustrative; the real names and values
come from the selected worksheet's headers and cells.

## Workflow B: build nested XML from related worksheets

`InteractiveExcelToXml_Corrected.ps1` is intended for workbooks where one sheet
contains primary records and other sheets contain related rows, such as
customers and their orders.

Run it with the required paths:

```powershell
.\InteractiveExcelToXml_Corrected.ps1 `
  -InputFile ".\CustomersAndOrders.xlsx" `
  -OutputFile ".\output\customers-and-orders.xml"
```

The script then guides the conversion:

1. Opens the workbook read-only and inspects every worksheet.
2. For each worksheet, checks up to the first ten rows in its used range and
   treats the row with the most non-empty cells as the header row.
3. Skips sheets where it cannot detect headers, then displays usable sheet
   names, detected header rows, and columns.
4. Asks you to choose the main (parent) worksheet and one of its columns as the
   key. The key should identify each parent row uniquely.
5. Suggests root and record XML element names based on the sheet name. You may
   accept each default or enter a replacement.
6. Checks the selected parent key for blank or duplicate values. If found, it
   asks whether to continue; continuing can make child matching ambiguous.
7. For each other worksheet, asks whether it should be linked as a child. For
   every included child sheet, you choose the parent matching column and the
   child matching column, plus container and child-record element names.
8. Indexes child rows by the selected child key, then reports the complete
   parent/child matching plan for confirmation.
9. Emits one parent XML record per non-empty main row. Matching child rows are
   nested inside a child container, and the child matching column is omitted
   from each nested child record because the relationship is represented by
   the nesting.
10. Writes indented UTF-8 XML without a byte-order mark and reports parent and
    nested child record counts.

Column matching is a string-value equality comparison based on Excel's
displayed cell text. Values used as join keys should therefore be consistently
formatted in both worksheets. Parent rows with no matching children are still
written; they simply do not get that child container.

The interactive converter sanitizes element names (including user-provided
names) to XML-safe names and makes duplicate detected headers unique by
appending suffixes such as `_2`. It does not alter the source workbook.

## Choosing a script

Use `ExcelToXml.ps1` when a single worksheet is already arranged as a header
row followed by records, and you want a compact command suitable for repeat
runs. Use `InteractiveExcelToXml_Corrected.ps1` when you need to inspect a
workbook, choose columns interactively, and nest related worksheet records.
The scripts have different assumptions and are not interchangeable: the
first does not create parent/child relationships, while the second is
interactive and performs key-based matching.

## Troubleshooting

- **Excel COM cannot be created:** Install desktop Excel and run on Windows
  under an account that can start it.
- **Worksheet does not contain data / no usable worksheets:** Check that the
  workbook opens in Excel and includes non-empty records and column headings.
- **Empty or duplicate header errors in `ExcelToXml.ps1`:** Give every column
  a non-empty, distinct heading. Different source headings can also sanitize
  to the same name if punctuation is replaced with underscores.
- **Unexpected child rows:** Verify that both chosen match columns contain the
  same displayed values (including leading zeros, spaces, and date/number
  formatting). Avoid blank or duplicate parent keys.
- **Output cannot be written:** Check that the destination directory is
  writable and that the output file is not open or locked by another process.

Both scripts stop with an error on conversion failures and close Excel in a
cleanup block. Review the output XML and the reported record counts before
using it downstream.
