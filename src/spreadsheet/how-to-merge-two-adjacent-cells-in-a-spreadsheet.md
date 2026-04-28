# Merge two adjacent cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to merge two adjacent cells in a spreadsheet document
programmatically.

--------------------------------------------------------------------------------

## Getting a SpreadsheetDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the `SpreadsheetDocument` class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using `SpreadsheetML` markup.

To create the class instance from the document that you call one of the
`DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open` methods. Several are provided, each
with a different signature. The sample code in this topic uses the @"DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open*?text=Open(String, Boolean)" method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either `true` or `false` and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is `false`.

------------------------------------------------------

## Basic structure of a spreadsheetML document

The basic document structure of a `SpreadsheetML` document consists of the `DocumentFormat.OpenXml.Spreadsheet.Sheets` and `DocumentFormat.OpenXml.Spreadsheet.Sheet` elements, which reference the worksheets in the workbook. A separate XML file is created for each worksheet. For example, the `SpreadsheetML` for a `DocumentFormat.OpenXml.Spreadsheet.Workbook` that has two worksheets name MySheet1 and MySheet2 is located in the Workbook.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
`DocumentFormat.OpenXml.Spreadsheet.SheetData` represents the cell table and contains
one or more `DocumentFormat.OpenXml.Spreadsheet.Row` elements. A `row` contains one or more `DocumentFormat.OpenXml.Spreadsheet.Cell` elements. Each cell contains a `DocumentFormat.OpenXml.Spreadsheet.CellValue` element that represents the value
of the cell. For example, the `SpreadsheetML`
for the first worksheet in a workbook, that only has the value 100 in
cell A1, is located in the Sheet1.xml file and is shown in the following
code example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK, you can create document structure and
content that uses strongly-typed classes that correspond to `SpreadsheetML` elements. You can find these
classes in the `DocumentFormat.OpenXML.Spreadsheet` namespace. The
following table lists the class names of the classes that correspond to
the `workbook`, `sheets`, `sheet`, `worksheet`, and `sheetData` elements.

| **SpreadsheetML Element** | **Open XML SDK Class** | **Description** |
|:---|:---|:---|
| `<workbook/>` | DocumentFormat.OpenXML.Spreadsheet.Workbook | The root element for the main document part. |
| `<sheets/>` | DocumentFormat.OpenXML.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| `<sheet/>` | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| `<worksheet/>` | DocumentFormat.OpenXML.Spreadsheet. Worksheet | A sheet definition file that contains the sheet data. |
| `<sheetData/>` | DocumentFormat.OpenXML.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| `<row/>` | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| `<c/>` | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| `<v/>` | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |

--------------------------------------------------------------------------------

## How the Sample Code Works

After you have opened the spreadsheet file for editing, the code
verifies that the specified cells exist, and if they do not exist, it
creates them by calling the `CreateSpreadsheetCellIfNotExist` method and append
it to the appropriate `DocumentFormat.OpenXml.Spreadsheet.Row` object.

### [C#](#tab/cs-1)
```csharp
// Given a Worksheet and a cell name, verifies that the specified cell exists.
// If it does not exist, creates a new cell. 
static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
{
    string columnName = GetColumnName(cellName);
    uint rowIndex = GetRowIndex(cellName);

    IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex?.Value == rowIndex);

    // If the Worksheet does not contain the specified row, create the specified row.
    // Create the specified cell in that row, and insert the row into the Worksheet.
    if (rows.Count() == 0)
    {
        Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
        Cell cell = new Cell() { CellReference = new StringValue(cellName) };
        row.Append(cell);
        worksheet.Descendants<SheetData>().First().Append(row);
    }
    else
    {
        Row row = rows.First();

        IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference?.Value == cellName);

        // If the row does not contain the specified cell, create the specified cell.
        if (cells.Count() == 0)
        {
            Cell cell = new Cell() { CellReference = new StringValue(cellName) };
            row.Append(cell);
        }
    }
}
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Given a Worksheet and a cell name, verifies that the specified cell exists.
    ' If it does not exist, creates a new cell. 
    Sub CreateSpreadsheetCellIfNotExist(worksheet As Worksheet, cellName As String)
        Dim columnName As String = GetColumnName(cellName)
        Dim rowIndex As UInteger = GetRowIndex(cellName)

        Dim rows As IEnumerable(Of Row) = worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex?.Value = rowIndex)

        ' If the Worksheet does not contain the specified row, create the specified row.
        ' Create the specified cell in that row, and insert the row into the Worksheet.
        If rows.Count() = 0 Then
            Dim row As New Row() With {
                .RowIndex = New UInt32Value(rowIndex)
            }
            Dim cell As New Cell() With {
                .CellReference = New StringValue(cellName)
            }
            row.Append(cell)
            worksheet.Descendants(Of SheetData)().First().Append(row)
        Else
            Dim row As Row = rows.First()

            Dim cells As IEnumerable(Of Cell) = row.Elements(Of Cell)().Where(Function(c) c.CellReference?.Value = cellName)

            ' If the row does not contain the specified cell, create the specified cell.
            If cells.Count() = 0 Then
                Dim cell As New Cell() With {
                    .CellReference = New StringValue(cellName)
                }
                row.Append(cell)
            End If
        End If
    End Sub
```
***

In order to get a column name, the code creates a new regular expression
to match the column name portion of the cell name. This regular
expression matches any combination of uppercase or lowercase letters.
For more information about regular expressions, see [Regular Expression Language Elements](https://learn.microsoft.com/dotnet/standard/base-types/regular-expressions). The
code gets the column name by calling the [Regex.Match](https://learn.microsoft.com/dotnet/api/system.text.regularexpressions.regex.match#overloads).

### [C#](#tab/cs-2)
```csharp
// Given a cell name, parses the specified cell to get the column name.
static string GetColumnName(string cellName)
{
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
}
```

### [Visual Basic](#tab/vb-2)
```vb
    ' Given a cell name, parses the specified cell to get the column name.
    Function GetColumnName(cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
    End Function
```
***

To get the row index, the code creates a new regular expression to match the row index portion of the cell name. This regular expression matches any combination of decimal digits. The following code creates a regular expression to match the row index portion of the cell name, comprised of decimal digits.

### [C#](#tab/cs-3)
```csharp
// Given a cell name, parses the specified cell to get the row index.
static uint GetRowIndex(string cellName)
{
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
}
```

### [Visual Basic](#tab/vb-3)
```vb
    ' Given a cell name, parses the specified cell to get the row index.
    Function GetRowIndex(cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
    End Function
```
***

## Sample Code

The following code merges two adjacent cells in a `DocumentFormat.OpenXml.Spreadsheet.Row` document package. When
merging two cells, only the content from one of the cells is preserved.
In left-to-right languages, the content in the upper-left cell is
preserved. In right-to-left languages, the content in the upper-right
cell is preserved.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        Worksheet? worksheet = GetWorksheet(document, sheetName);
        if (worksheet is null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
        {
            return;
        }

        // Verify if the specified cells exist, and if they do not exist, create them.
        CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
        CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

        MergeCells mergeCells;
        if (worksheet.Elements<MergeCells>().Count() > 0)
        {
            mergeCells = worksheet.Elements<MergeCells>().First();
        }
        else
        {
            mergeCells = new MergeCells();

            // Insert a MergeCells object into the specified position.
            if (worksheet.Elements<CustomSheetView>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
            }
            else if (worksheet.Elements<DataConsolidate>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
            }
            else if (worksheet.Elements<SortState>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
            }
            else if (worksheet.Elements<AutoFilter>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
            }
            else if (worksheet.Elements<Scenarios>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
            }
            else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
            }
            else if (worksheet.Elements<SheetProtection>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
            }
            else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
            }
            else
            {
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            }
        }

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
        mergeCells.Append(mergeCell);
    }
}
// Given a Worksheet and a cell name, verifies that the specified cell exists.
// If it does not exist, creates a new cell. 
static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
{
    string columnName = GetColumnName(cellName);
    uint rowIndex = GetRowIndex(cellName);

    IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex?.Value == rowIndex);

    // If the Worksheet does not contain the specified row, create the specified row.
    // Create the specified cell in that row, and insert the row into the Worksheet.
    if (rows.Count() == 0)
    {
        Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
        Cell cell = new Cell() { CellReference = new StringValue(cellName) };
        row.Append(cell);
        worksheet.Descendants<SheetData>().First().Append(row);
    }
    else
    {
        Row row = rows.First();

        IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference?.Value == cellName);

        // If the row does not contain the specified cell, create the specified cell.
        if (cells.Count() == 0)
        {
            Cell cell = new Cell() { CellReference = new StringValue(cellName) };
            row.Append(cell);
        }
    }
}
// Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
static Worksheet? GetWorksheet(SpreadsheetDocument document, string worksheetName)
{
    WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

    string? id = sheets.First().Id;
    WorksheetPart? worksheetPart = id is not null ? (WorksheetPart)workbookPart.GetPartById(id) : null;

    return worksheetPart?.Worksheet;
}
// Given a cell name, parses the specified cell to get the column name.
static string GetColumnName(string cellName)
{
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
}
// Given a cell name, parses the specified cell to get the row index.
static uint GetRowIndex(string cellName)
{
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub MergeTwoCells(docName As String, sheetName As String, cell1Name As String, cell2Name As String)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim worksheet As Worksheet = GetWorksheet(document, sheetName)
            If worksheet Is Nothing OrElse String.IsNullOrEmpty(cell1Name) OrElse String.IsNullOrEmpty(cell2Name) Then
                Return
            End If

            ' Verify if the specified cells exist, and if they do not exist, create them.
            CreateSpreadsheetCellIfNotExist(worksheet, cell1Name)
            CreateSpreadsheetCellIfNotExist(worksheet, cell2Name)

            Dim mergeCells As MergeCells
            If worksheet.Elements(Of MergeCells)().Count() > 0 Then
                mergeCells = worksheet.Elements(Of MergeCells)().First()
            Else
                mergeCells = New MergeCells()

                ' Insert a MergeCells object into the specified position.
                If worksheet.Elements(Of CustomSheetView)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of CustomSheetView)().First())
                ElseIf worksheet.Elements(Of DataConsolidate)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of DataConsolidate)().First())
                ElseIf worksheet.Elements(Of SortState)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SortState)().First())
                ElseIf worksheet.Elements(Of AutoFilter)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of AutoFilter)().First())
                ElseIf worksheet.Elements(Of Scenarios)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of Scenarios)().First())
                ElseIf worksheet.Elements(Of ProtectedRanges)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of ProtectedRanges)().First())
                ElseIf worksheet.Elements(Of SheetProtection)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetProtection)().First())
                ElseIf worksheet.Elements(Of SheetCalculationProperties)().Count() > 0 Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetCalculationProperties)().First())
                Else
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetData)().First())
                End If
            End If

            ' Create the merged cell and append it to the MergeCells collection.
            Dim mergeCell As New MergeCell() With {
                .Reference = New StringValue(cell1Name & ":" & cell2Name)
            }
            mergeCells.Append(mergeCell)
        End Using
    End Sub
    ' Given a Worksheet and a cell name, verifies that the specified cell exists.
    ' If it does not exist, creates a new cell. 
    Sub CreateSpreadsheetCellIfNotExist(worksheet As Worksheet, cellName As String)
        Dim columnName As String = GetColumnName(cellName)
        Dim rowIndex As UInteger = GetRowIndex(cellName)

        Dim rows As IEnumerable(Of Row) = worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex?.Value = rowIndex)

        ' If the Worksheet does not contain the specified row, create the specified row.
        ' Create the specified cell in that row, and insert the row into the Worksheet.
        If rows.Count() = 0 Then
            Dim row As New Row() With {
                .RowIndex = New UInt32Value(rowIndex)
            }
            Dim cell As New Cell() With {
                .CellReference = New StringValue(cellName)
            }
            row.Append(cell)
            worksheet.Descendants(Of SheetData)().First().Append(row)
        Else
            Dim row As Row = rows.First()

            Dim cells As IEnumerable(Of Cell) = row.Elements(Of Cell)().Where(Function(c) c.CellReference?.Value = cellName)

            ' If the row does not contain the specified cell, create the specified cell.
            If cells.Count() = 0 Then
                Dim cell As New Cell() With {
                    .CellReference = New StringValue(cellName)
                }
                row.Append(cell)
            End If
        End If
    End Sub
    ' Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
    Function GetWorksheet(document As SpreadsheetDocument, worksheetName As String) As Worksheet
        Dim workbookPart As WorkbookPart = If(document.WorkbookPart, document.AddWorkbookPart())
        Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)

        Dim id As String = sheets.First().Id
        Dim worksheetPart As WorksheetPart = If(id IsNot Nothing, CType(workbookPart.GetPartById(id), WorksheetPart), Nothing)

        Return If(worksheetPart IsNot Nothing, worksheetPart.Worksheet, Nothing)
    End Function
    ' Given a cell name, parses the specified cell to get the column name.
    Function GetColumnName(cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
    End Function
    ' Given a cell name, parses the specified cell to get the row index.
    Function GetRowIndex(cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
    End Function
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
