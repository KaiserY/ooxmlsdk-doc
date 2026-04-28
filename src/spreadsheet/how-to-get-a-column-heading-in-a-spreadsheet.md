# Get a column heading in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to retrieve a column heading in a spreadsheet document
programmatically.

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

## How the Sample Code Works

The code in this how-to consists of three methods (functions in Visual
Basic): `GetColumnHeading`, `GetColumnName`, and `GetRowIndex`. The last two methods are called from
within the `GetColumnHeading` method.

The `GetColumnName` method takes the cell
name as a parameter. It parses the cell name to get the column name by
creating a regular expression to match the column name portion of the
cell name. For more information about regular expressions, see [Regular Expression Language Elements](https://learn.microsoft.com/dotnet/standard/base-types/regular-expression-language-quick-reference).

### [C#](#tab/cs-1)
```csharp
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
```
### [Visual Basic](#tab/vb-1)
```vb
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
```
***

The `GetRowIndex` method takes the cell name
as a parameter. It parses the cell name to get the row index by creating
a regular expression to match the row index portion of the cell name.

### [C#](#tab/cs-2)
```csharp
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
```
### [Visual Basic](#tab/vb-2)
```vb
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
```
***

The `GetColumnHeading` method uses three
parameters, the full path to the source spreadsheet file, the name of
the worksheet that contains the specified column, and the name of a cell
in the column for which to get the heading.

The code gets the name of the column of the specified cell by calling
the `GetColumnName` method. The code also
gets the cells in the column and orders them by row using the `GetRowIndex` method.

### [C#](#tab/cs-3)
```csharp
        // Get the column name for the specified cell.
        string columnName = GetColumnName(cellName);

        // Get the cells in the specified column and order them by row.
        IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0)
            .OrderBy(r => GetRowIndex(r.CellReference) ?? 0);
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Get the column name for the specified cell.
            Dim columnName As String = GetColumnName(cellName)

            ' Get the cells in the specified column and order them by row.
            Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) String.Compare(GetColumnName(c.CellReference?.Value), columnName, True) = 0) _
                .OrderBy(Function(r) If(GetRowIndex(r.CellReference), 0))
```
***

If the specified column exists, it gets the first cell in the column
using the
`System.Linq.Enumerable.First`
method. The first cell contains the heading. Otherwise the specified column does not exist and the method returns `null` / `Nothing`

### [C#](#tab/cs-4)
```csharp
        if (cells.Count() == 0)
        {
            // The specified column does not exist.
            return null;
        }

        // Get the first cell in the column.
        Cell headCell = cells.First();
```
### [Visual Basic](#tab/vb-4)
```vb
            If cells.Count() = 0 Then
                ' The specified column does not exist.
                Return Nothing
            End If

            ' Get the first cell in the column.
            Dim headCell As Cell = cells.First()
```
***

If the content of the cell is stored in the `DocumentFormat.OpenXml.Packaging.SharedStringTablePart` object, it gets the
shared string items and returns the content of the column heading using
the
`System.Int32.Parse`
method. If the content of the cell is not in the `DocumentFormat.OpenXml.Spreadsheet.SharedStringTable` object, it returns the
content of the cell.

### [C#](#tab/cs-5)
```csharp
        // If the content of the first cell is stored as a shared string, get the text of the first cell
        // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
        if (headCell.DataType is not null && headCell.DataType.Value == CellValues.SharedString && int.TryParse(headCell.CellValue?.Text, out int index))
        {
            SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

            return items[index].InnerText;
        }
        else
        {
            return headCell.CellValue?.Text;
        }
```
### [Visual Basic](#tab/vb-5)
```vb
            ' If the content of the first cell is stored as a shared string, get the text of the first cell
            ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            Dim idx As Integer

            If headCell.DataType IsNot Nothing AndAlso headCell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(headCell.CellValue?.Text, idx) Then
                Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
                Dim items As SharedStringItem() = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()

                Return items(idx).InnerText
            Else
                Return headCell.CellValue?.Text
            End If
```
***

## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static string? GetColumnHeading(string docName, string worksheetName, string cellName)
{
    // Open the document as read-only.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return null;
        }

        string? id = sheets.First().Id;

        if (id is null)
        {
            // The worksheet does not have an ID.
            return null;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
        // Get the column name for the specified cell.
        string columnName = GetColumnName(cellName);

        // Get the cells in the specified column and order them by row.
        IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0)
            .OrderBy(r => GetRowIndex(r.CellReference) ?? 0);
        if (cells.Count() == 0)
        {
            // The specified column does not exist.
            return null;
        }

        // Get the first cell in the column.
        Cell headCell = cells.First();
        // If the content of the first cell is stored as a shared string, get the text of the first cell
        // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
        if (headCell.DataType is not null && headCell.DataType.Value == CellValues.SharedString && int.TryParse(headCell.CellValue?.Text, out int index))
        {
            SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

            return items[index].InnerText;
        }
        else
        {
            return headCell.CellValue?.Text;
        }
    }
}
// Given a cell name, parses the specified cell to get the column name.
static string GetColumnName(string? cellName)
{
    if (cellName is null)
    {
        return string.Empty;
    }
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
}

// Given a cell name, parses the specified cell to get the row index.
static uint? GetRowIndex(string? cellName)
{
    if (cellName is null)
    {
        return null;
    }
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetColumnHeading(docName As String, worksheetName As String, cellName As String) As String
        ' Open the document as read-only.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, False)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)

            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return Nothing
            End If

            Dim id As String = sheets.First().Id

            If id Is Nothing Then
                ' The worksheet does not have an ID.
                Return Nothing
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(id), WorksheetPart)
            ' Get the column name for the specified cell.
            Dim columnName As String = GetColumnName(cellName)

            ' Get the cells in the specified column and order them by row.
            Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) String.Compare(GetColumnName(c.CellReference?.Value), columnName, True) = 0) _
                .OrderBy(Function(r) If(GetRowIndex(r.CellReference), 0))
            If cells.Count() = 0 Then
                ' The specified column does not exist.
                Return Nothing
            End If

            ' Get the first cell in the column.
            Dim headCell As Cell = cells.First()
            ' If the content of the first cell is stored as a shared string, get the text of the first cell
            ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            Dim idx As Integer

            If headCell.DataType IsNot Nothing AndAlso headCell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(headCell.CellValue?.Text, idx) Then
                Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
                Dim items As SharedStringItem() = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()

                Return items(idx).InnerText
            Else
                Return headCell.CellValue?.Text
            End If
        End Using
    End Function

    ' Given a cell name, parses the specified cell to get the column name.
    Function GetColumnName(cellName As String) As String
        If cellName Is Nothing Then
            Return String.Empty
        End If
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
    End Function

    ' Given a cell name, parses the specified cell to get the row index.
    Function GetRowIndex(cellName As String) As UInteger?
        If cellName Is Nothing Then
            Return Nothing
        End If
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
    End Function
```
***

## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
