# Delete text from a cell in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to delete text from a cell in a spreadsheet document
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

## How the sample code works

In the following code example, you delete text from a cell in a `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` document package. Then, you verify if other cells within the spreadsheet document still reference the text removed from the row, and if they do not, you remove the text from the `DocumentFormat.OpenXml.Packaging.SharedStringTablePart` object by using the `DocumentFormat.OpenXml.OpenXmlElement.Remove` method. Then you clean up the `SharedStringTablePart` object by calling the `RemoveSharedStringItem` method.

### [C#](#tab/cs-1)
```csharp
// Given a document, a worksheet name, a column name, and a one-based row index,
// deletes the text from the cell at the specified column and row on the specified worksheet.
static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.Where(s => s.Name is not null && s.Name == sheetName);
        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }
        string? relationshipId = sheets.First()?.Id?.Value;

        if (relationshipId is null)
        {
            // The worksheet does not have a relationship ID.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(relationshipId);

        // Get the cell at the specified column and row.
        Cell? cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
        if (cell is null)
        {
            // The specified cell does not exist.
            return;
        }

        cell.Remove();
    }
}
```
### [Visual Basic](#tab/vb-1)
```vb
    ' Given a document, a worksheet name, a column name, and a one-based row index,
    ' deletes the text from the cell at the specified column and row on the specified worksheet.
    Sub DeleteTextFromCell(docName As String, sheetName As String, colName As String, rowIndex As UInteger)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook?.GetFirstChild(Of Sheets)()?.Elements(Of Sheet)()?.Where(Function(s) s.Name IsNot Nothing AndAlso s.Name = sheetName)
            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim relationshipId As String = sheets.First()?.Id?.Value

            If relationshipId Is Nothing Then
                ' The worksheet does not have a relationship ID.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If cell Is Nothing Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
        End Using
    End Sub
```
***

In the following code example, you verify that the cell specified by the column name and row index exists. If so, the code returns the cell; otherwise, it returns `null`.

### [C#](#tab/cs-2)
```csharp
// Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
static Cell? GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
{
    IEnumerable<Row>? rows = worksheet.GetFirstChild<SheetData>()?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex);
    if (rows is null || rows.Count() == 0)
    {
        // A cell does not exist at the specified row.
        return null;
    }

    IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference?.Value, columnName + rowIndex, true) == 0);

    if (cells.Count() == 0)
    {
        // A cell does not exist at the specified column, in the specified row.
        return null;
    }

    return cells.FirstOrDefault();
}
```
### [Visual Basic](#tab/vb-2)
```vb
    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Function GetSpreadsheetCell(worksheet As Worksheet, columnName As String, rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)()?.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex))
        If rows Is Nothing OrElse rows.Count() = 0 Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference?.Value, columnName & rowIndex, True) = 0)

        If cells.Count() = 0 Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.FirstOrDefault()
    End Function
```
***

In the following code example, you verify if other cells within the
spreadsheet document reference the text specified by the `shareStringId` parameter. If they do not reference
the text, you remove it from the `SharedStringTablePart` object. You do that by
passing a parameter that represents the ID of the text to remove and a
parameter that represents the `SpreadsheetDocument` document package. Then you
iterate through each `Worksheet` object and
compare the contents of each `Cell` object to
the shared string ID. If other cells within the spreadsheet document
still reference the `DocumentFormat.OpenXml.Spreadsheet.SharedStringItem` object, you do not remove
the item from the `SharedStringTablePart`
object. If other cells within the spreadsheet document no longer
reference the `SharedStringItem` object, you
remove the item from the `SharedStringTablePart` object. Then you iterate
through each `Worksheet` object and `Cell` object and refresh the shared string
references.

### [C#](#tab/cs-3)
```csharp
// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
// reference the specified SharedStringItem and removes the item.
static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
{
    bool remove = true;

    if (document.WorkbookPart is null)
    {
        return;
    }

    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
    {
        var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

        if (cells is null)
        {
            continue;
        }

        foreach (var cell in cells)
        {
            // Verify if other cells in the document reference the item.
            if (cell.DataType is not null &&
                cell.DataType.Value == CellValues.SharedString &&
                cell.CellValue?.Text == shareStringId.ToString())
            {
                // Other cells in the document still reference the item. Do not remove the item.
                remove = false;
                break;
            }
        }

        if (!remove)
        {
            break;
        }
    }

    // Other cells in the document do not reference the item. Remove the item.
    if (remove)
    {
        SharedStringTablePart? shareStringTablePart = document.WorkbookPart.SharedStringTablePart;

        if (shareStringTablePart is null)
        {
            return;
        }

        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
        if (item is not null)
        {
            item.Remove();

            // Refresh all the shared string references.
            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

                if (cells is null)
                {
                    continue;
                }

                foreach (var cell in cells)
                {
                    if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int itemIndex))
                    {
                        if (itemIndex > shareStringId)
                        {
                            cell.CellValue.Text = (itemIndex - 1).ToString();
                        }
                    }
                }
            }
        }
    }
}
```
### [Visual Basic](#tab/vb-3)
```vb
    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Sub RemoveSharedStringItem(shareStringId As Integer, document As SpreadsheetDocument)
        Dim remove As Boolean = True

        If document.WorkbookPart Is Nothing Then
            Return
        End If

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

            If cells Is Nothing Then
                Continue For
            End If

            For Each cell In cells
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso
                   cell.DataType.Value = CellValues.SharedString AndAlso
                   cell.CellValue?.Text = shareStringId.ToString() Then
                    ' Other cells in the document still reference the item. Do not remove the item.
                    remove = False
                    Exit For
                End If
            Next

            If Not remove Then
                Exit For
            End If
        Next

        ' Other cells in the document do not reference the item. Remove the item.
        If remove Then
            Dim shareStringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart

            If shareStringTablePart Is Nothing Then
                Return
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

                    If cells Is Nothing Then
                        Continue For
                    End If

                    For Each cell In cells
                        Dim itemIndex As Integer

                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(cell.CellValue?.Text, itemIndex) Then
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub
```
***

## Sample code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.Where(s => s.Name is not null && s.Name == sheetName);
        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }
        string? relationshipId = sheets.First()?.Id?.Value;

        if (relationshipId is null)
        {
            // The worksheet does not have a relationship ID.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(relationshipId);

        // Get the cell at the specified column and row.
        Cell? cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
        if (cell is null)
        {
            // The specified cell does not exist.
            return;
        }

        cell.Remove();
    }
}
// Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
static Cell? GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
{
    IEnumerable<Row>? rows = worksheet.GetFirstChild<SheetData>()?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex);
    if (rows is null || rows.Count() == 0)
    {
        // A cell does not exist at the specified row.
        return null;
    }

    IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference?.Value, columnName + rowIndex, true) == 0);

    if (cells.Count() == 0)
    {
        // A cell does not exist at the specified column, in the specified row.
        return null;
    }

    return cells.FirstOrDefault();
}
// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
// reference the specified SharedStringItem and removes the item.
static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
{
    bool remove = true;

    if (document.WorkbookPart is null)
    {
        return;
    }

    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
    {
        var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

        if (cells is null)
        {
            continue;
        }

        foreach (var cell in cells)
        {
            // Verify if other cells in the document reference the item.
            if (cell.DataType is not null &&
                cell.DataType.Value == CellValues.SharedString &&
                cell.CellValue?.Text == shareStringId.ToString())
            {
                // Other cells in the document still reference the item. Do not remove the item.
                remove = false;
                break;
            }
        }

        if (!remove)
        {
            break;
        }
    }

    // Other cells in the document do not reference the item. Remove the item.
    if (remove)
    {
        SharedStringTablePart? shareStringTablePart = document.WorkbookPart.SharedStringTablePart;

        if (shareStringTablePart is null)
        {
            return;
        }

        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
        if (item is not null)
        {
            item.Remove();

            // Refresh all the shared string references.
            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

                if (cells is null)
                {
                    continue;
                }

                foreach (var cell in cells)
                {
                    if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int itemIndex))
                    {
                        if (itemIndex > shareStringId)
                        {
                            cell.CellValue.Text = (itemIndex - 1).ToString();
                        }
                    }
                }
            }
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub DeleteTextFromCell(docName As String, sheetName As String, colName As String, rowIndex As UInteger)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook?.GetFirstChild(Of Sheets)()?.Elements(Of Sheet)()?.Where(Function(s) s.Name IsNot Nothing AndAlso s.Name = sheetName)
            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim relationshipId As String = sheets.First()?.Id?.Value

            If relationshipId Is Nothing Then
                ' The worksheet does not have a relationship ID.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If cell Is Nothing Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
        End Using
    End Sub
    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Function GetSpreadsheetCell(worksheet As Worksheet, columnName As String, rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)()?.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex))
        If rows Is Nothing OrElse rows.Count() = 0 Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference?.Value, columnName & rowIndex, True) = 0)

        If cells.Count() = 0 Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.FirstOrDefault()
    End Function
    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Sub RemoveSharedStringItem(shareStringId As Integer, document As SpreadsheetDocument)
        Dim remove As Boolean = True

        If document.WorkbookPart Is Nothing Then
            Return
        End If

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

            If cells Is Nothing Then
                Continue For
            End If

            For Each cell In cells
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso
                   cell.DataType.Value = CellValues.SharedString AndAlso
                   cell.CellValue?.Text = shareStringId.ToString() Then
                    ' Other cells in the document still reference the item. Do not remove the item.
                    remove = False
                    Exit For
                End If
            Next

            If Not remove Then
                Exit For
            End If
        Next

        ' Other cells in the document do not reference the item. Remove the item.
        If remove Then
            Dim shareStringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart

            If shareStringTablePart Is Nothing Then
                Return
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

                    If cells Is Nothing Then
                        Continue For
                    End If

                    For Each cell In cells
                        Dim itemIndex As Integer

                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(cell.CellValue?.Text, itemIndex) Then
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub
```

## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
