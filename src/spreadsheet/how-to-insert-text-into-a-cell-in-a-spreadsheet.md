# Insert text into a cell in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to insert text into a cell in a new worksheet in a spreadsheet
document programmatically.

--------------------------------------------------------------------------------

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
After opening the `SpreadsheetDocument`
document for editing, the code inserts a blank `DocumentFormat.OpenXml.Packaging.WorksheetPart.Worksheet` object into a `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` document package. Then,
inserts a new `DocumentFormat.OpenXml.Spreadsheet.Cell` object into the new worksheet and
inserts the specified text into that cell.

### [C#](#tab/cs-1)
```csharp
// Given a document name and text, 
// inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
static void InsertText(string docName, string text)
{
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
    {
        WorkbookPart workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

        // Get the SharedStringTablePart. If it does not exist, create a new one.
        SharedStringTablePart shareStringPart;
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }
        else
        {
            shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        }

        // Insert the text into the SharedStringTablePart.
        int index = InsertSharedStringItem(text, shareStringPart);

        // Insert a new worksheet.
        WorksheetPart worksheetPart = InsertWorksheet(workbookPart);

        // Insert cell A1 into the new worksheet.
        Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

        // Set the value of cell A1.
        cell.CellValue = new CellValue(index.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
    }
}
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Given a document name and text, 
    ' inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
    Sub InsertText(docName As String, text As String)
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim workbookPart As WorkbookPart = If(spreadSheet.WorkbookPart, spreadSheet.AddWorkbookPart())

            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If workbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0 Then
                shareStringPart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = workbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(workbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString())
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        End Using
    End Sub
```
***

The code passes in a parameter that represents the text to insert into
the cell and a parameter that represents the `SharedStringTablePart` object for the spreadsheet.
If the `ShareStringTablePart` object does not
contain a `DocumentFormat.OpenXml.Spreadsheet.SharedStringTable` object, the code creates
one. If the text already exists in the `ShareStringTable` object, the code returns the
index for the `DocumentFormat.OpenXml.Spreadsheet.SharedStringItem` object that represents the
text. Otherwise, it creates a new `SharedStringItem` object that represents the text.

The following code verifies if the specified text exists in the `SharedStringTablePart` object and add the text if
it does not exist.

### [C#](#tab/cs-2)
```csharp
// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
{
    // If the part does not contain a SharedStringTable, create one.
    shareStringPart.SharedStringTable ??= new SharedStringTable();

    int i = 0;

    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
    {
        if (item.InnerText == text)
        {
            return i;
        }

        i++;
    }

    // The text does not exist in the part. Create the SharedStringItem and return its index.
    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

    return i;
}
```

### [Visual Basic](#tab/vb-2)
```vb
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                Return i
            End If

            i += 1
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
```
***

The code adds a new `WorksheetPart` object to
the `WorkbookPart` object by using the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart` method. It then adds a new `Worksheet` object to the `WorksheetPart` object, and gets a unique ID for
the new worksheet by selecting the maximum `DocumentFormat.OpenXml.Spreadsheet.Sheet.SheetId` object used within the spreadsheet
document and adding one to create the new sheet ID. It gives the
worksheet a name by concatenating the word "Sheet" with the sheet ID. It
then appends the new `Sheet` object to the
`Sheets` collection.

The following code inserts a new `Worksheet`
object by adding a new `WorksheetPart` object
to the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart` object.

### [C#](#tab/cs-3)
```csharp
// Given a WorkbookPart, inserts a new worksheet.
static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
{
    // Add a new worksheet part to the workbook.
    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    newWorksheetPart.Worksheet = new Worksheet(new SheetData());

    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

    // Get a unique ID for the new sheet.
    uint sheetId = 1;
    if (sheets.Elements<Sheet>().Count() > 0)
    {
        sheetId = sheets.Elements<Sheet>().Select<Sheet, uint>(s =>
        {
            if (s.SheetId is not null && s.SheetId.HasValue)
            {
                return s.SheetId.Value;
            }

            return 0;
        }).Max() + 1;
    }

    string sheetName = "Sheet" + sheetId;

    // Append the new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
    sheets.Append(sheet);

    return newWorksheetPart;
}
```

### [Visual Basic](#tab/vb-3)
```vb
    ' Given a WorkbookPart, inserts a new worksheet.
    Function InsertWorksheet(workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())

        Dim sheets As Sheets = If(workbookPart.Workbook.GetFirstChild(Of Sheets)(), workbookPart.Workbook.AppendChild(New Sheets()))
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If sheets.Elements(Of Sheet)().Count() > 0 Then
            sheetId = sheets.Elements(Of Sheet)().Select(Function(s)
                                                             If s.SheetId IsNot Nothing AndAlso s.SheetId.HasValue Then
                                                                 Return s.SheetId.Value
                                                             End If

                                                             Return 0
                                                         End Function).Max() + 1
        End If

        Dim sheetName As String = "Sheet" & sheetId

        ' Append the new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {
            .Id = relationshipId,
            .SheetId = sheetId,
            .Name = sheetName
        }
        sheets.Append(sheet)

        Return newWorksheetPart
    End Function
```
***

To insert a cell into a worksheet, the code determines where to insert
the new cell in the column by iterating through the row elements to find
the cell that comes directly after the specified row, in sequential
order. It saves that row in the `refCell`
variable. It then inserts the new cell before the cell referenced by
`refCell` using the `DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore` method.

In the following code, insert a new `Cell`
object into a `Worksheet` object.

### [C#](#tab/cs-4)
```csharp
// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
// If the cell already exists, returns it. 
static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
{
    Worksheet worksheet = worksheetPart.Worksheet;
    SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
    string cellReference = columnName + rowIndex;

    // If the worksheet does not contain a row with the specified row index, insert one.
    Row row;

    if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
    {
        row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
    }
    else
    {
        row = new Row() { RowIndex = rowIndex };
        sheetData.Append(row);
    }

    // If there is not a cell with the specified column name, insert one.  
    if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
    {
        return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
    }
    else
    {
        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        Cell? refCell = null;

        foreach (Cell cell in row.Elements<Cell>())
        {
            if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
            {
                refCell = cell;
                break;
            }
        }

        Cell newCell = new Cell() { CellReference = cellReference };
        row.InsertBefore(newCell, refCell);

        return newCell;
    }
}
```

### [Visual Basic](#tab/vb-4)
```vb
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, returns it. 
    Function InsertCellInWorksheet(columnName As String, rowIndex As UInteger, worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = columnName & rowIndex

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row

        If sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).Count() <> 0 Then
            row = sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).First()
        Else
            row = New Row() With {
                .RowIndex = rowIndex
            }
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = columnName & rowIndex).Count() > 0 Then
            Return row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing

            For Each cell As Cell In row.Elements(Of Cell)()
                If String.Compare(cell.CellReference?.Value, cellReference, True) > 0 Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As New Cell() With {
                .CellReference = cellReference
            }
            row.InsertBefore(newCell, refCell)

            Return newCell
        End If
    End Function
```
***

--------------------------------------------------------------------------------
## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void InsertText(string docName, string text)
{
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
    {
        WorkbookPart workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

        // Get the SharedStringTablePart. If it does not exist, create a new one.
        SharedStringTablePart shareStringPart;
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }
        else
        {
            shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        }

        // Insert the text into the SharedStringTablePart.
        int index = InsertSharedStringItem(text, shareStringPart);

        // Insert a new worksheet.
        WorksheetPart worksheetPart = InsertWorksheet(workbookPart);

        // Insert cell A1 into the new worksheet.
        Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

        // Set the value of cell A1.
        cell.CellValue = new CellValue(index.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
    }
}
// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
{
    // If the part does not contain a SharedStringTable, create one.
    shareStringPart.SharedStringTable ??= new SharedStringTable();

    int i = 0;

    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
    {
        if (item.InnerText == text)
        {
            return i;
        }

        i++;
    }

    // The text does not exist in the part. Create the SharedStringItem and return its index.
    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

    return i;
}
// Given a WorkbookPart, inserts a new worksheet.
static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
{
    // Add a new worksheet part to the workbook.
    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    newWorksheetPart.Worksheet = new Worksheet(new SheetData());

    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

    // Get a unique ID for the new sheet.
    uint sheetId = 1;
    if (sheets.Elements<Sheet>().Count() > 0)
    {
        sheetId = sheets.Elements<Sheet>().Select<Sheet, uint>(s =>
        {
            if (s.SheetId is not null && s.SheetId.HasValue)
            {
                return s.SheetId.Value;
            }

            return 0;
        }).Max() + 1;
    }

    string sheetName = "Sheet" + sheetId;

    // Append the new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
    sheets.Append(sheet);

    return newWorksheetPart;
}
// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
// If the cell already exists, returns it. 
static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
{
    Worksheet worksheet = worksheetPart.Worksheet;
    SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
    string cellReference = columnName + rowIndex;

    // If the worksheet does not contain a row with the specified row index, insert one.
    Row row;

    if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
    {
        row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
    }
    else
    {
        row = new Row() { RowIndex = rowIndex };
        sheetData.Append(row);
    }

    // If there is not a cell with the specified column name, insert one.  
    if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
    {
        return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
    }
    else
    {
        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        Cell? refCell = null;

        foreach (Cell cell in row.Elements<Cell>())
        {
            if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
            {
                refCell = cell;
                break;
            }
        }

        Cell newCell = new Cell() { CellReference = cellReference };
        row.InsertBefore(newCell, refCell);

        return newCell;
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub InsertText(docName As String, text As String)
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim workbookPart As WorkbookPart = If(spreadSheet.WorkbookPart, spreadSheet.AddWorkbookPart())

            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If workbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0 Then
                shareStringPart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = workbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(workbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString())
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        End Using
    End Sub
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                Return i
            End If

            i += 1
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
    ' Given a WorkbookPart, inserts a new worksheet.
    Function InsertWorksheet(workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())

        Dim sheets As Sheets = If(workbookPart.Workbook.GetFirstChild(Of Sheets)(), workbookPart.Workbook.AppendChild(New Sheets()))
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If sheets.Elements(Of Sheet)().Count() > 0 Then
            sheetId = sheets.Elements(Of Sheet)().Select(Function(s)
                                                             If s.SheetId IsNot Nothing AndAlso s.SheetId.HasValue Then
                                                                 Return s.SheetId.Value
                                                             End If

                                                             Return 0
                                                         End Function).Max() + 1
        End If

        Dim sheetName As String = "Sheet" & sheetId

        ' Append the new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {
            .Id = relationshipId,
            .SheetId = sheetId,
            .Name = sheetName
        }
        sheets.Append(sheet)

        Return newWorksheetPart
    End Function
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, returns it. 
    Function InsertCellInWorksheet(columnName As String, rowIndex As UInteger, worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = columnName & rowIndex

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row

        If sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).Count() <> 0 Then
            row = sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).First()
        Else
            row = New Row() With {
                .RowIndex = rowIndex
            }
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = columnName & rowIndex).Count() > 0 Then
            Return row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing

            For Each cell As Cell In row.Elements(Of Cell)()
                If String.Compare(cell.CellReference?.Value, cellReference, True) > 0 Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As New Cell() With {
                .CellReference = cellReference
            }
            row.InsertBefore(newCell, refCell)

            Return newCell
        End If
    End Function
```

--------------------------------------------------------------------------------
## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
