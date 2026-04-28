# Calculate the sum of a range of cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to calculate the sum of a contiguous range of cells in a spreadsheet document programmatically.

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

The sample code starts by passing in to the method `CalculateSumOfCellRange` a parameter that represents the full path to the source `SpreadsheetML` file, a parameter that represents the name of the worksheet that contains the cells, a parameter that represents the name of the first cell in the contiguous range, a parameter that represent the name of the last cell in the contiguous range, and a parameter that represents the name of the cell where you want the result displayed.

The code then opens the file for editing as a `SpreadsheetDocument` document package for read/write access, the code gets the specified `Worksheet` object. It then gets the index of the row for the first and last cell in the contiguous range by calling the `GetRowIndex` method. It gets the name of the column for the first and last cell in the contiguous range by calling the `GetColumnName` method.

For each `Row` object within the contiguous range, the code iterates through each `Cell` object and determines if the column of the cell is within the contiguous
range by calling the `CompareColumn` method. If the cell is within the contiguous range, the code adds the value of the cell to the sum. Then it gets the `SharedStringTablePart` object if it exists. If it does not exist, it creates one using the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart` method. It inserts the result into the `SharedStringTablePart` object by calling the `InsertSharedStringItem` method.

The code inserts a new cell for the result into the worksheet by calling the `InsertCellInWorksheet` method and sets the value of the cell. For more information, see [how to insert a cell in a spreadsheet](how-to-insert-text-into-a-cell-in-a-spreadsheet.md#how-the-sample-code-works).

### [C#](#tab/cs-1)
```csharp
static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
        string? firstId = sheets?.First().Id;
        if (sheets is null || firstId is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(firstId);
        Worksheet worksheet = worksheetPart.Worksheet;

        // Get the row number and column name for the first and last cells in the range.
        uint firstRowNum = GetRowIndex(firstCellName);
        uint lastRowNum = GetRowIndex(lastCellName);
        string firstColumn = GetColumnName(firstCellName);
        string lastColumn = GetColumnName(lastCellName);

        double sum = 0;

        // Iterate through the cells within the range and add their values to the sum.
        foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex is not null && r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
        {
            foreach (Cell cell in row)
            {
                if (cell.CellReference is not null && cell.CellReference.Value is not null)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0 && double.TryParse(cell.CellValue?.Text, out double num))
                    {
                        sum += num;
                    }
                }
            }
        }

        // Get the SharedStringTablePart and add the result to it.
        // If the SharedStringPart does not exist, create a new one.
        SharedStringTablePart shareStringPart;
        if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }
        else
        {
            shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
        }

        // Insert the result into the SharedStringTablePart.
        int index = InsertSharedStringItem("Result: " + sum, shareStringPart);

        Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

        // Set the value of the cell.
        result.CellValue = new CellValue(index.ToString());
        result.DataType = new EnumValue<CellValues>(CellValues.SharedString);
    }
}
```
### [Visual Basic](#tab/vb-1)
```vb
    Private Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String,
    ByVal lastCellName As String, ByVal resultCell As String)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) =
                document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If (sheets.Count() = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)
            Dim worksheet As Worksheet = worksheetPart.Worksheet

            ' Get the row number and column name for the first and last cells in the range.
            Dim firstRowNum As UInteger = GetRowIndex(firstCellName)
            Dim lastRowNum As UInteger = GetRowIndex(lastCellName)
            Dim firstColumn As String = GetColumnName(firstCellName)
            Dim lastColumn As String = GetColumnName(lastCellName)

            Dim sum As Double = 0

            ' Iterate through the cells within the range and add their values to the sum.
            For Each row As Row In worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value >= firstRowNum _
                                                                             AndAlso r.RowIndex.Value <= lastRowNum)
                For Each cell As Cell In row
                    Dim columnName As String = GetColumnName(cell.CellReference.Value)
                    If ((CompareColumn(columnName, firstColumn) >= 0) AndAlso (CompareColumn(columnName, lastColumn) <= 0)) Then
                        sum = (sum + Double.Parse(cell.CellValue.Text))
                    End If
                Next
            Next

            ' Get the SharedStringTablePart and add the result to it.
            ' If the SharedStringPart does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If (document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0) Then
                shareStringPart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = document.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the result into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(("Result:" + sum.ToString()), shareStringPart)

            Dim result As Cell = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart)

            ' Set the value of the cell.
            result.CellValue = New CellValue(index.ToString())
            result.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        End Using
    End Sub
```
***

To get the row index the code passes a parameter that represents the name of the cell, and creates a new regular expression to match the row
index portion of the cell name. For more information about regular expressions, see [Regular Expression Language Elements](https://learn.microsoft.com/dotnet/standard/base-types/regular-expression-language-quick-reference). It gets the row index by calling the `System.Text.RegularExpressions.Regex.Match` method, and then returns the row index.

### [C#](#tab/cs-2)
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
### [Visual Basic](#tab/vb-2)
```vb
    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
```
***

The code then gets the column name by passing a parameter that represents the name of the cell, and creates a new regular expression to match the column name portion of the cell name. This regular expression matches any combination of uppercase or lowercase letters. It gets the column name by calling the `System.Text.RegularExpressions.Regex.Match` method, and then returns the column name.

### [C#](#tab/cs-3)
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
### [Visual Basic](#tab/vb-3)
```vb
    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function
```
***

To compare two columns the code passes in two parameters that represent the columns to compare. If the first column is longer than the second column, it returns 1. If the second column is longer than the first column, it returns -1. Otherwise, it compares the values of the columns using the `System.String.Compare` and returns the result.

### [C#](#tab/cs-4)
```csharp
// Given two columns, compares the columns.
static int CompareColumn(string column1, string column2)
{
    if (column1.Length > column2.Length)
    {
        return 1;
    }
    else if (column1.Length < column2.Length)
    {
        return -1;
    }
    else
    {
        return string.Compare(column1, column2, true);
    }
}
```
### [Visual Basic](#tab/vb-4)
```vb
    ' Given two columns, compares the columns.
    Private Function CompareColumn(ByVal column1 As String, ByVal column2 As String) As Integer
        If (column1.Length > column2.Length) Then
            Return 1
        ElseIf (column1.Length < column2.Length) Then
            Return -1
        Else
            Return String.Compare(column1, column2, True)
        End If
    End Function
```
***

To insert a `SharedStringItem`, the code passes in a parameter that represents the text to insert into the cell and a parameter that represents the  `SharedStringTablePart` object for the spreadsheet. If the `ShareStringTablePart` object does not contain a `DocumentFormat.OpenXml.Spreadsheet.SharedStringTable` object then it creates one. If the text already exists in the `ShareStringTable` object, then it returns the index for the `DocumentFormat.OpenXml.Spreadsheet.SharedStringItem` object that represents the text. If the text does not exist, create a new `SharedStringItem` object that represents the text. It then returns the index for the `SharedStringItem` object that represents the text.

### [C#](#tab/cs-5)
```csharp
// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
{
    // If the part does not contain a SharedStringTable, create it.
    if (shareStringPart.SharedStringTable is null)
    {
        shareStringPart.SharedStringTable = new SharedStringTable();
    }

    int i = 0;
    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
    {
        if (item.InnerText == text)
        {
            // The text already exists in the part. Return its index.
            return i;
        }

        i++;
    }

    // The text does not exist in the part. Create the SharedStringItem.
    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

    return i;
}
```
### [Visual Basic](#tab/vb-5)
```vb
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create it.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                ' The text already exists in the part. Return its index.
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
```
***

The final step is to insert a cell into the worksheet. The code does that by passing in parameters that represent the name of the column and the number of the row of the cell, and a parameter that represents the worksheet that contains the cell. If the specified row does not exist, it creates the row and append it to the worksheet. If the specified column exists, it finds the cell that matches the row in that column and returns the cell. If the specified column does not exist, it creates the column and inserts it into the worksheet. It then determines where to insert the new cell in the column by iterating through the row elements to find the cell that comes directly after the specified row, in sequential order. It saves this row in the `refCell` variable. It inserts the new cell before the cell referenced by `refCell` using the `DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore` method. It then returns the new `Cell` object.

### [C#](#tab/cs-6)
```csharp
// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
// If the cell already exists, returns it.
static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
{
    Worksheet worksheet = worksheetPart.Worksheet;
    SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
    string cellReference = columnName + rowIndex;

    // If the worksheet does not contain a row with the specified row index, insert one.
    Row row;
    if (sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
    {
        row = sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
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
### [Visual Basic](#tab/vb-6)
```vb
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)

            Return newCell
        End If
    End Function
```
***

## Sample Code
The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
        string? firstId = sheets?.First().Id;
        if (sheets is null || firstId is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(firstId);
        Worksheet worksheet = worksheetPart.Worksheet;

        // Get the row number and column name for the first and last cells in the range.
        uint firstRowNum = GetRowIndex(firstCellName);
        uint lastRowNum = GetRowIndex(lastCellName);
        string firstColumn = GetColumnName(firstCellName);
        string lastColumn = GetColumnName(lastCellName);

        double sum = 0;

        // Iterate through the cells within the range and add their values to the sum.
        foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex is not null && r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
        {
            foreach (Cell cell in row)
            {
                if (cell.CellReference is not null && cell.CellReference.Value is not null)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0 && double.TryParse(cell.CellValue?.Text, out double num))
                    {
                        sum += num;
                    }
                }
            }
        }

        // Get the SharedStringTablePart and add the result to it.
        // If the SharedStringPart does not exist, create a new one.
        SharedStringTablePart shareStringPart;
        if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }
        else
        {
            shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
        }

        // Insert the result into the SharedStringTablePart.
        int index = InsertSharedStringItem("Result: " + sum, shareStringPart);

        Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

        // Set the value of the cell.
        result.CellValue = new CellValue(index.ToString());
        result.DataType = new EnumValue<CellValues>(CellValues.SharedString);
    }
}
// Given a cell name, parses the specified cell to get the row index.
static uint GetRowIndex(string cellName)
{
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
}
// Given a cell name, parses the specified cell to get the column name.
static string GetColumnName(string cellName)
{
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
}
// Given two columns, compares the columns.
static int CompareColumn(string column1, string column2)
{
    if (column1.Length > column2.Length)
    {
        return 1;
    }
    else if (column1.Length < column2.Length)
    {
        return -1;
    }
    else
    {
        return string.Compare(column1, column2, true);
    }
}
// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
{
    // If the part does not contain a SharedStringTable, create it.
    if (shareStringPart.SharedStringTable is null)
    {
        shareStringPart.SharedStringTable = new SharedStringTable();
    }

    int i = 0;
    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
    {
        if (item.InnerText == text)
        {
            // The text already exists in the part. Return its index.
            return i;
        }

        i++;
    }

    // The text does not exist in the part. Create the SharedStringItem.
    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

    return i;
}
// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
// If the cell already exists, returns it.
static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
{
    Worksheet worksheet = worksheetPart.Worksheet;
    SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
    string cellReference = columnName + rowIndex;

    // If the worksheet does not contain a row with the specified row index, insert one.
    Row row;
    if (sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
    {
        row = sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
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
    Private Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String,
    ByVal lastCellName As String, ByVal resultCell As String)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) =
                document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If (sheets.Count() = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)
            Dim worksheet As Worksheet = worksheetPart.Worksheet

            ' Get the row number and column name for the first and last cells in the range.
            Dim firstRowNum As UInteger = GetRowIndex(firstCellName)
            Dim lastRowNum As UInteger = GetRowIndex(lastCellName)
            Dim firstColumn As String = GetColumnName(firstCellName)
            Dim lastColumn As String = GetColumnName(lastCellName)

            Dim sum As Double = 0

            ' Iterate through the cells within the range and add their values to the sum.
            For Each row As Row In worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value >= firstRowNum _
                                                                             AndAlso r.RowIndex.Value <= lastRowNum)
                For Each cell As Cell In row
                    Dim columnName As String = GetColumnName(cell.CellReference.Value)
                    If ((CompareColumn(columnName, firstColumn) >= 0) AndAlso (CompareColumn(columnName, lastColumn) <= 0)) Then
                        sum = (sum + Double.Parse(cell.CellValue.Text))
                    End If
                Next
            Next

            ' Get the SharedStringTablePart and add the result to it.
            ' If the SharedStringPart does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If (document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0) Then
                shareStringPart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = document.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the result into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(("Result:" + sum.ToString()), shareStringPart)

            Dim result As Cell = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart)

            ' Set the value of the cell.
            result.CellValue = New CellValue(index.ToString())
            result.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        End Using
    End Sub
    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function
    ' Given two columns, compares the columns.
    Private Function CompareColumn(ByVal column1 As String, ByVal column2 As String) As Integer
        If (column1.Length > column2.Length) Then
            Return 1
        ElseIf (column1.Length < column2.Length) Then
            Return -1
        Else
            Return String.Compare(column1, column2, True)
        End If
    End Function
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create it.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                ' The text already exists in the part. Return its index.
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)

            Return newCell
        End If
    End Function
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
