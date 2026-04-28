# Retrieve a list of the hidden rows or columns in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve a list of hidden rows or columns in a Microsoft Excel worksheet. It contains an example `GetHiddenRowsOrCols` method to illustrate this task.

---------------------------------------------------------------------------------

## GetHiddenRowsOrCols Method

You can use the `GetHiddenRowsOrCols` method to retrieve a list of the hidden rows or columns in a worksheet. The method returns a list of unsigned integers that contain each index for the hidden rows or columns, if the specified worksheet contains any hidden rows or columns (rows and columns are numbered starting at 1, rather than 0). The `GetHiddenRowsOrCols` method accepts three parameters:

- The name of the document to examine (string).

- The name of the sheet to examine (string).

- Whether to detect rows (true) or columns (false) (Boolean).

---------------------------------------------------------------------------------

## How the Code Works

The code opens the document, by using the <DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open*> method and indicating that the document should be open for read-only access (the final `false` parameter value). Next the code retrieves a reference to the workbook part, by using the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart` property of the document.

### [C#](#tab/cs-3)
```csharp
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        if (document is not null)
        {
            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
```

### [Visual Basic](#tab/vb-3)
```vb
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            If document IsNot Nothing Then
                Dim wbPart As WorkbookPart = If(document.WorkbookPart, document.AddWorkbookPart())
```
***

To find the hidden rows or columns, the code must first retrieve a reference to the specified sheet, given its name. This is not as easy as you might think. The code must look through all the sheet-type descendants of the workbook part's `DocumentFormat.OpenXml.Packaging.WorkbookPart.Workbook` property, examining the `DocumentFormat.OpenXml.Spreadsheet.Sheet.Name` property of each sheet that it finds.
Note that this search simply looks through the relations of the workbook, and does not actually find a worksheet part. It simply finds a reference to a `DocumentFormat.OpenXml.Spreadsheet.Sheet` object, which contains information such as the name and `DocumentFormat.OpenXml.Spreadsheet.Sheet.Id` property of the sheet. The simplest way to accomplish this is to use a LINQ query.

### [C#](#tab/cs-4)
```csharp
            Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault((s) => s.Name == sheetName);
```

### [Visual Basic](#tab/vb-4)
```vb
                Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)
```
***

The sheet information you already retrieved provides an `Id` property, and given that `Id` property, the code can retrieve a reference to the corresponding `DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart` property by calling the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.GetPartById` method of the `DocumentFormat.OpenXml.Packaging.WorkbookPart` object.

### [C#](#tab/cs-5)
```csharp
                // The sheet does exist.
                WorksheetPart? wsPart = wbPart.GetPartById(theSheet.Id!) as WorksheetPart;
                Worksheet? ws = wsPart?.Worksheet;
```

### [Visual Basic](#tab/vb-5)
```vb
                    ' The sheet does exist.
                    Dim wsPart As WorksheetPart = TryCast(wbPart.GetPartById(theSheet.Id), WorksheetPart)
                    Dim ws As Worksheet = wsPart?.Worksheet
```
***

---------------------------------------------------------------------------------

## Retrieving the List of Hidden Row or Column Index Values

The code uses the `detectRows` parameter that you specified when you called the method to determine whether to retrieve information about rows or columns.The code that actually retrieves the list of hidden rows requires only a single line of code.

### [C#](#tab/cs-7)
```csharp
                        // Retrieve hidden rows.
                        itemList = ws.Descendants<Row>()
                            .Where((r) => r?.Hidden is not null && r.Hidden.Value)
                            .Select(r => r.RowIndex?.Value)
                            .Cast<uint>()
                            .ToList();
```

### [Visual Basic](#tab/vb-7)
```vb
                            ' Retrieve hidden rows.
                            itemList = ws.Descendants(Of Row)() _
                                .Where(Function(r) r?.Hidden IsNot Nothing AndAlso r.Hidden.Value) _
                                .Select(Function(r) r.RowIndex?.Value) _
                                .Cast(Of UInteger)() _
                                .ToList()
```
***

Retrieving the list of hidden columns is a bit trickier, because Excel collapses groups of hidden columns into a single element, and provides `DocumentFormat.OpenXml.Spreadsheet.Column.Min` and `DocumentFormat.OpenXml.Spreadsheet.Column.Max` properties that describe the first and last columns in the group. Therefore, the code that retrieves the list of hidden columns starts the same as the code that retrieves hidden rows. However, it must iterate through the index values (looping each item in the collection of hidden columns, adding each index from the `Min` to the `Max` value, inclusively).

### [C#](#tab/cs-8)
```csharp
                        var cols = ws.Descendants<Column>().Where((c) => c?.Hidden is not null && c.Hidden.Value);

                        foreach (Column item in cols)
                        {
                            if (item.Min is not null && item.Max is not null)
                            {
                                for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                                {
                                    itemList.Add(i);
                                }
                            }
                        }
```

### [Visual Basic](#tab/vb-8)
```vb
                            Dim cols = ws.Descendants(Of Column)().Where(Function(c) c?.Hidden IsNot Nothing AndAlso c.Hidden.Value)

                            For Each item As Column In cols
                                If item.Min IsNot Nothing AndAlso item.Max IsNot Nothing Then
                                    For i As UInteger = item.Min.Value To item.Max.Value
                                        itemList.Add(i)
                                    Next
                                End If
                            Next
```
***

---------------------------------------------------------------------------------

## Sample Code

The following is the complete `GetHiddenRowsOrCols` code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static List<uint> GetHiddenRowsOrCols(string fileName, string sheetName, string detectRows = "false")
{
    // Given a workbook and a worksheet name, return 
    // either a list of hidden row numbers, or a list 
    // of hidden column numbers. If detectRows is true, return
    // hidden rows. If detectRows is false, return hidden columns. 
    // Rows and columns are numbered starting with 1.
    List<uint> itemList = new List<uint>();
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        if (document is not null)
        {
            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
            Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault((s) => s.Name == sheetName);
            if (theSheet is null || theSheet.Id is null)
            {
                throw new ArgumentException("sheetName");
            }
            else
            {
                // The sheet does exist.
                WorksheetPart? wsPart = wbPart.GetPartById(theSheet.Id!) as WorksheetPart;
                Worksheet? ws = wsPart?.Worksheet;
                if (ws is not null)
                {
                    if (detectRows.ToLower() == "true")
                    {
                        // Retrieve hidden rows.
                        itemList = ws.Descendants<Row>()
                            .Where((r) => r?.Hidden is not null && r.Hidden.Value)
                            .Select(r => r.RowIndex?.Value)
                            .Cast<uint>()
                            .ToList();
                    }
                    else
                    {
                        // Retrieve hidden columns.
                        var cols = ws.Descendants<Column>().Where((c) => c?.Hidden is not null && c.Hidden.Value);

                        foreach (Column item in cols)
                        {
                            if (item.Min is not null && item.Max is not null)
                            {
                                for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                                {
                                    itemList.Add(i);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    return itemList;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetHiddenRowsOrCols(fileName As String, sheetName As String, Optional detectRows As String = "false") As List(Of UInteger)
        ' Given a workbook and a worksheet name, return 
        ' either a list of hidden row numbers, or a list 
        ' of hidden column numbers. If detectRows is true, return
        ' hidden rows. If detectRows is false, return hidden columns. 
        ' Rows and columns are numbered starting with 1.
        Dim itemList As New List(Of UInteger)()
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            If document IsNot Nothing Then
                Dim wbPart As WorkbookPart = If(document.WorkbookPart, document.AddWorkbookPart())
                Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)
                If theSheet Is Nothing OrElse theSheet.Id Is Nothing Then
                    Throw New ArgumentException("sheetName")
                Else
                    ' The sheet does exist.
                    Dim wsPart As WorksheetPart = TryCast(wbPart.GetPartById(theSheet.Id), WorksheetPart)
                    Dim ws As Worksheet = wsPart?.Worksheet
                    If ws IsNot Nothing Then
                        If detectRows.ToLower() = "true" Then
                            ' Retrieve hidden rows.
                            itemList = ws.Descendants(Of Row)() _
                                .Where(Function(r) r?.Hidden IsNot Nothing AndAlso r.Hidden.Value) _
                                .Select(Function(r) r.RowIndex?.Value) _
                                .Cast(Of UInteger)() _
                                .ToList()
                        Else
                            ' Retrieve hidden columns.
                            Dim cols = ws.Descendants(Of Column)().Where(Function(c) c?.Hidden IsNot Nothing AndAlso c.Hidden.Value)

                            For Each item As Column In cols
                                If item.Min IsNot Nothing AndAlso item.Max IsNot Nothing Then
                                    For i As UInteger = item.Min.Value To item.Max.Value
                                        itemList.Add(i)
                                    Next
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End Using

        Return itemList
    End Function
```
***

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
