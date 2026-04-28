# Retrieve a list of the worksheets in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a list of the worksheets in a
Microsoft Excel workbook, without loading
the document into Excel. It contains an example `GetAllWorksheets` method to illustrate this task.

--------------------------------------------------------------------------------

## GetAllWorksheets Method

You can use the `GetAllWorksheets` method,
which is shown in the following code, to retrieve a list of the
worksheets in a workbook. The `GetAllWorksheets` method accepts a single
parameter, a string that indicates the path of the file that you want to
examine.

### [C#](#tab/cs-0)
```csharp
Sheets? sheets = GetAllWorksheets(args[0]);
```

### [Visual Basic](#tab/vb-0)
```vb
        Dim sheets As Sheets = GetAllWorksheets(args(0))
```
***

The method works with the workbook you specify, returning an instance of
the `DocumentFormat.OpenXml.Spreadsheet.Sheets` object, from which you can retrieve
a reference to each `DocumentFormat.OpenXml.Spreadsheet.Sheet` object.

--------------------------------------------------------------------------------

## Calling the GetAllWorksheets Method

To call the `GetAllWorksheets` method, pass
the required value, as shown in the following code.

### [C#](#tab/cs-1)
```csharp
Sheets? sheets = GetAllWorksheets(args[0]);
if (sheets is not null)
{
    foreach (Sheet sheet in sheets)
    {
        Console.WriteLine(sheet.Name);
    }
}
```

### [Visual Basic](#tab/vb-1)
```vb
        Dim sheets As Sheets = GetAllWorksheets(args(0))
        If sheets IsNot Nothing Then
            For Each sheet As Sheet In sheets
                Console.WriteLine(sheet.Name)
            Next
        End If
```
***

--------------------------------------------------------------------------------

## How the Code Works

The sample method, `GetAllWorksheets`,
creates a variable that will contain a reference to the `Sheets` collection of the workbook. At the end of
its work, the method returns the variable, which contains either a
reference to the `Sheets` collection, or
`null`/`Nothing` if there were no sheets (this cannot occur in a well-formed
workbook).

### [C#](#tab/cs-2)
```csharp
    Sheets? theSheets = null;
```

### [Visual Basic](#tab/vb-2)
```vb
        Dim theSheets As Sheets = Nothing
```
***

The code then continues by opening the document in read-only mode, and
retrieving a reference to the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart`.

### [C#](#tab/cs-3)
```csharp
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        theSheets = document?.WorkbookPart?.Workbook.Sheets;
```

### [Visual Basic](#tab/vb-3)
```vb
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            theSheets = document?.WorkbookPart?.Workbook.Sheets
```
***

To get access to the `DocumentFormat.OpenXml.Spreadsheet.Workbook` object, the code retrieves the value of the `DocumentFormat.OpenXml.Packaging.WorkbookPart.Workbook` property from the `WorkbookPart`, and then retrieves a reference to the `Sheets` object from the `DocumentFormat.OpenXml.Spreadsheet.Workbook.Sheets` property of the `Workbook`. The `Sheets` object contains the collection of `DocumentFormat.OpenXml.Spreadsheet.Sheet` objects that provide the method's return value.

### [C#](#tab/cs-4)
```csharp
        theSheets = document?.WorkbookPart?.Workbook.Sheets;
```

### [Visual Basic](#tab/vb-4)
```vb
            theSheets = document?.WorkbookPart?.Workbook.Sheets
```
***

--------------------------------------------------------------------------------

## Sample Code

The following is the complete `GetAllWorksheets` code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
```csharp
static Sheets? GetAllWorksheets(string fileName)
{
    Sheets? theSheets = null;
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        theSheets = document?.WorkbookPart?.Workbook.Sheets;
    }

    return theSheets;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetAllWorksheets(fileName As String) As Sheets
        Dim theSheets As Sheets = Nothing
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            theSheets = document?.WorkbookPart?.Workbook.Sheets
        End Using

        Return theSheets
    End Function
```
***

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
