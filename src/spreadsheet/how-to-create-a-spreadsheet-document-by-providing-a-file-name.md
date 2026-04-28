# Create a spreadsheet document by providing a file name

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a spreadsheet document.

--------------------------------------------------------------------------------
## Creating a SpreadsheetDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` class represents an
Excel document package. To create an Excel document, create an instance
of the `SpreadsheetDocument` class and
populate it with parts. At a minimum, the document must have a workbook
part that serves as a container for the document, and at least one
worksheet part. The text is represented in the package as XML using
`SpreadsheetML` markup.

To create the class instance, call the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create`
method. Several `Create` methods are
provided, each with a different signature. The sample code in this topic
uses the `Create` method with a signature
that requires two parameters. The first parameter, `package`, takes a full
path string that represents the document that you want to create. The
second parameter, *type*, is a member of the `DocumentFormat.OpenXml.SpreadsheetDocumentType` enumeration. This
parameter represents the document type. For example, there are different
members of the `SpreadsheetDocumentType`
enumeration for add-ins, templates, workbooks, and macro-enabled
templates and workbooks.

> **Note**
> Select the appropriate `SpreadsheetDocumentType` and ensure that the persisted file has the correct, matching file name extension. If the `SpreadsheetDocumentType` does not match the file name extension, an error occurs when you open the file in Excel.

The following code example calls the `Create` method.

### [C#](#tab/cs-0)
```csharp
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
```
### [Visual Basic](#tab/vb-0)
```vb
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)
```
***

When you have created the Excel document package, you can add parts to
it. To add the workbook part you call the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.AddWorkbookPart`
method of the `SpreadsheetDocument` class.

### [C#](#tab/cs-100)
```csharp
        // Add a WorkbookPart to the document.
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
```
### [Visual Basic](#tab/vb-100)
```vb
            ' Add a WorkbookPart to the document.
            Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()
```
***

A workbook part must
have at least one worksheet. To add a worksheet, create a new `Sheet`. When you create a new `Sheet`, associate the `Sheet` with the `DocumentFormat.OpenXml.Spreadsheet.Workbook` by passing the `Id`, `SheetId` and `Name` parameters. Use the
`DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.GetIdOfPart` method to get the
`Id` of the `Sheet`. Then add the new sheet to the `Sheet` collection by calling the
`DocumentFormat.OpenXml.OpenXmlElement.Append` method of the `DocumentFormat.OpenXml.Spreadsheet.Sheets` class.

To create the basic document structure using the Open XML SDK, instantiate the `Workbook` class, assign it
to the `DocumentFormat.OpenXml.Spreadsheet.Workbook.WorkbookPart` property of the main document
part, and then add instances of the `DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart`, `Worksheet`, and `Sheet`. The following code example
creates a new worksheet, associates the worksheet, and appends the
worksheet to the workbook.

### [C#](#tab/cs-1)
```csharp
        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Add a WorksheetPart to the WorkbookPart.
            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            worksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Add Sheets to the Workbook.
            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())

            ' Append a new worksheet and associate it with the workbook.
            Dim sheet As New Sheet() With {
                .Id = workbookPart.GetIdOfPart(worksheetPart),
                .SheetId = 1,
                .Name = "mySheet"
            }
            sheets.Append(sheet)
```
***

--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void CreateSpreadsheetWorkbook(string filepath)
{
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
    {
        // Add a WorkbookPart to the document.
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub CreateSpreadsheetWorkbook(filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)
            ' Add a WorkbookPart to the document.
            Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()
            ' Add a WorksheetPart to the WorkbookPart.
            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            worksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Add Sheets to the Workbook.
            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())

            ' Append a new worksheet and associate it with the workbook.
            Dim sheet As New Sheet() With {
                .Id = workbookPart.GetIdOfPart(worksheetPart),
                .SheetId = 1,
                .Name = "mySheet"
            }
            sheets.Append(sheet)
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
