# Insert a new worksheet into a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to insert a new worksheet into a spreadsheet document
programmatically.

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

The code that calls the `Open` method is
shown in the following `using` statement.

### [C#](#tab/cs-1)
```csharp
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
```

### [Visual Basic](#tab/vb-1)
```vb
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
```
***

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

--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void InsertWorksheet(string docName)
{
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
    {
        WorkbookPart workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();
        // Add a blank WorksheetPart.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new worksheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = (sheets.Elements<Sheet>().Select(s => s.SheetId?.Value).Max() + 1) ?? (uint)sheets.Elements<Sheet>().Count() + 1;
        }

        // Give the new worksheet a name.
        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub InsertWorksheet(docName As String)
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim workbookPart As WorkbookPart = If(spreadSheet.WorkbookPart, spreadSheet.AddWorkbookPart())
            ' Add a blank WorksheetPart.
            Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            newWorksheetPart.Worksheet = New Worksheet(New SheetData())

            Dim sheets As Sheets = If(workbookPart.Workbook.GetFirstChild(Of Sheets)(), workbookPart.Workbook.AppendChild(New Sheets()))
            Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

            ' Get a unique ID for the new worksheet.
            Dim sheetId As UInteger = 1
            If sheets.Elements(Of Sheet)().Count() > 0 Then
                sheetId = sheets.Elements(Of Sheet)().Select(Function(s) s.SheetId?.Value).Max() + 1
            End If

            ' Give the new worksheet a name.
            Dim sheetName As String = "Sheet" & sheetId

            ' Append the new worksheet and associate it with the workbook.
            Dim sheet As New Sheet() With {
                .Id = relationshipId,
                .SheetId = sheetId,
                .Name = sheetName
            }
            sheets.Append(sheet)
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
