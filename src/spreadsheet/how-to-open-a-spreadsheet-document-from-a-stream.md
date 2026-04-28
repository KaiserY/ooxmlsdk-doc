# Open a spreadsheet document from a stream

This topic shows how to use the classes in the Open XML SDK for
Office to open a spreadsheet document from a stream programmatically.

---------------------------------------------------------------------------------
## When to Open From a Stream
If you have an application, such as Microsoft SharePoint Foundation
2010, that works with documents by using stream input/output, and you
want to use the Open XML SDK to work with one of the documents, this
is designed to be easy to do. This is especially true if the document
exists and you can open it using the Open XML SDK. However, suppose
that the document is an open stream at the point in your code where you
must use the SDK to work with it? That is the scenario for this topic.
The sample method in the sample code accepts an open stream as a
parameter and then adds text to the document behind the stream using the
Open XML SDK.

--------------------------------------------------------------------------------

## The SpreadsheetDocument Object

The basic document structure of a SpreadsheetML document consists of the
`DocumentFormat.OpenXml.Spreadsheet.Sheets` and `DocumentFormat.OpenXml.Spreadsheet.Sheet` elements, which reference the
worksheets in the `DocumentFormat.OpenXml.Spreadsheet.Workbook`. A separate XML file is created
for each `DocumentFormat.OpenXml.Spreadsheet.Worksheet`. For example, the SpreadsheetML
for a workbook that has two worksheets name MySheet1 and MySheet2 is
located in the Workbook.xml file and is as follows.

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
`DocumentFormat.OpenXml.Spreadsheet.SheetData`. `sheetData` represents the cell table and contains
one or more `DocumentFormat.OpenXml.Spreadsheet.Row` elements. A `row` contains one or more `DocumentFormat.OpenXml.Spreadsheet.Cell` elements. Each cell contains a `DocumentFormat.OpenXml.Spreadsheet.CellValue` element that represents the value
of the cell. For example, the SpreadsheetML for the first worksheet in a
workbook, that only has the value 100 in cell A1, is located in the
Sheet1.xml file and is as follows.

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
content that uses strongly-typed classes that correspond to
SpreadsheetML elements. You can find these classes in the `DocumentFormat.OpenXML.Spreadsheet` namespace. The
following table lists the class names of the classes that correspond to
the `workbook`, `sheets`, `sheet`, `worksheet`, and `sheetData` elements.

| **SpreadsheetML Element**|**Open XML SDK Class**|**Description** |
|--|--|--|
| `<workbook/>`|`DocumentFormat.OpenXml.Spreadsheet.Workbook`|The root element for the main document part. |
| `<sheets/>`|`DocumentFormat.OpenXml.Spreadsheet.Sheets`|The container for the block level structures such as sheet, fileVersion, and  |others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.
| `<sheet/>`|`DocumentFormat.OpenXml.Spreadsheet.Sheet`|A sheet that points to a sheet definition file. |
| `<worksheet/>`|`DocumentFormat.OpenXml.Spreadsheet.Worksheet`|A sheet definition file that contains the sheet data. |
| `<sheetData/>`|`DocumentFormat.OpenXml.Spreadsheet.SheetData`|The cell table, grouped together by rows. |
| `<row/>`|`DocumentFormat.OpenXml.Spreadsheet.Row`|A row in the cell table. |
| `<c/>`|`DocumentFormat.OpenXml.Spreadsheet.Cell`|A cell in a row. |
| `<v/>`|`DocumentFormat.OpenXml.Spreadsheet.CellValue`|The value of a cell. |

--------------------------------------------------------------------------------
## Generating the SpreadsheetML Markup to Add a Worksheet

When you have access to the body of the main document part, you add a
worksheet by calling `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.AddNewPart` method to
create a new `DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart`. The following code example
adds the new `WorksheetPart`.

### [C#](#tab/cs-2)
```csharp
            // Add a new worksheet.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
```

### [Visual Basic](#tab/vb-2)
```vb
                ' Add a new worksheet.
                Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                newWorksheetPart.Worksheet = New Worksheet(New SheetData())
```
***

--------------------------------------------------------------------------------
## Sample Code

In this example, the `OpenAndAddToSpreadsheetStream` method can be used
to open a spreadsheet document from an already open stream and append
some text to it. The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs-3)
```csharp
using (FileStream fileStream = new FileStream(args[0], FileMode.Open, FileAccess.ReadWrite))
{
    OpenAndAddToSpreadsheetStream(fileStream);
}
```

### [Visual Basic](#tab/vb-3)
```vb
        Using fileStream As New FileStream(args(0), FileMode.Open, FileAccess.ReadWrite)
            OpenAndAddToSpreadsheetStream(fileStream)
        End Using
```
***

Notice that the `OpenAddAndAddToSpreadsheetStream` method does not
close the stream passed to it. The calling code must do that manually
or with a `using` statement.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void OpenAndAddToSpreadsheetStream(Stream stream)
{
    // Open a SpreadsheetDocument based on a stream.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true))
    {

        if (spreadsheetDocument is not null)
        {
            // Get or create the WorkbookPart
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
            // Add a new worksheet.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            Workbook workbook = workbookPart.Workbook ?? new Workbook();

            if (workbookPart.Workbook is null)
            {
                workbookPart.Workbook = workbook;
            }

            Sheets sheets = workbook.GetFirstChild<Sheets>() ?? workbook.AppendChild(new Sheets());
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
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub OpenAndAddToSpreadsheetStream(stream As Stream)
        ' Open a SpreadsheetDocument based on a stream.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(stream, True)

            If spreadsheetDocument IsNot Nothing Then
                ' Get or create the WorkbookPart
                Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
                ' Add a new worksheet.
                Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                newWorksheetPart.Worksheet = New Worksheet(New SheetData())
                Dim workbook As Workbook = If(workbookPart.Workbook, New Workbook())

                If workbookPart.Workbook Is Nothing Then
                    workbookPart.Workbook = workbook
                End If

                Dim sheets As Sheets = If(workbook.GetFirstChild(Of Sheets)(), workbook.AppendChild(New Sheets()))
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
            End If
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
