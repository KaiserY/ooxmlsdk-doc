# Parse and read a large spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically read a large Excel file. For more information
about the basic structure of a `SpreadsheetML` document, see [Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md).

> **Note**
> Interested in developing solutions that extend the Office experience across multiple platforms? Check out the new [Office Add-ins model](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). Office Add-ins have a small footprint compared to VSTO Add-ins and solutions, and you can build them by using almost any web programming technology, such as HTML5, JavaScript, CSS3, and XML.

--------------------------------------------------------------------------------
## Approaches to Parsing Open XML Files

The Open XML SDK provides two approaches to parsing Open XML files. You
can use the SDK Document Object Model (DOM), or the Simple API for XML
(SAX) reading and writing features. The SDK DOM is designed to make it
easy to query and parse Open XML files by using strongly-typed classes.
However, the DOM approach requires loading entire Open XML parts into
memory, which can cause an `Out of Memory`
exception when you are working with really large files. Using the SAX
approach, you can employ an OpenXMLReader to read the XML in the file
one element at a time, without having to load the entire file into
memory. Consider using SAX when you need to handle very large files.

The following code segment is used to read a very large Excel file using
the DOM approach.

### [C#](#tab/cs-2)
```csharp
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        string? text;

        foreach (Row r in sheetData.Elements<Row>())
        {
            foreach (Cell c in r.Elements<Cell>())
            {
                text = c?.CellValue?.Text;
                Console.Write(text + " ");
            }
        }
```

### [Visual Basic](#tab/vb-2)
```vb
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
            Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()
            Dim text As String = Nothing

            For Each r As Row In sheetData.Elements(Of Row)()
                For Each c As Cell In r.Elements(Of Cell)()
                    text = c?.CellValue?.Text
                    Console.Write(text & " ")
                Next
            Next
```
***

The following code segment performs an identical task to the preceding
sample (reading a very large Excel file), but uses the SAX approach.
This is the recommended approach for reading very large files.

### [C#](#tab/cs-3)
```csharp
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

        OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
        string text;
        while (reader.Read())
        {
            if (reader.ElementType == typeof(CellValue))
            {
                text = reader.GetText();
                Console.Write(text + " ");
            }
        }
```

### [Visual Basic](#tab/vb-3)
```vb
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

            Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
            Dim text As String
            While reader.Read()
                If reader.ElementType = GetType(CellValue) Then
                    text = reader.GetText()
                    Console.Write(text & " ")
                End If
            End While
```
***

--------------------------------------------------------------------------------
## Sample Code

You can imagine a scenario where you work for a financial company that
handles very large Excel spreadsheets. Those spreadsheets are updated
daily by analysts and can easily grow to sizes exceeding hundreds of
megabytes. You need a solution to read and extract relevant data from
every spreadsheet. The following code example contains two methods that
correspond to the two approaches, DOM and SAX. The latter technique will
avoid memory exceptions when using very large files. To try them, you
can call them in your code one after the other or you can call each
method separately by commenting the call to the one you would like to
exclude.

### [C#](#tab/cs-4)
```csharp
// Comment one of the following lines to test the method separately.
ReadExcelFileDOM(args[0]);    // DOM
ReadExcelFileSAX(args[0]);    // SAX
```

### [Visual Basic](#tab/vb-4)
```vb
        ' Comment one of the following lines to test the method separately.
        ReadExcelFileDOM(args(0))    ' DOM
        ReadExcelFileSAX(args(0))    ' SAX
```
***

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// The DOM approach.
// Note that the code below works only for cells that contain numeric values
static void ReadExcelFileDOM(string fileName)
{
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        string? text;

        foreach (Row r in sheetData.Elements<Row>())
        {
            foreach (Cell c in r.Elements<Cell>())
            {
                text = c?.CellValue?.Text;
                Console.Write(text + " ");
            }
        }
        Console.WriteLine();
        Console.ReadKey();
    }
}

// The SAX approach.
static void ReadExcelFileSAX(string fileName)
{
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

        OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
        string text;
        while (reader.Read())
        {
            if (reader.ElementType == typeof(CellValue))
            {
                text = reader.GetText();
                Console.Write(text + " ");
            }
        }
        Console.WriteLine();
        Console.ReadKey();
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' The DOM approach.
    ' Note that the code below works only for cells that contain numeric values
    Sub ReadExcelFileDOM(fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
            Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()
            Dim text As String = Nothing

            For Each r As Row In sheetData.Elements(Of Row)()
                For Each c As Cell In r.Elements(Of Cell)()
                    text = c?.CellValue?.Text
                    Console.Write(text & " ")
                Next
            Next
            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub

    ' The SAX approach.
    Sub ReadExcelFileSAX(fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

            Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
            Dim text As String
            While reader.Read()
                If reader.ElementType = GetType(CellValue) Then
                    text = reader.GetText()
                    Console.Write(text & " ")
                End If
            End While
            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also

[Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md)

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
