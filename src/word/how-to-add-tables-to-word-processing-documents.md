# Add tables to word processing documents

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add a table to a word processing document. It contains an example `AddTable` method to illustrate this task.

## AddTable method

You can use the `AddTable` method to add a simple table to a word processing document. The `AddTable` method accepts two parameters, indicating the following:

- The name of the document to modify (string).

- A two-dimensional array of strings to insert into the document as a
    table.

### [C#](#tab/cs-0)
```csharp
static void AddTable(string fileName, string[,] data)
```
### [Visual Basic](#tab/vb-0)
```vb
    Sub AddTable(fileName As String, data As String(,))
```
***

## Call the AddTable method

The `AddTable` method modifies the document you specify, adding a table that contains the information in the two-dimensional array that you provide. To call the method, pass both of the parameter values, as shown in the following code.

### [C#](#tab/cs-1)
```csharp
string fileName = args[0];

AddTable(fileName, new string[,] {
    { "Hawaii", "HI" },
    { "California", "CA" },
    { "New York", "NY" },
    { "Massachusetts", "MA" }
});
```
### [Visual Basic](#tab/vb-1)
```vb
        Dim fileName As String = args(0)

        AddTable(fileName, New String(,) {
            {"Hawaii", "HI"},
            {"California", "CA"},
            {"New York", "NY"},
            {"Massachusetts", "MA"}
        })
```
***

## How the code works

The following code starts by opening the document, using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open` method and indicating that the document should be open for read/write access (the final `true` parameter value). Next the code retrieves a reference to the root element of the main document part, using the `DocumentFormat.OpenXml.Packaging.MainDocumentPart.Document` property of the`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart` of the word processing document.

### [C#](#tab/cs-2)
```csharp
        using (var document = WordprocessingDocument.Open(fileName, true))
        {
            if (document.MainDocumentPart is null || document.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            var doc = document.MainDocumentPart.Document;
```
### [Visual Basic](#tab/vb-2)
```vb
            Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
                If document.MainDocumentPart Is Nothing OrElse document.MainDocumentPart.Document.Body Is Nothing Then
                    Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
                End If

                Dim doc = document.MainDocumentPart.Document
```
***

## Create the table object and set its properties

Before you can insert a table into a document, you must create the `DocumentFormat.OpenXml.Wordprocessing.Table` object and set its properties. To set a table's properties, you create and supply values for a `DocumentFormat.OpenXml.Wordprocessing.TableProperties` object. The `TableProperties` class provides many table-oriented properties, like `DocumentFormat.OpenXml.Wordprocessing.TableProperties.Shading`, `DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableBorders`, `DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableCaption`, `DocumentFormat.OpenXml.Wordprocessing.TableCellProperties`, `DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableJustification`, and more. The sample method includes the following code.

### [C#](#tab/cs-3)
```csharp
            Table table = new();

            TableProperties props = new(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                }));

            table.AppendChild<TableProperties>(props);
```
### [Visual Basic](#tab/vb-3)
```vb
                Dim table As New Table()

                Dim props As New TableProperties(
                    New TableBorders(
                        New TopBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New BottomBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New LeftBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New RightBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideHorizontalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideVerticalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        }))

                table.AppendChild(Of TableProperties)(props)
```
***

The constructor for the `TableProperties` class allows you to specify as many child elements as you like (much like the `System.Xml.Linq.XElement` constructor). In this case, the code creates `DocumentFormat.OpenXml.Wordprocessing.TopBorder`, `DocumentFormat.OpenXml.Wordprocessing.BottomBorder`, `DocumentFormat.OpenXml.Wordprocessing.LeftBorder`, `DocumentFormat.OpenXml.Wordprocessing.RightBorder`, `DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder`, and `DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder` child elements, each describing one of the border elements for the table. For each element, the code sets the `Val` and `Size` properties as part of calling the constructor. Setting the size is simple, but setting the `Val` property requires a bit more effort: this property, for this particular object, represents the border style, and you must set it to an enumerated value. To do that, create an instance of the `DocumentFormat.OpenXml.EnumValue%601` generic type, passing the specific border type (`DocumentFormat.OpenXml.Wordprocessing.BorderValues`) as a parameter to the constructor. Once the code has set all the table border value it needs to set, it calls the `DocumentFormat.OpenXml.OpenXmlElement.AppendChild` method of the table, indicating that the generic type is `DocumentFormat.OpenXml.Wordprocessing.TableProperties` i.e., it is appending an instance of the `TableProperties` class, using the variable `props` as the value.

## Fill the table with data

Given that table and its properties, now it is time to fill the table with data. The sample procedure iterates first through all the rows of data in the array of strings that you specified, creating a new `DocumentFormat.OpenXml.Wordprocessing.TableRow` instance for each row of data. The following code shows how you create and append the row to the table. Then for each column, the code creates a new `DocumentFormat.OpenXml.Wordprocessing.TableCell` object, fills it with data, and appends it to the row. 

Next, the code does the following:

- Creates a new `DocumentFormat.OpenXml.Wordprocessing.Text` object that contains a value from the array of strings.
- Passes the `DocumentFormat.OpenXml.Wordprocessing.Text` object to the constructor for a new `DocumentFormat.OpenXml.Wordprocessing.Run` object.
- Passes the `DocumentFormat.OpenXml.Wordprocessing.Run` object to the constructor for a new `DocumentFormat.OpenXml.Wordprocessing.Paragraph` object.
- Passes the `DocumentFormat.OpenXml.Wordprocessing.Paragraph` object to the `DocumentFormat.OpenXml.OpenXmlElement.Append` method of the cell.

The code then appends a new `DocumentFormat.OpenXml.Wordprocessing.TableCellProperties` object to the cell. This `TableCellProperties` object, like the `TableProperties` object you already saw, can accept as many objects in its constructor as you care to supply. In this case, the code passes only a new `DocumentFormat.OpenXml.Wordprocessing.TableCellWidth` object, with its `DocumentFormat.OpenXml.Wordprocessing.TableWidthType.Type` property set to `DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues` (so that the table automatically sizes the width of each column).

### [C#](#tab/cs-4)
```csharp
            for (var i = 0; i < data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j < data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
```
### [Visual Basic](#tab/vb-4)
```vb
                For i As Integer = 0 To data.GetUpperBound(0) - 1
                    Dim tr As New TableRow()
                    For j As Integer = 0 To data.GetUpperBound(1) - 1
                        Dim tc As New TableCell()
                        tc.Append(New Paragraph(New Run(New Text(data(i, j)))))

                        ' Assume you want columns that are automatically sized.
                        tc.Append(New TableCellProperties(
                            New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))

                        tr.Append(tc)
                    Next
                    table.Append(tr)
                Next
```
***

## Finish up

The following code concludes by appending the table to the body of the document, and then saving the document.

### [C#](#tab/cs-8)
```csharp
            doc.Body.Append(table);
```
### [Visual Basic](#tab/vb-8)
```vb
                doc.Body.Append(table)
```
***

## Sample Code

The following is the complete **AddTable** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void AddTable(string fileName, string[,] data)
{
    if (data is not null)
    {
        using (var document = WordprocessingDocument.Open(fileName, true))
        {
            if (document.MainDocumentPart is null || document.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            var doc = document.MainDocumentPart.Document;
            Table table = new();

            TableProperties props = new(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                }));

            table.AppendChild<TableProperties>(props);
            for (var i = 0; i < data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j < data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
            doc.Body.Append(table);
        }
    }
}
```
### [Visual Basic](#tab/vb)
```vb
    Sub AddTable(fileName As String, data As String(,))
        If data IsNot Nothing Then
            Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
                If document.MainDocumentPart Is Nothing OrElse document.MainDocumentPart.Document.Body Is Nothing Then
                    Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
                End If

                Dim doc = document.MainDocumentPart.Document
                Dim table As New Table()

                Dim props As New TableProperties(
                    New TableBorders(
                        New TopBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New BottomBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New LeftBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New RightBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideHorizontalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideVerticalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        }))

                table.AppendChild(Of TableProperties)(props)
                For i As Integer = 0 To data.GetUpperBound(0) - 1
                    Dim tr As New TableRow()
                    For j As Integer = 0 To data.GetUpperBound(1) - 1
                        Dim tc As New TableCell()
                        tc.Append(New Paragraph(New Run(New Text(data(i, j)))))

                        ' Assume you want columns that are automatically sized.
                        tc.Append(New TableCellProperties(
                            New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))

                        tr.Append(tc)
                    Next
                    table.Append(tr)
                Next
                doc.Body.Append(table)
            End Using
        End If
    End Sub
End Module
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
