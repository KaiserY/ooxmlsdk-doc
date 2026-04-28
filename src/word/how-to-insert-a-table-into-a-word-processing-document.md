# Insert a table into a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically insert a table into a word processing
document.

## Getting a WordprocessingDocument Object

To open an existing document, instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in the
following `using` statement. In the same
statement, open the word processing file at the specified filepath by
using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` method, with the Boolean
parameter set to `true` in order to enable
editing the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
```
### [Visual Basic](#tab/vb-0)
```vb
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
It ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement. Because the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class in the Open XML SDK
automatically saves and closes the object as part of its `System.IDisposable` implementation, and because
`System.IDisposable.Dispose` is automatically called when you
exit the block, you do not have to explicitly call `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save` or
`DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Dispose` as long as you use a `using` statement.

## Structure of a Table

The basic document structure of a `WordProcessingML` document consists of the `document` and `body`
elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph
contains one or more `r` elements. The r stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more `t` elements. The `t`
element contains a range of text.The document might contain a table as
in this example. A table is a set of paragraphs (and other block-level
content) arranged in rows and columns. Tables in `WordprocessingML`
are defined via the `tbl` element, which is analogous to the HTML table
tag. Consider an empty one-cell table (i.e. a table with one row, one
column) and 1 point borders on all sides. This table is represented by
the following `WordprocessingML` markup
segment.

```xml
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="10296"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="0" w:type="auto"/>
          </w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
```

This table specifies table-wide properties of 100% of page width using
the `tblW` element, a set of table borders
using the `tblBorders` element, the table
grid, which defines a set of shared vertical edges within the table
using the `tblGrid` element, and a single
table row using the `tr` element.

## How the Sample Code Works

In sample code, after you open the document in the `using` statement, you create a new
`DocumentFormat.OpenXml.Wordprocessing.Table` object. Then you create 
a `DocumentFormat.OpenXml.Wordprocessing.TableProperties` object and specify its border information.
The `DocumentFormat.OpenXml.Wordprocessing.TableProperties` class contains an overloaded 
constructor `DocumentFormat.OpenXml.Wordprocessing.TableProperties.%23ctor`
that takes a `params` array of type `DocumentFormat.OpenXml.OpenXmlElement`. The code uses this
constructor to instantiate a `TableProperties` object with `DocumentFormat.OpenXml.Wordprocessing.BorderType`
objects for each border, instantiating each `BorderType` and specifying its value using object initializers.
After it has been instantiated, append the `TableProperties` object to the table.

### [C#](#tab/cs-1)
```csharp
        // Create an empty table.
        Table table = new Table();

        // Create a TableProperties object and specify its border information.
        TableProperties tblProp = new TableProperties(
            new TableBorders(
                new TopBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new BottomBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new LeftBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new RightBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new InsideHorizontalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new InsideVerticalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                }
            )
        );

        // Append the TableProperties object to the empty table.
        table.AppendChild<TableProperties>(tblProp);
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Create an empty table.
            Dim table As New Table()

            ' Create a TableProperties object and specify its border information.
            Dim tblProp As New TableProperties(New TableBorders(
            New TopBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                },
            New BottomBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                },
            New LeftBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                },
            New RightBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                },
            New InsideHorizontalBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                },
            New InsideVerticalBorder() With
                {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Dashed),
                    .Size = 24
                })
            )
            ' Append the TableProperties object to the empty table.
            table.AppendChild(Of TableProperties)(tblProp)
```
***

The code creates a table row. This section of the code makes extensive
use of the overloaded `DocumentFormat.OpenXml.OpenXmlElement.Append%2A` methods,
which classes derived from `OpenXmlElement` inherit. The `Append` methods provide
a way to either append a single element or to append a portion of an XML tree,
to the end of the list of child elements under a given parent element. Next, the code
creates a `DocumentFormat.OpenXml.Wordprocessing.TableCell` object, which represents
an individual table cell, and specifies the width property of the table cell using a 
`DocumentFormat.OpenXml.Wordprocessing.TableCellProperties` object, and the cell
content ("Hello, World!") using a `DocumentFormat.OpenXml.Wordprocessing.Text` object.
In the Open XML Wordprocessing schema, a paragraph element (`<p\>`) contains run elements (`<r\>`)
which, in turn, contain text elements (`<t\>`). To insert text within a table cell using the API, you must create a
`DocumentFormat.OpenXml.Wordprocessing.Paragraph` object that contains a `DocumentFormat.OpenXml.Wordprocessing.Run`
object that contains a `Text` object that contains the text you want to insert in the cell.
You then append the `Paragraph` object to the `TableCell` object. This creates the proper XML
structure for inserting text into a cell. The `TableCell` is then appended to the
`DocumentFormat.OpenXml.Wordprocessing.TableRow` object.

### [C#](#tab/cs-2)
```csharp
        // Create a row.
        TableRow tr = new TableRow();

        // Create a cell.
        TableCell tc1 = new TableCell();

        // Specify the width property of the table cell.
        tc1.Append(new TableCellProperties(
            new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

        // Specify the table cell content.
        tc1.Append(new Paragraph(new Run(new Text("some text"))));

        // Append the table cell to the table row.
        tr.Append(tc1);
```
### [Visual Basic](#tab/vb-2)
```vb
            ' Create a row.
            Dim tr As New TableRow()

            ' Create a cell.
            Dim tc1 As New TableCell()

            ' Specify the width property of the table cell.
            tc1.Append(New TableCellProperties(New TableCellWidth()))

            ' Specify the table cell content.
            tc1.Append(New Paragraph(New Run(New Text("some text"))))

            ' Append the table cell to the table row.
            tr.Append(tc1)
```
***

The code then creates a second table cell. The final section of code creates another table cell
using the overloaded `TableCell` constructor `DocumentFormat.OpenXml.Wordprocessing.TableCell.%23ctor(System.String)`
that takes the `DocumentFormat.OpenXml.OpenXmlElement.OuterXml` property of an existing 
`TableCell` object as its only argument. After creating the second table cell, the code appends
the `TableCell` to the `TableRow`, appends the `TableRow` to the `Table`, and the `Table`
to the `DocumentFormat.OpenXml.Wordprocessing.Document` object.

### [C#](#tab/cs-3)
```csharp
        // Create a second table cell by copying the OuterXml value of the first table cell.
        TableCell tc2 = new TableCell(tc1.OuterXml);

        // Append the table cell to the table row.
        tr.Append(tc2);

        // Append the table row to the table.
        table.Append(tr);

        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Append the table to the document.
        doc.MainDocumentPart.Document.Body.Append(table);
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Create a second table cell by copying the OuterXml value of the first table cell.
            Dim tc2 As New TableCell(tc1.OuterXml)

            ' Append the table cell to the table row.
            tr.Append(tc2)

            ' Append the table row to the table.
            table.Append(tr)

            ' Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table)
```
***

## Sample Code

The following code example shows how to create a table, set its
properties, insert text into a cell in the table, copy a cell, and then
insert the table into a word processing document. You can invoke the
method `CreateTable` by using the following
call.

### [C#](#tab/cs-4)
```csharp
string filePath = args[0];

CreateTable(filePath);
```
### [Visual Basic](#tab/vb-4)
```vb
        Dim filePath As String = args(0)
        CreateTable(filePath)
```
***

After you run the program inspect the file to see the inserted table.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

// Insert a table into a word processing document.
static void CreateTable(string fileName)
{
    // Use the file name and path passed in as an argument 
    // to open an existing Word document.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder
Imports LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder
Imports RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder
Imports Run = DocumentFormat.OpenXml.Wordprocessing.Run
Imports Table = DocumentFormat.OpenXml.Wordprocessing.Table
Imports Text = DocumentFormat.OpenXml.Wordprocessing.Text
Imports TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder

Module MyModule

    Sub Main(args As String())
        Dim filePath As String = args(0)
        CreateTable(filePath)
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)

[Object Initializers: Named and Anonymous Types (Visual Basic .NET)](https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/objects-and-classes/object-initializers-named-and-anonymous-types)

[Object and Collection Initializers (C\# Programming Guide)](https://learn.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/object-and-collection-initializers)
