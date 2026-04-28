# Change text in a table in a word processing document

This topic shows how to use the Open XML SDK for Office to programmatically change text in a table in an existing word processing document.

## Open the Existing Document

To open an existing document, instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in the following `using` statement. In the same statement, open the word processing file at the specified `filepath` by using the `Open` method, with the Boolean parameter set to `true` to enable editing the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
```
### [Visual Basic](#tab/vb-0)
```vb
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
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

## The Structure of a Table

The basic document structure of a `WordProcessingML` document consists of the `document` and `body`
elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph
contains one or more `r` elements. The `r` stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more `t` elements. The `t` element contains a range of text.

The document might contain a table as in this example. A `table` is a set of paragraphs (and other block-level
content) arranged in `rows` and `columns`. Tables in `WordprocessingML` are defined via the `tbl` element, which is analogous to the HTML table tag. Consider an empty one-cell table (that is, a table with one row and
one column) and 1 point borders on all sides. This table is represented
by the following `WordprocessingML` code
example.

```xml
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="4 w:space="0" w:color="auto"/>
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

In the sample code, after you open the document in the `using` statement, you locate the first table in
the document. Then you locate the second row in the table by finding the
row whose index is 1. Next, you locate the third cell in that row whose
index is 2, as shown in the following code example.

### [C#](#tab/cs-1)
```csharp
        // Find the first table in the document.
        Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

        // Find the second row in the table.
        TableRow row = table.Elements<TableRow>().ElementAt(1);

        // Find the third cell in the row.
        TableCell cell = row.Elements<TableCell>().ElementAt(2);
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()

            ' Find the second row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)

            ' Find the third cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)
```
***

After you have located the target cell, you locate the first run in the
first paragraph of the cell and replace the text with the passed in
text. The following code example shows these actions.

### [C#](#tab/cs-2)
```csharp
        // Find the first paragraph in the table cell.
        Paragraph p = cell.Elements<Paragraph>().First();

        // Find the first run in the paragraph.
        Run r = p.Elements<Run>().First();

        // Set the text for the run.
        Text t = r.Elements<Text>().First();
        t.Text = txt;
```
### [Visual Basic](#tab/vb-2)
```vb
            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().First()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().First()

            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().First()
            t.Text = txt
```
***

## Change Text in a Cell in a Table

The following code example shows how to change the text in the specified
table cell in a word processing document. The code example expects that
the document, whose file name and path are passed as an argument to the
`ChangeTextInCell` method, contains a table.
The code example also expects that the table has at least two rows and
three columns, and that the table contains text in the cell that is
located at the second row and the third column position. When you call
the `ChangeTextInCell` method in your
program, the text in the cell at the specified location will be replaced
by the text that you pass in as the second argument to the `ChangeTextInCell` method.

| **Some text** | **Some text** | **Some text** |
|---------------|---------------|---------------|
| Some text     | Some text     |The text from the second argument |

## Sample Code

The `ChangeTextInCell` method changes the
text in the second row and the third column of the first table found in
the file. You call it by passing a full path to the file as the first
parameter, and the text to use as the second parameter. For example, the
following call to the `ChangeTextInCell`
method changes the text in the specified cell to "The text from the API
example."

### [C#](#tab/cs-3)
```csharp
ChangeTextInCell(args[0], args[1]);
```
### [Visual Basic](#tab/vb-3)
```vb
        ChangeTextInCell(args(0), args(1))
```
***

Following is the complete code example.

### [C#](#tab/cs)
```csharp
// Change the text in a table in a word processing document.
static void ChangeTextInCell(string filePath, string txt)
{
    // Use the file name and path passed in as an argument to 
    // open an existing document.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
    {
        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }
        // Find the first table in the document.
        Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

        // Find the second row in the table.
        TableRow row = table.Elements<TableRow>().ElementAt(1);

        // Find the third cell in the row.
        TableCell cell = row.Elements<TableCell>().ElementAt(2);
        // Find the first paragraph in the table cell.
        Paragraph p = cell.Elements<Paragraph>().First();

        // Find the first run in the paragraph.
        Run r = p.Elements<Run>().First();

        // Set the text for the run.
        Text t = r.Elements<Text>().First();
        t.Text = txt;
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' Change the text in a table in a word processing document.
    Public Sub ChangeTextInCell(ByVal filepath As String, ByVal txt As String)
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()

            ' Find the second row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)

            ' Find the third cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)
            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().First()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().First()

            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().First()
            t.Text = txt
        End Using
    End Sub
```
***

## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)

[How to: Change Text in a Table in a Word Processing Document](https://learn.microsoft.com/previous-versions/office/developer/office-2010/cc840870(v=office.14))

[Language-Integrated Query (LINQ)](https://learn.microsoft.com/previous-versions/bb397926(v=vs.140))

[Extension Methods (C\# Programming Guide)](https://learn.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/extension-methods)

[Extension Methods (Visual Basic)](https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/extension-methods)
