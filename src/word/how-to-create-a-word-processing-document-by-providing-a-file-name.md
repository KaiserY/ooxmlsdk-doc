# Create a word processing document by providing a file name

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a word processing document.

--------------------------------------------------------------------------------
## Creating a WordprocessingDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class represents a
Word document package. To create a Word document, you create an instance
of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using
WordprocessingML markup.

To create the class instance you call the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(System.String,DocumentFormat.OpenXml.WordprocessingDocumentType)` 
method. Several `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create%2A` methods are provided, each with a
different signature. The sample code in this topic uses the `Create` method with a signature that requires two
parameters. The first parameter takes a full path string that represents
the document that you want to create. The second parameter is a member
of the `DocumentFormat.OpenXml.WordprocessingDocumentType` enumeration.
This parameter represents the type of document. For example, there is a
different member of the `WordProcessingDocumentType` enumeration for each
of document, template, and the macro enabled variety of document and
template.

> **Note**
> Carefully select the appropriate `WordProcessingDocumentType` and verify that the persisted file has the correct, matching file extension. If the `WordProcessingDocumentType` does not match the file extension, an error occurs when you open the file in Microsoft Word.

The code that calls the `Create` method is
part of a `using` statement followed by a
bracketed block, as shown in the following code example.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
```
### [Visual Basic](#tab/vb-0)
```vb
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)
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

Once you have created the Word document package, you can add parts to
it. To add the main document part you call the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddMainDocumentPart%2A` method of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class. Having done that,
you can set about adding the document structure and text.

--------------------------------------------------------------------------------

## Structure of a WordProcessingML Document

The basic document structure of a `WordProcessingML` document consists of the `document` and `body` elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph contains one or more `r` elements. The `r` stands for run, which is a region of text with a common set of properties, such as formatting. A run contains one or more `t` elements. The `t` element contains a range of text. The following code example shows the `WordprocessingML` markup for a document that contains the text "Example text."

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK, you can create document structure and content using strongly-typed classes that correspond to `WordprocessingML` elements. You will find these classes in the `DocumentFormat.OpenXml.Wordprocessing` namespace. The following table lists the class names of the classes that correspond to the `document`, `body`, `p`, `r`, and `t` elements.

| **WordprocessingML Element** | **Open XML SDK Class** | **Description** |
|---|---|---|
| `<document/>` | `DocumentFormat.OpenXml.Wordprocessing.Document` | The root element for the main document part. |
| `<body/>` | `DocumentFormat.OpenXml.Wordprocessing.Body` | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| `<p/>` | `DocumentFormat.OpenXml.Wordprocessing.Paragraph` | A paragraph. |
| `<r/>` | `DocumentFormat.OpenXml.Wordprocessing.Run` | A run. |
| `<t/>` | `DocumentFormat.OpenXml.Wordprocessing.Text` | A range of text. |

For more information about the overall structure of the parts and elements of a WordprocessingML document, see [Structure of a WordprocessingML document](structure-of-a-wordprocessingml-document.md).

--------------------------------------------------------------------------------
## Generating the WordprocessingML Markup

To create the basic document structure using the Open XML SDK, you
instantiate the `Document` class, assign it
to the `Document` property of the main
document part, and then add instances of the `Body`, `Paragraph`,
`Run` and `Text`
classes. This is shown in the sample code listing, and does the work of
generating the required WordprocessingML markup. While the code in the
sample listing calls the `AppendChild` method
of each class, you can sometimes make code shorter and easier to read by
using the technique shown in the following code example.

### [C#](#tab/cs-1)
```csharp
    mainPart.Document = new Document(
       new Body(
          new Paragraph(
             new Run(
                new Text("Create text in body - CreateWordprocessingDocument")))));
```

### [Visual Basic](#tab/vb-1)
```vb
    mainPart.Document = New Document(New Body(New Paragraph(New Run(New Text("Create text in body - CreateWordprocessingDocument")))))
```
***

--------------------------------------------------------------------------------
## Sample Code
The `CreateWordprocessingDocument` method can
be used to create a basic Word document. You call it by passing a full
path as the only parameter. The following code example creates the
Invoice.docx file in the Public Documents folder.

### [C#](#tab/cs-2)
```csharp
CreateWordprocessingDocument(args[0]);
```
### [Visual Basic](#tab/vb-2)
```vb
        CreateWordprocessingDocument(args(0))
```
***

The file extension, .docx, matches the type of file specified by the
**WordprocessingDocumentType.Document**
parameter in the call to the **Create** method.

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void CreateWordprocessingDocument(string filepath)
{
    // Create a document by supplying the filepath. 
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
        // Add a main document part. 
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

        // Create the document structure and add some text.
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
    }
```

### [Visual Basic](#tab/vb)
```vb
    Public Sub CreateWordprocessingDocument(ByVal filepath As String)
        ' Create a document by supplying the filepath.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)
            ' Add a main document part. 
            Dim mainPart As MainDocumentPart = wordDocument.AddMainDocumentPart()

            ' Create the document structure and add some text.
            mainPart.Document = New Document()
            Dim body As Body = mainPart.Document.AppendChild(New Body())
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Create text in body - CreateWordprocessingDocument"))
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
