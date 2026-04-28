# Create a package

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a word processing document package
from content in the form of `WordprocessingML` XML markup.

## Packages and Document Parts 

An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500](https://www.iso.org/standard/71691.html). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.

## Getting a WordprocessingDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class represents a Word document package. To create a Word document, you create an instance
of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using `WordprocessingML` markup.

To create the class instance you call `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(System.String,DocumentFormat.OpenXml.WordprocessingDocumentType)`. Several `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create%2A` methods are
provided, each with a different signature. The first parameter takes a full path
string that represents the document that you want to create. The second
parameter is a member of the `DocumentFormat.OpenXml.WordprocessingDocumentType` enumeration.
This parameter represents the type of document. For example, there is a
different member of the `DocumentFormat.OpenXml.WordprocessingDocumentType` enumeration for each
of document, template, and the macro enabled variety of document and
template.

> **Note**
> Carefully select the appropriate `DocumentFormat.OpenXml.WordprocessingDocumentType` and verify that the persisted file has the correct, matching file extension. If the `DocumentFormat.OpenXml.WordprocessingDocumentType` does not match the file extension, an error occurs when you open the file in Microsoft Word.

### [C#](#tab/cs-1)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
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
it. To add the main document part you call `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddMainDocumentPart%2A`. Having done that,
you can set about adding the document structure and text.

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

For more information about the overall structure of the parts and elements of a WordprocessingML document, see [Structure of a WordprocessingML document](../word/structure-of-a-wordprocessingml-document.md).

## Sample Code

The following is the complete code sample that you can use to create an
Open XML word processing document package from XML content in the form
of `WordprocessingML` markup. 

After you run the program, open the created file and
examine its content; it should be one paragraph that contains the phrase
"Hello world!"

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// To create a new package as a Word document.
static void CreateNewWordDocument(string document)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
    {
        // Set the content of the document so that Word can open it.
        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

        SetMainDocumentContent(mainPart);
    }
}

// Set the content of MainDocumentPart.
static void SetMainDocumentContent(MainDocumentPart part)
{
    const string docXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
                            <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                              <w:body>
                                <w:p>
                                  <w:r>
                                    <w:t>Hello World</w:t>
                                  </w:r>
                                </w:p>
                              </w:body>
                            </w:document>";

    using (Stream stream = part.GetStream())
    {
        byte[] buf = (new UTF8Encoding()).GetBytes(docXml);
        stream.Write(buf, 0, buf.Length);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' To create a new package as a Word document.
    Sub CreateNewWordDocument(document As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
            ' Set the content of the document so that Word can open it.
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart()

            SetMainDocumentContent(mainPart)
        End Using
    End Sub

    ' Set the content of MainDocumentPart.
    Sub SetMainDocumentContent(part As MainDocumentPart)
        Const docXml As String = "<?xml version=""1.0"" encoding=""utf-8""?>" &
                                 "<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" &
                                 "<w:body>" &
                                 "<w:p>" &
                                 "<w:r>" &
                                 "<w:t>Hello World</w:t>" &
                                 "</w:r>" &
                                 "</w:p>" &
                                 "</w:body>" &
                                 "</w:document>"

        Using stream As Stream = part.GetStream()
            Dim buf As Byte() = (New UTF8Encoding()).GetBytes(docXml)
            stream.Write(buf, 0, buf.Length)
        End Using
    End Sub
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
