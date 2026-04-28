# Working with paragraphs

This topic discusses the Open XML SDK `DocumentFormat.OpenXml.Wordprocessing.Paragraph` class and how it relates to the
Open XML File Format WordprocessingML schema.

--------------------------------------------------------------------------------
## Paragraphs in WordprocessingML
The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the Open XML WordprocessingML element used to represent a
paragraph in a WordprocessingML document.

The most basic unit of block-level content within a WordprocessingML
document, paragraphs are stored using the `<p>` element. A paragraph
defines a distinct division of content that begins on a new line. A
paragraph can contain three pieces of information: optional paragraph
properties, inline content (typically runs), and a set of optional
revision IDs used to compare the content of two documents.

A paragraph's properties are specified via the `<pPr>` element. Some
examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control.

&copy; ISO/IEC 29500: 2016

The following table lists the most common Open XML SDK classes used when
working with paragraphs.

| **WordprocessingML element** | **Open XML SDK Class** |
|------------------------------|----------------------------|
|            **p**             |         Paragraph          |
|           **pPr**            |    ParagraphProperties     |
|            **r**             |            Run             |
|            **t**             |            Text            |

---------------------------------------------------------------------------------
## Paragraph Class
The Open XML SDK `DocumentFormat.OpenXml.Wordprocessing.Paragraph` class represents the paragraph
`<p>` element defined in the Open XML
File Format schema for WordprocessingML documents as discussed above.
Use the **Paragraph** object to manipulate
individual `<p>` elements in a
WordprocessingML document.

### ParagraphProperties Class

In WordprocessingML, a paragraph's properties are specified via the
paragraph properties `<pPr>` element.
Some examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control. The OXML SDK `DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties` class represents the
`<pPr>` element.

### Run Class

Paragraphs in a word-processing document most often contain text. In the
OXML File Format schema for WordprocessingML documents, the run `<r>` element is provided to demarcate a region of
text. The OXML SDK `DocumentFormat.OpenXml.Wordprocessing.Run` class represents the `<r>` element.

### Text Object

With the `<r>` element, the text `<t>` element is the container for the text that
makes up the document content. The OXML SDK `DocumentFormat.OpenXml.Wordprocessing.Text` class represents the `<t>` element.

--------------------------------------------------------------------------------
## Open XML SDK Code Example
The following code instantiates an Open XML SDK**Paragraph** object and then uses it to add text to
a WordprocessingML document.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void WriteToWordDoc(string filepath, string txt)
{
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        if (wordprocessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordprocessingDocument));
        }
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Add a paragraph with some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
    }
}
WriteToWordDoc(args[0], args[1]);
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    Sub Main(args As String())
        WriteToWordDoc(args(0), args(1))
    End Sub

    Public Sub WriteToWordDoc(ByVal filepath As String, ByVal txt As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument =
            WordprocessingDocument.Open(filepath, True)
            ' Assign a reference to the existing document body.
            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

            ' Add a paragraph with some text.            
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text(txt))
        End Using

    End Sub
End Module
```

When this code is run, the following XML is written to the
WordprocessingML document referenced in the code.

```xml
    <w:p>
      <w:r>
        <w:t>String from WriteToWordDoc method.</w:t>
      </w:r>
    </w:p>
```

--------------------------------------------------------------------------------
## See also

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[Working with runs](working-with-runs.md)  

[How to: Apply a style to a paragraph in a word processing document](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)  

[How to: Open and add text to a word processing document](how-to-open-and-add-text-to-a-word-processing-document.md)
