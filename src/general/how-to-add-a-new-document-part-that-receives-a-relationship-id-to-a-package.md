# Add a new document part that receives a relationship ID to a package

This topic shows how to use the classes in the Open XML SDK for
Office to add a document part (file) that receives a relationship `Id` parameter for a word
processing document.

-----------------------------------------------------------------------------
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

-----------------------------------------------------------------------------

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

-----------------------------------------------------------------------------

## How the Sample Code Works

The sample code, in this how-to, starts by passing in a parameter that represents the path to the Word document. It then creates
a new WordprocessingDocument object within a using statement.

### [C#](#tab/cs-1)
```csharp
static void AddNewPart(string document)
{
    // Create a new word processing document.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
```

### [Visual Basic](#tab/vb-1)
```vb
    Sub AddNewPart(document As String)
        ' Create a new word processing document.
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
```
***

It then adds the MainDocumentPart part in the new word processing document, with the relationship ID, rId1. It also adds the `CustomFilePropertiesPart` part and a `CoreFilePropertiesPart` in the new word processing document.

### [C#](#tab/cs-2)
```csharp
        // Add the MainDocumentPart part in the new word processing document.
        MainDocumentPart mainDocPart = wordDoc.AddMainDocumentPart();
        mainDocPart.Document = new Document();

        // Add the CustomFilePropertiesPart part in the new word processing document.
        var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
        customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Add the CoreFilePropertiesPart part in the new word processing document.
        var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
        using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" />
                """);
            writer.Flush();
        }
```

### [Visual Basic](#tab/vb-2)
```vb
            ' Add the MainDocumentPart part in the new word processing document.
            Dim mainDocPart As MainDocumentPart = wordDoc.AddMainDocumentPart()
            mainDocPart.Document = New Document()

            ' Add the CustomFilePropertiesPart part in the new word processing document.
            Dim customFilePropPart = wordDoc.AddCustomFilePropertiesPart()
            customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()

            ' Add the CoreFilePropertiesPart part in the new word processing document.
            Dim coreFilePropPart = wordDoc.AddCoreFilePropertiesPart()
            Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8)
                writer.WriteRaw("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                                "<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" />")
                writer.Flush()
            End Using
```
***

The code then adds the `DigitalSignatureOriginPart` part, the `ExtendedFilePropertiesPart` part, and the `ThumbnailPart` part in the new word processing document with realtionship IDs rId4, rId5, and rId6.

> **Note**
> The `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddNewPart` method creates a relationship from the current document part to the new document part. This method returns the new document part. Also, you can use the <DocumentFormat.OpenXml.Packaging.DataPart.FeedData*> method to fill the document part.

## Sample Code

The following code, adds a new document part that contains custom XML
from an external file and then populates the document part. Below is the
complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void AddNewPart(string document)
{
    // Create a new word processing document.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
    {
        // Add the MainDocumentPart part in the new word processing document.
        MainDocumentPart mainDocPart = wordDoc.AddMainDocumentPart();
        mainDocPart.Document = new Document();

        // Add the CustomFilePropertiesPart part in the new word processing document.
        var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
        customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Add the CoreFilePropertiesPart part in the new word processing document.
        var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
        using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" />
                """);
            writer.Flush();
        }
        // Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

        // Add the ExtendedFilePropertiesPart part in the new word processing document.
        var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
        extendedFilePropPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

        // Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub AddNewPart(document As String)
        ' Create a new word processing document.
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
            ' Add the MainDocumentPart part in the new word processing document.
            Dim mainDocPart As MainDocumentPart = wordDoc.AddMainDocumentPart()
            mainDocPart.Document = New Document()

            ' Add the CustomFilePropertiesPart part in the new word processing document.
            Dim customFilePropPart = wordDoc.AddCustomFilePropertiesPart()
            customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()

            ' Add the CoreFilePropertiesPart part in the new word processing document.
            Dim coreFilePropPart = wordDoc.AddCoreFilePropertiesPart()
            Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8)
                writer.WriteRaw("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                                "<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" />")
                writer.Flush()
            End Using
            ' Add the DigitalSignatureOriginPart part in the new word processing document.
            wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")

            ' Add the ExtendedFilePropertiesPart part in the new word processing document.
            Dim extendedFilePropPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
            extendedFilePropPart.Properties = New DocumentFormat.OpenXml.ExtendedProperties.Properties()

            ' Add the ThumbnailPart part in the new word processing document.
            wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")
        End Using
    End Sub
```
***

-----------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
