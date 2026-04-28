# Add a new document part to a package

This topic shows how to use the classes in the Open XML SDK for Office to add a document part (file) to a word processing document programmatically.

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

## Get a WordprocessingDocument object

The code starts with opening a package file by passing a file name to one of the overloaded `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` methods of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` that takes a string and a Boolean value that specifies whether the file should be opened for editing or for read-only access. In this case, the Boolean value is `true` specifying that the file should be opened in read/write mode.

### [C#](#tab/cs-1)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
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

## How the sample code works

After opening the document for editing, in the `using` statement, as a `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` object, the code creates a reference to the `MainDocumentPart` part and adds a new custom XML part. It then reads the contents of the external
file that contains the custom XML and writes it to the `CustomXmlPart` part.

> **Note**
> To use the new document part in the document, add a link to the document part in the relationship part for the new part.

### [C#](#tab/cs-2)
```csharp
        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();

        CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        {
            myXmlPart.FeedData(stream);
        }
```

### [Visual Basic](#tab/vb-2)
```vb
            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())

            Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)

            Using stream As New FileStream(fileName, FileMode.Open)
                myXmlPart.FeedData(stream)
            End Using
```
***

## Sample code

The following code adds a new document part that contains custom XML from an external file and then populates the part. To call the `AddCustomXmlPart` method in your program, use the following example that modifies a file by adding a new document part to it.

### [C#](#tab/cs-3)
```csharp
string document = args[0];
string fileName = args[1];

AddNewPart(args[0], args[1]);
```

### [Visual Basic](#tab/vb-3)
```vb
        Dim document As String = args(0)
        Dim fileName As String = args(1)

        AddNewPart(document, fileName)
```
***

> **Note**
> Before you run the program, change the Word file extension from .docx to .zip, and view the content of the zip file. Then change the extension back to .docx and run the program. After running the program, change the file extension again to .zip and view its content. You will see an extra folder named &quot;customXML.&quot; This folder contains the XML file that represents the added part

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void AddNewPart(string document, string fileName)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();

        CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        {
            myXmlPart.FeedData(stream);
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub AddNewPart(document As String, fileName As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())

            Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)

            Using stream As New FileStream(fileName, FileMode.Open)
                myXmlPart.FeedData(stream)
            End Using
        End Using
    End Sub
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
