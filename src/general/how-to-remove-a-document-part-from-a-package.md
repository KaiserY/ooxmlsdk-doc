# Remove a document part from a package

This topic shows how to use the classes in the Open XML SDK for
Office to remove a document part (file) from a Wordprocessing document
programmatically.

--------------------------------------------------------------------------------
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

---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object

The code example starts with opening a package file by passing a file
name as an argument to one of the overloaded `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open` methods of the 
`DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
that takes a string and a Boolean value that specifies whether the file
should be opened in read/write mode or not. In this case, the Boolean
value is `true` specifying that the file
should be opened in read/write mode.

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

---------------------------------------------------------------------------------

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

--------------------------------------------------------------------------------
## Settings Element
The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the settings element in a `PresentationML` package.

> This element specifies the settings that are applied to a
> WordprocessingML document. This element is the root element of the
> Document Settings part in a WordprocessingML document.   
> **Example**:
> Consider the following WordprocessingML fragment for the settings part
> of a document:

```xml
    <w:settings>
      <w:defaultTabStop w:val="720" />
      <w:characterSpacingControl w:val="dontCompress" />
    </w:settings>
```

> The **settings** element contains all of the
> settings for this document. In this case, the two settings applied are
> automatic tab stop increments of 0.5" using the **defaultTabStop** element, and no character level
> white space compression using the **characterSpacingControl** element. 
> 
> &copy; ISO/IEC 29500: 2016

--------------------------------------------------------------------------------
## How the Sample Code Works

After you have opened the document, in the `using` statement, as a `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` object, you create a
reference to the `DocumentSettingsPart` part.
You can then check if that part exists, if so, delete that part from the
package. In this instance, the `settings.xml`
part is removed from the package.

### [C#](#tab/cs-2)
```csharp
        MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

        if (mainPart is not null && mainPart.DocumentSettingsPart is not null)
        {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
        }
```

### [Visual Basic](#tab/vb-2)
```vb
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            If mainPart IsNot Nothing AndAlso mainPart.DocumentSettingsPart IsNot Nothing Then
                mainPart.DeletePart(mainPart.DocumentSettingsPart)
            End If
```
***

--------------------------------------------------------------------------------
## Sample Code

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// To remove a document part from a package.
static void RemovePart(string document)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

        if (mainPart is not null && mainPart.DocumentSettingsPart is not null)
        {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' To remove a document part from a package.
    Sub RemovePart(document As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            If mainPart IsNot Nothing AndAlso mainPart.DocumentSettingsPart IsNot Nothing Then
                mainPart.DeletePart(mainPart.DocumentSettingsPart)
            End If
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
