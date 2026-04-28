# Get the contents of a document part from a package

This topic shows how to use the classes in the Open XML SDK for
Office to retrieve the contents of a document part in a Wordprocessing
document programmatically.

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

The code starts with opening a package file by passing a file name to
one of the overloaded `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open` methods (Visual Basic .NET Shared
method or C\# static method) of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class that takes a
string and a Boolean value that specifies whether the file should be
opened in read/write mode or not. In this case, the Boolean value is
`false` specifying that the file should be
opened in read-only mode to avoid accidental changes.

### [C#](#tab/cs-1)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
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
## Comments Element

In this how-to, you are going to work with comments. Therefore, it is
useful to familiarize yourself with the structure of the `<comments/>` element. The following information
from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification can be useful when working with this element.

This element specifies all of the comments defined in the current
document. It is the root element of the comments part of a
WordprocessingML document.Consider the following WordprocessingML
fragment for the content of a comments part in a WordprocessingML
document:

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

The **comments** element contains the single
comment specified by this document in this example.

&copy; ISO/IEC 29500: 2016

The following XML schema fragment defines the contents of this element.

```xml
    <complexType name="CT_Comments">
       <sequence>
           <element name="comment" type="CT_Comment" minOccurs="0" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

--------------------------------------------------------------------------------
## How the Sample Code Works

After you have opened the source file for reading, you create a `mainPart` object by instantiating the `MainDocumentPart`. Then you can create a reference
to the `WordprocessingCommentsPart` part of
the document.

### [C#](#tab/cs-2)
```csharp
static string GetCommentsFromDocument(string document)
{
    string? comments = null;
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
    {
        if (wordDoc is null)
        {
            throw new ArgumentNullException(nameof(wordDoc));
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
        WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
```

### [Visual Basic](#tab/vb-2)
```vb
    Function GetCommentsFromDocument(document As String) As String
        Dim comments As String = Nothing
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
            If wordDoc Is Nothing Then
                Throw New ArgumentNullException(NameOf(wordDoc))
            End If

            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())
            Dim WordprocessingCommentsPart As WordprocessingCommentsPart = If(mainPart.WordprocessingCommentsPart, mainPart.AddNewPart(Of WordprocessingCommentsPart)())
```
***

You can then use a `StreamReader` object to
read the contents of the `WordprocessingCommentsPart` part of the document
and return its contents.

### [C#](#tab/cs-3)
```csharp
        using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
        {
            comments = streamReader.ReadToEnd();
        }
    }

    return comments;
```

### [Visual Basic](#tab/vb-3)
```vb
            Using streamReader As New StreamReader(WordprocessingCommentsPart.GetStream())
                comments = streamReader.ReadToEnd()
            End Using
        End Using

        Return comments
```
***

--------------------------------------------------------------------------------
## Sample Code
The following code retrieves the contents of a `WordprocessingCommentsPart` part contained in a
`WordProcessing` document package. You can
run the program by calling the `GetCommentsFromDocument` method as shown in the
following example.

### [C#](#tab/cs-4)
```csharp
string document = args[0];
GetCommentsFromDocument(document);
```

### [Visual Basic](#tab/vb-4)
```vb
        Dim document As String = args(0)
        GetCommentsFromDocument(document)
```
***

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static string GetCommentsFromDocument(string document)
{
    string? comments = null;
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
    {
        if (wordDoc is null)
        {
            throw new ArgumentNullException(nameof(wordDoc));
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
        WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
        {
            comments = streamReader.ReadToEnd();
        }
    }

    return comments;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetCommentsFromDocument(document As String) As String
        Dim comments As String = Nothing
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
            If wordDoc Is Nothing Then
                Throw New ArgumentNullException(NameOf(wordDoc))
            End If

            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())
            Dim WordprocessingCommentsPart As WordprocessingCommentsPart = If(mainPart.WordprocessingCommentsPart, mainPart.AddNewPart(Of WordprocessingCommentsPart)())
            Using streamReader As New StreamReader(WordprocessingCommentsPart.GetStream())
                comments = streamReader.ReadToEnd()
            End Using
        End Using

        Return comments
    End Function
```
***

--------------------------------------------------------------------------------
## See also

[Open XML SDK class library
reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
