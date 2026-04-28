# Open a word processing document from a stream

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically open a Word processing document from a
stream.

## When to Open a Document from a Stream

If you have an application, such as a SharePoint application, that works
with documents using stream input/output, and you want to employ the
Open XML SDK to work with one of the documents, this is designed to
be easy to do. This is particularly true if the document exists and you
can open it using the Open XML SDK. However, suppose the document is
an open stream at the point in your code where you need to employ the
SDK to work with it? That is the scenario for this topic. The sample
method in the sample code accepts an open stream as a parameter and then
adds text to the document behind the stream using the Open XML SDK.

## Creating a WordprocessingDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class represents a
Word document package. To work with a Word document, first create an
instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class from the document, and then work with that instance. When you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using `WordprocessingML`
markup.

To create the class instance from the document call the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean)` method.Several `Open` methods are provided, each with a different
signature. The sample code in this topic uses the `Open` method with a signature that requires two
parameters. The first parameter takes a handle to the stream from which
you want to open the document. The second parameter is either `true` or `false` and
represents whether the stream is opened for editing.

The following code example calls the `Open`
method.

### [C#](#tab/cs-0)
```csharp
    // Open a WordProcessingDocument based on a stream.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true))
    {
```
### [Visual Basic](#tab/vb-0)
```vb
        ' Open a WordProcessingDocument based on a stream.
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, True)
```
***

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

## How the Sample Code Works

When you open the Word document package, you can add text to the main
document part. To access the body of the main document part you assign a
reference to the document body, as shown in the following code segment.

### [C#](#tab/cs-1)
```csharp
        // Assign a reference to the document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Assign a reference to the document body. 
            Dim mainDocumentPart As MainDocumentPart = If(wordprocessingDocument.MainDocumentPart, wordprocessingDocument.AddMainDocumentPart())

            If wordprocessingDocument.MainDocumentPart.Document Is Nothing Then
                wordprocessingDocument.MainDocumentPart.Document = New Document()
            End If

            If wordprocessingDocument.MainDocumentPart.Document.Body Is Nothing Then
                wordprocessingDocument.MainDocumentPart.Document.Body = New Body()
            End If

            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
```
***

When you access to the body of the main document part, add text by
adding instances of the `DocumentFormat.OpenXml.Wordprocessing.Paragraph`,
`DocumentFormat.OpenXml.Wordprocessing.Run`, and `DocumentFormat.OpenXml.Wordprocessing.Text`
classes. This generates the required `WordprocessingML` markup. The following lines from
the sample code add the paragraph, run, and text.

### [C#](#tab/cs-2)
```csharp
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
```
### [Visual Basic](#tab/vb-2)
```vb
            ' Add new text.
            Dim para As Paragraph = body.AppendChild(New Paragraph)
            Dim run As Run = para.AppendChild(New Run)
            run.AppendChild(New Text(txt))
```
***

## Sample Code

The example `OpenAndAddToWordprocessingStream` method shown
here can be used to open a Word document from an already open stream and
append some text using the Open XML SDK. You can call it by passing a
handle to an open stream as the first parameter and the text to add as
the second. For example, the following code example opens the
file specified in the first argument and adds text from the second argument to it.

### [C#](#tab/cs-3)
```csharp
string filePath = args[0];
string txt = args[1];

using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
{
    OpenAndAddToWordprocessingStream(fileStream, txt);
}
```
### [Visual Basic](#tab/vb-3)
```vb
        Dim filePath As String = args(0)
        Dim txt As String = args(1)

        Using fileStream As FileStream = New FileStream(filePath, FileMode.Open)
            OpenAndAddToWordprocessingStream(fileStream, txt)
        End Using
```
***

> **Note**
> Notice that the `OpenAddAddToWordprocessingStream` method does not close the stream passed to it. The calling code must do that
> by wrapping the method call in a `using` statement or explicitly calling Dispose.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
{
    // Open a WordProcessingDocument based on a stream.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true))
    {
        // Assign a reference to the document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
    }
    // Caller must close the stream.
}
```

### [Visual Basic](#tab/vb)
```vb
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    Sub Main(args As String())
        Dim filePath As String = args(0)
        Dim txt As String = args(1)

        Using fileStream As FileStream = New FileStream(filePath, FileMode.Open)
            OpenAndAddToWordprocessingStream(fileStream, txt)
        End Using
    End Sub

    Public Sub OpenAndAddToWordprocessingStream(ByVal stream As Stream, ByVal txt As String)
        ' Open a WordProcessingDocument based on a stream.
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, True)
            ' Assign a reference to the document body. 
            Dim mainDocumentPart As MainDocumentPart = If(wordprocessingDocument.MainDocumentPart, wordprocessingDocument.AddMainDocumentPart())

            If wordprocessingDocument.MainDocumentPart.Document Is Nothing Then
                wordprocessingDocument.MainDocumentPart.Document = New Document()
            End If

            If wordprocessingDocument.MainDocumentPart.Document.Body Is Nothing Then
                wordprocessingDocument.MainDocumentPart.Document.Body = New Body()
            End If

            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
            ' Add new text.
            Dim para As Paragraph = body.AppendChild(New Paragraph)
            Dim run As Run = para.AppendChild(New Run)
            run.AppendChild(New Text(txt))
        End Using
        ' Caller must close the stream.
    End Sub
End Module
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
