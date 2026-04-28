# Open and add text to a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically open and add text to a Word processing
document.

--------------------------------------------------------------------------------
## How to Open and Add Text to a Document

The Open XML SDK helps you create Word processing document structure
and content using strongly-typed classes that correspond to `WordprocessingML` elements. This topic shows how
to use the classes in the Open XML SDK to open a Word processing
document and add text to it. In addition, this topic introduces the
basic document structure of a `WordprocessingML` document, the associated XML
elements, and their corresponding Open XML SDK classes.

--------------------------------------------------------------------------------
## Create a WordprocessingDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class represents a
Word document package. To open and work with a Word document, create an
instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class from the document. When you create the instance from the document,
you can then obtain access to the main document part that contains the
text of the document. The text in the main document part is represented
in the package as XML using `WordprocessingML` markup.

To create the class instance from the document you call one of the `Open` methods. Several are provided, each with a
different signature. The sample code in this topic uses the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)` method with a signature that requires two parameters. The first parameter takes a full
path string that represents the document to open. The second parameter
is either `true` or `false` and represents whether you want the file to
be opened for editing. Changes you make to the document will not be
saved if this parameter is `false`.

The following code example calls the `Open` method.

### [C#](#tab/cs-0)
```csharp
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {

        if (wordprocessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordprocessingDocument));
        }
```
### [Visual Basic](#tab/vb-0)
```vb
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)

            If wordprocessingDocument Is Nothing Then
                Throw New ArgumentNullException(NameOf(wordprocessingDocument))
            End If
```
***

When you have opened the Word document package, you can add text to the
main document part. To access the body of the main document part, create
any missing elements and assign a reference to the document body, 
as shown in the following code example.

### [C#](#tab/cs-1)
```csharp
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        mainDocumentPart.Document.Body ??= mainDocumentPart.Document.AppendChild(new Body());
        Body body = wordprocessingDocument.MainDocumentPart!.Document!.Body!;
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Assign a reference to the existing document body. 
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
## Generate the WordprocessingML Markup to Add the Text
When you have access to the body of the main document part, add text by
adding instances of the `DocumentFormat.OpenXml.Wordprocessing.Paragraph`, `DocumentFormat.OpenXml.Wordprocessing.Run`,
and `DocumentFormat.OpenXml.Wordprocessing.Text` classes. 
This generates the required WordprocessingML markup. The
following code example adds the paragraph.

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

--------------------------------------------------------------------------------
## Sample Code
The example `OpenAndAddTextToWordDocument`
method shown here can be used to open a Word document and append some
text using the Open XML SDK. To call this method, pass a full path
filename as the first parameter and the text to add as the second.

### [C#](#tab/cs-3)
```csharp
string file = args[0];
string txt = args[1];

OpenAndAddTextToWordDocument(args[0], args[1]);
```
### [Visual Basic](#tab/vb-3)
```vb
        Dim file As String = args(0)
        Dim txt As String = args(1)

        OpenAndAddTextToWordDocument(file, txt)
```
***

Following is the complete sample code in both C\# and Visual Basic.

Notice that the `OpenAndAddTextToWordDocument` method does not
include an explicit call to `Save`. That is
because the AutoSave feature is on by default and has not been disabled
in the call to the `Open` method through use
of `OpenSettings`.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void OpenAndAddTextToWordDocument(string filepath, string txt)
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
        mainDocumentPart.Document.Body ??= mainDocumentPart.Document.AppendChild(new Body());
        Body body = wordprocessingDocument.MainDocumentPart!.Document!.Body!;
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
    }
}
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    Sub Main(args As String())
        Dim file As String = args(0)
        Dim txt As String = args(1)

        OpenAndAddTextToWordDocument(file, txt)
    End Sub

    Public Sub OpenAndAddTextToWordDocument(ByVal filepath As String, ByVal txt As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)

            If wordprocessingDocument Is Nothing Then
                Throw New ArgumentNullException(NameOf(wordprocessingDocument))
            End If
            ' Assign a reference to the existing document body. 
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
    End Sub
End Module
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
