# Open a word processing document for read-only access

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically open a word processing document for read only
access.

---------------------------------------------------------------------------------
## When to Open a Document for Read-only Access

Sometimes you want to open a document to inspect or retrieve some
information, and you want to do so in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This how-to topic discusses several ways to
programmatically open a read-only word processing document.

--------------------------------------------------------------------------------
## Create a WordprocessingDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class represents a
Word document package. To work with a Word document, first create an
instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class from the document, and then work with that instance. Once you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using WordprocessingML markup.

To create the class instance from the document you call one of the `Open` methods. Several `Open` methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

Open Method|Class Library Reference Topic|Description
--|--|--
`Open(String, Boolean)`|`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)` |Create an instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class from the specified file.
`Open(Stream, Boolean)`|`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean)` |Create an instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class from the specified IO stream.
`Open(String, Boolean, OpenSettings)`|`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean,DocumentFormat.OpenXml.Packaging.OpenSettings)` |Create an instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class from the specified file.
`Open(Stream, Boolean, OpenSettings)`|`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean,DocumentFormat.OpenXml.Packaging.OpenSettings)` |Create an instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class from the specified I/O stream.

The table above lists only those `Open`
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read only access,
you specify false for this parameter.

Notice that two of the `Open` methods create
an instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class based on a string as the first parameter. The first example in the
sample code uses this technique. It uses the first `Open` method in the table above; with a signature
that requires two parameters. The first parameter takes a string that
represents the full path filename from which you want to open the
document. The second parameter is either `true` or `false`; this
example uses `false` and indicates whether
you want to open the file for editing.

The following code example calls the `Open`
Method.

### [C#](#tab/cs-0)
```csharp
    // Open a WordprocessingDocument based on a filepath.
    using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filepath, false))
```
### [Visual Basic](#tab/vb-0)
```vb
        ' Open a WordprocessingDocument based on a filepath.
        Using wordProcessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
```
***

The other two `Open` methods create an
instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class based on an input/output stream. You might employ this approach,
for instance, if you have a Microsoft SharePoint Online
application that uses stream input/output, and you want to use the Open
XML SDK to work with a document.

The following code example opens a document based on a stream.

### [C#](#tab/cs-1)
```csharp
    // Get a stream of the wordprocessing document
    using (FileStream fileStream = new FileStream(filepath, FileMode.Open))

    // Open a WordprocessingDocument for read-only access based on a stream.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(fileStream, false))
```
### [Visual Basic](#tab/vb-1)
```vb
        ' Get a stream of the wordprocessing document
        Using fileStream As FileStream = New FileStream(filepath, FileMode.Open)
            ' Open a WordprocessingDocument for read-only access based on a stream.
            Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(fileStream, False)
```
***

Suppose you have an application that employs the Open XML support in the
System.IO.Packaging namespace of the .NET Framework Class Library, and
you want to use the Open XML SDK to work with a package read only.
While the Open XML SDK includes method overloads that accept a `System.IO.Packaging.Package`
as the first parameter, there is not one that takes a Boolean as
the second parameter to indicate whether the document should be opened for editing.

The recommended method is to open the package as read-only to begin with
prior to creating the instance of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class, as shown in the
second example in the sample code. The following code example performs
this operation.

### [C#](#tab/cs-2)
```csharp
    // Open System.IO.Packaging.Package.
    using (Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read))
    // Open a WordprocessingDocument based on a package.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wordPackage))
```
### [Visual Basic](#tab/vb-2)
```vb
        ' Open System.IO.Packaging.Package.
        Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

        ' Open a WordprocessingDocument based on a package.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
```
***

Once you open the Word document package, you can access the main
document part. To access the body of the main document part, you assign
a reference to the existing document body, as shown in the following
code example.

### [C#](#tab/cs-3)
```csharp
        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordProcessingDocument.MainDocumentPart ?? wordProcessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Assign a reference to the document body. 
            Dim mainDocumentPart As MainDocumentPart = If(wordProcessingDocument.MainDocumentPart, wordProcessingDocument.AddMainDocumentPart())

            If wordProcessingDocument.MainDocumentPart.Document Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document = New Document()
            End If

            If wordProcessingDocument.MainDocumentPart.Document.Body Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document.Body = New Body()
            End If

            Dim body As Body = wordProcessingDocument.MainDocumentPart.Document.Body
```
***

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

For more information about the overall structure of the parts and elements of a WordprocessingML document, see [Structure of a WordprocessingML document](structure-of-a-wordprocessingml-document.md).

--------------------------------------------------------------------------------
## Generate the WordprocessingML Markup to Add Text and Attempt to Save
The sample code shows how you can add some text and attempt to save the
changes to show that access is read-only. Once you have access to the
body of the main document part, you add text by adding instances of the
`DocumentFormat.OpenXml.Wordprocessing.Paragraph`,
`DocumentFormat.OpenXml.Wordprocessing.Run`, and `DocumentFormat.OpenXml.Wordprocessing.Text`
classes. This generates the required WordprocessingML markup. The
following code example adds the paragraph, run, and text.

### [C#](#tab/cs-4)
```csharp
        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
```
### [Visual Basic](#tab/vb-4)
```vb
            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            'wordProcessingDocument.MainDocumentPart.Document.Save()
```
***

--------------------------------------------------------------------------------
## Sample Code
The first example method shown here, `OpenWordprocessingDocumentReadOnly`, opens a Word
document for read-only access. Call it by passing a full path to the
file that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-5)
```csharp
OpenWordprocessingDocumentReadonly(args[0]);
```
### [Visual Basic](#tab/vb-5)
```vb
        OpenWordprocessingDocumentReadonly(args(0))
```
***

The second example method, `OpenWordprocessingPackageReadonly`, shows how to
open a Word document for read-only access from a `System.IO.Packaging.Package`.
Call it by passing a full path to the file
that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-6)
```csharp
OpenWordprocessingPackageReadonly(args[0]);
```
### [Visual Basic](#tab/vb-6)
```vb
        OpenWordprocessingPackageReadonly(args(0))
```
***

The third example method, `OpenWordprocessingStreamReadonly`, shows how to
open a Word document for read-only access from a a stream.
Call it by passing a full path to the file
that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-6)
```csharp
OpenWordprocessingStreamReadonly(args[0]);
```
### [Visual Basic](#tab/vb-6)
```vb
        OpenWordprocessingStreamReadonly(args(0))
```
***

> **Important**
> If you uncomment the statement that saves the file, the program would throw an **IOException** because the file is opened for read-only access.

The following is the complete sample code in C\# and VB.

### [C#](#tab/cs)
```csharp
static void OpenWordprocessingDocumentReadonly(string filepath)
{
    // Open a WordprocessingDocument based on a filepath.
    using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filepath, false))
    {
        if (wordProcessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordProcessingDocument));
        }
        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordProcessingDocument.MainDocumentPart ?? wordProcessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
    }
}

static void OpenWordprocessingPackageReadonly(string filepath)
{
    // Open System.IO.Packaging.Package.
    using (Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read))
    // Open a WordprocessingDocument based on a package.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wordPackage))
    {
        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart ?? wordDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
    }
}

static void OpenWordprocessingStreamReadonly(string filepath)
{
    // Get a stream of the wordprocessing document
    using (FileStream fileStream = new FileStream(filepath, FileMode.Open))

    // Open a WordprocessingDocument for read-only access based on a stream.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(fileStream, false))
    {

        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart ?? wordDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingStreamReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Public Sub OpenWordprocessingDocumentReadonly(ByVal filepath As String)
        ' Open a WordprocessingDocument based on a filepath.
        Using wordProcessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
            ' Assign a reference to the document body. 
            Dim mainDocumentPart As MainDocumentPart = If(wordProcessingDocument.MainDocumentPart, wordProcessingDocument.AddMainDocumentPart())

            If wordProcessingDocument.MainDocumentPart.Document Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document = New Document()
            End If

            If wordProcessingDocument.MainDocumentPart.Document.Body Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document.Body = New Body()
            End If

            Dim body As Body = wordProcessingDocument.MainDocumentPart.Document.Body
            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            'wordProcessingDocument.MainDocumentPart.Document.Save()
        End Using
    End Sub

    Public Sub OpenWordprocessingPackageReadonly(ByVal filepath As String)
        ' Open System.IO.Packaging.Package.
        Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

        ' Open a WordprocessingDocument based on a package.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
            ' Assign a reference to the existing document body. 
            Dim body As Body = wordDocument.MainDocumentPart.Document.Body

            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            ' wordDocument.MainDocumentPart.Document.Save()
        End Using

        ' Close the package.
        wordPackage.Close()
    End Sub

    Public Sub OpenWordprocessingStreamReadonly(ByVal filepath As String)
        ' Get a stream of the wordprocessing document
        Using fileStream As FileStream = New FileStream(filepath, FileMode.Open)
            ' Open a WordprocessingDocument for read-only access based on a stream.
            Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(fileStream, False)
                ' Assign a reference to the existing document body. 
                Dim body As Body = wordDocument.MainDocumentPart.Document.Body

                ' Attempt to add some text.
                Dim para As Paragraph = body.AppendChild(New Paragraph())
                Dim run As Run = para.AppendChild(New Run())
                run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingStreamReadonly"))

                ' Call Save to generate an exception and show that access is read-only.
                ' wordDocument.MainDocumentPart.Document.Save()
            End Using
        End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
