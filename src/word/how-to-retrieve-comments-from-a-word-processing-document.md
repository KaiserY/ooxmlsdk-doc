# Retrieve comments from a word processing document

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically retrieve the comments from the main document
part in a word processing document.

--------------------------------------------------------------------------------
## Open the Existing Document for Read-only Access
To open an existing document, instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in
the following `using` statement. In the same
statement, open the word processing file at the specified `fileName` by using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean,DocumentFormat.OpenXml.Packaging.OpenSettings)` method. To open the file for editing the Boolean parameter is set to `true`. In this example you just need to read the file; therefore, you can open the file for read-only access by setting
the Boolean parameter to `false`.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileName, false))
    {
        if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.WordprocessingCommentsPart is null)
        {
            throw new System.ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.");
        }
```
### [Visual Basic](#tab/vb-0)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)

            If wordDoc.MainDocumentPart Is Nothing Or wordDoc.MainDocumentPart.WordprocessingCommentsPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.")
            End If
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

--------------------------------------------------------------------------------
## Comments Element

The `comments` and `comment` elements are crucial to working with
comments in a word processing file. It is important in this code example
to familiarize yourself with those elements.

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the comments element.

> **comments (Comments Collection)**
> 
> This element specifies all of the comments defined in the current
> document. It is the root element of the comments part of a
> WordprocessingML document.
> 
> Consider the following WordprocessingML fragment for the content of a
> comments part in a WordprocessingML document:

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

> © ISO/IEC 29500: 2016

The following XML schema segment defines the contents of the comments
element.

```xml
    <complexType name="CT_Comments">
       <sequence>
           <element name="comment" type="CT_Comment" minOccurs="0" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

---------------------------------------------------------------------------------
## Comment Element

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the comment element.

> **comment (Comment Content)**
> 
> This element specifies the content of a single comment stored in the
> comments part of a WordprocessingML document.
> 
> If a comment is not referenced by document content via a matching
> **id** attribute on a valid use of the **commentReference** element,
> then it may be ignored when loading the document. If more than one
> comment shares the same value for the **id** attribute, then only one
> comment shall be loaded and the others may be ignored.
> 
> Consider a document with text with an annotated comment as follows:

![Document text with annotated comment](../media/w-comment01.gif)

> This comment is represented by the following WordprocessingML
> fragment.

```xml
    <w:comment w:id="1" w:initials="User">
      …
    </w:comment>
```
> The **comment** element specifies the presence of a single comment
> within the comments part.
> 
> © ISO/IEC 29500: 2016

  
The following XML schema segment defines the contents of the comment element.

```xml
    <complexType name="CT_Comment">
       <complexContent>
           <extension base="CT_TrackChange">
              <sequence>
                  <group ref="EG_BlockLevelElts" minOccurs="0" maxOccurs="unbounded"/>
              </sequence>
              <attribute name="initials" type="ST_String" use="optional"/>
           </extension>
       </complexContent>
    </complexType>
```

--------------------------------------------------------------------------------
## How the Sample Code Works
After you have opened the file for read-only access, you instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingCommentsPart` class. You can
then display the inner text of the `DocumentFormat.OpenXml.Wordprocessing.Comment` element.

### [C#](#tab/cs-1)
```csharp
        WordprocessingCommentsPart commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;

        if (commentsPart is not null && commentsPart.Comments is not null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                Console.WriteLine(comment.InnerText);
            }
        }
```
### [Visual Basic](#tab/vb-1)
```vb
            Dim commentsPart As WordprocessingCommentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart

            If commentsPart IsNot Nothing AndAlso commentsPart.Comments IsNot Nothing Then
                For Each comment As Comment In
                    commentsPart.Comments.Elements(Of Comment)()
                    Console.WriteLine(comment.InnerText)
                Next
            End If
```
***

--------------------------------------------------------------------------------
## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void GetCommentsFromDocument(string fileName)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileName, false))
    {
        if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.WordprocessingCommentsPart is null)
        {
            throw new System.ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.");
        }
        WordprocessingCommentsPart commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;

        if (commentsPart is not null && commentsPart.Comments is not null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                Console.WriteLine(comment.InnerText);
            }
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        GetCommentsFromDocument(args(0))
    End Sub

    Public Sub GetCommentsFromDocument(ByVal fileName As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)

            If wordDoc.MainDocumentPart Is Nothing Or wordDoc.MainDocumentPart.WordprocessingCommentsPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.")
            End If
            Dim commentsPart As WordprocessingCommentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart

            If commentsPart IsNot Nothing AndAlso commentsPart.Comments IsNot Nothing Then
                For Each comment As Comment In
                    commentsPart.Comments.Elements(Of Comment)()
                    Console.WriteLine(comment.InnerText)
                Next
            End If
        End Using
    End Sub
End Module
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
