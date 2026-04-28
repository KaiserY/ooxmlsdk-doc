# Insert a comment into a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically add a comment to the first paragraph in a
word processing document.

--------------------------------------------------------------------------------
## Open the Existing Document for Editing
To open an existing document, instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in
the following `using` statement. In the same
statement, open the word processing file at the specified *filepath* by
using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)`
method, with the Boolean parameter set to `true` to enable
editing in the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
```
### [Visual Basic](#tab/vb-0)
```vb
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
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
## How the Sample Code Works

After you open the document, you can find the first paragraph to attach
a comment. The code finds the first paragraph by calling the `System.Linq.Enumerable.First%2A`
extension method on all the descendant elements of the document element
that are of type `DocumentFormat.OpenXml.Wordprocessing.Paragraph`. The `First` method is a member
of the `System.Linq.Enumerable` class. The `System.Linq.Enumerable` class
provides extension methods for objects that implement the `System.Collections.Generic.IEnumerable%601` interface.

### [C#](#tab/cs-1)
```csharp
        Paragraph firstParagraph = document.MainDocumentPart.Document.Descendants<Paragraph>().First();
        wordprocessingCommentsPart.Comments ??= new Comments();
        string id = "0";
```
### [Visual Basic](#tab/vb-1)
```vb
            Dim firstParagraph As Paragraph = document.MainDocumentPart.Document.Descendants(Of Paragraph)().First()
            Dim comments As Comments = Nothing
            Dim id As String = "0"
```
***

The code first determines whether a `DocumentFormat.OpenXml.Packaging.WordprocessingCommentsPart`
part exists. To do this, call the `DocumentFormat.OpenXml.Packaging.MainDocumentPart` generic method,
`GetPartsCountOfType`, and specify a kind of `WordprocessingCommentsPart`.

If a `WordprocessingCommentsPart` part exists, the code obtains a new `Id` value for
the `DocumentFormat.OpenXml.Wordprocessing.Comment` object that it will add to the
existing `WordprocessingCommentsPart` `DocumentFormat.OpenXml.Wordprocessing.Comments`
collection object. It does this by finding the highest `Id` attribute value
given to a `Comment` in the `Comments` collection object, incrementing the
value by one, and then storing that as the `Id` value.If no `WordprocessingCommentsPart` part exists, the code
creates one using the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart%2A`
method of the `DocumentFormat.OpenXml.Packaging.MainDocumentPart` object and then adds a
`Comments` collection object to it.

### [C#](#tab/cs-2)
```csharp
        if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Count() > 0)
        {
            if (wordprocessingCommentsPart.Comments.HasChildren)
            {
                // Obtain an unused ID.
                id = (wordprocessingCommentsPart.Comments.Descendants<Comment>().Select(e =>
                {
                    if (e.Id is not null && e.Id.Value is not null)
                    {
                        return int.Parse(e.Id.Value);
                    }
                    else
                    {
                        throw new ArgumentNullException("Comment id and/or value are null.");
                    }
                })
                    .Max() + 1).ToString();
            }
        }
```
### [Visual Basic](#tab/vb-2)
```vb
            If document.MainDocumentPart.GetPartsOfType(Of WordprocessingCommentsPart).Count() > 0 Then
                comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments
                If comments.HasChildren Then
                    ' Obtain an unused ID.
                    id = comments.Descendants(Of Comment)().[Select](Function(e) e.Id.Value).Max()
                End If
            Else
                ' No WordprocessingCommentsPart part exists, so add one to the package.
                Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.AddNewPart(Of WordprocessingCommentsPart)()
                commentPart.Comments = New Comments()
                comments = commentPart.Comments
            End If
```
***

The `Comment` and `Comments` objects represent comment and comments
elements, respectively, in the Open XML Wordprocessing schema. A `Comment`
must be added to a `Comments` object so the code first instantiates a
`Comments` object (using the string arguments `author`, `initials`,
and `comments` that were passed in to the `AddCommentOnFirstParagraph` method).

The comment is represented by the following WordprocessingML code
example. .

```xml
    <w:comment w:id="1" w:initials="User">
      ...
    </w:comment>
```

The code then appends the `Comment` to the `Comments` object. This
creates the required XML document object model (DOM) tree structure in
memory which consists of a `comments` parent element with `comment` child elements under
it.

### [C#](#tab/cs-3)
```csharp
        Paragraph p = new Paragraph(new Run(new Text(comment)));
        Comment cmt =
            new Comment()
            {
                Id = id,
                Author = author,
                Initials = initials,
                Date = DateTime.Now
            };
        cmt.AppendChild(p);
        wordprocessingCommentsPart.Comments.AppendChild(cmt);
```
### [Visual Basic](#tab/vb-3)
```vb
            Dim p As New Paragraph(New Run(New Text(comment)))
            Dim cmt As New Comment() With {.Id = id, .Author = author, .Initials = initials, .Date = DateTime.Now}
            cmt.AppendChild(p)
            comments.AppendChild(cmt)
```
***

The following WordprocessingML code example represents the content of a
comments part in a WordprocessingML document.

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

With the `Comment` object instantiated, the code associates the `Comment` with a range in
the Wordprocessing document. `DocumentFormat.OpenXml.Wordprocessing.CommentRangeStart` and
`DocumentFormat.OpenXml.Wordprocessing.CommentRangeEnd` objects correspond to the
`commentRangeStart` and `commentRangeEnd` elements in the Open XML Wordprocessing schema.
A `CommentRangeStart` object is given as the argument to the `DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore%2A`
method of the `DocumentFormat.OpenXml.Wordprocessing.Paragraph` object and a `CommentRangeEnd`
object is passed to the `DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertAfter%2A` method.
This creates a comment range that extends from immediately before the first character of the first paragraph
in the Wordprocessing document to immediately after the last character of the first paragraph.

A `DocumentFormat.OpenXml.Wordprocessing.CommentReference` object represents a
`commentReference` element in the Open XML Wordprocessing schema. A
commentReference links a specific comment in the `WordprocessingCommentsPart` part (the Comments.xml
file in the Wordprocessing package) to a specific location in the
document body (the `MainDocumentPart` part
contained in the Document.xml file in the Wordprocessing package). The
`id` attribute of the comment,
commentRangeStart, commentRangeEnd, and commentReference is the same for
a given comment, so the commentReference `id`
attribute must match the comment `id` attribute
value that it links to. In the sample, the code adds a `commentReference` element by using the API, and
instantiates a `CommentReference` object,
specifying the `Id` value, and then adds it to a `DocumentFormat.OpenXml.Wordprocessing.Run` object.

### [C#](#tab/cs-4)
```csharp
        firstParagraph.InsertBefore(new CommentRangeStart()
        { Id = id }, firstParagraph.GetFirstChild<Run>());

        // Insert the new CommentRangeEnd after last run of paragraph.
        var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd()
        { Id = id }, firstParagraph.Elements<Run>().Last());

        // Compose a run with CommentReference and insert it.
        firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
```
### [Visual Basic](#tab/vb-4)
```vb
            firstParagraph.InsertBefore(New CommentRangeStart() With {.Id = id}, firstParagraph.GetFirstChild(Of Run)())

            ' Insert the new CommentRangeEnd after last run of paragraph.
            Dim cmtEnd = firstParagraph.InsertAfter(New CommentRangeEnd() With {.Id = id}, firstParagraph.Elements(Of Run)().Last())

            ' Compose a run with CommentReference and insert it.
            firstParagraph.InsertAfter(New Run(New CommentReference() With {.Id = id}), cmtEnd)
```
***

--------------------------------------------------------------------------------
## Sample Code
The following code example shows how to create a comment and associate
it with a range in a word processing document. To call the method `AddCommentOnFirstParagraph` pass in the path of
the document, your name, your initials, and the comment text.

### [C#](#tab/cs-5)
```csharp
string fileName = args[0];
string author = args[1];
string initials = args[2];
string comment = args[3];

AddCommentOnFirstParagraph(fileName, author, initials, comment);
```
### [Visual Basic](#tab/vb-5)
```vb
        Dim fileName As String = args(0)
        Dim author As String = args(1)
        Dim initials As String = args(2)
        Dim comment As String = args(3)

        AddCommentOnFirstParagraph(fileName, author, initials, comment)
```
***

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// Insert a comment on the first paragraph.
static void AddCommentOnFirstParagraph(string fileName, string author, string initials, string comment)
{
    // Use the file name and path passed in as an 
    // argument to open an existing Wordprocessing document. 
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
```

### [Visual Basic](#tab/vb)
```vb
    ' Insert a comment on the first paragraph.
    Public Sub AddCommentOnFirstParagraph(ByVal fileName As String, ByVal author As String, ByVal initials As String, ByVal comment As String)
        ' Use the file name and path passed in as an 
        ' argument to open an existing Wordprocessing document. 
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)

- [Language-Integrated Query (LINQ)](https://learn.microsoft.com/previous-versions/bb397926(v=vs.140))

- [Extension Methods (C\# Programming Guide)](https://learn.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/extension-methods)

- [Extension Methods (Visual Basic)](https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/extension-methods)
