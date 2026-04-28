# Delete comments by all or a specific author in a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically delete comments by all or a specific author
in a word processing document, without having to load the document into
Microsoft Word. It contains an example `DeleteComments` method to illustrate this task.

--------------------------------------------------------------------------------

## DeleteComments Method

You can use the `DeleteComments` method to
delete all of the comments from a word processing document, or only
those written by a specific author. As shown in the following code, the
method accepts two parameters that indicate the name of the document to
modify (string) and, optionally, the name of the author whose comments
you want to delete (string). If you supply an author name, the code
deletes comments written by the specified author. If you do not supply
an author name, the code deletes all comments.

### [C#](#tab/cs-0)
```csharp
// Delete comments by a specific author. Pass an empty string for the 
// author to delete all comments, by all authors.
static void DeleteComments(string fileName, string author = "")
```
### [Visual Basic](#tab/vb-0)
```vb
    ' Delete comments by a specific author. Pass an empty string for the author 
    ' to delete all comments, by all authors.
    Public Sub DeleteComments(ByVal fileName As String, Optional ByVal author As String = "")
```
***

--------------------------------------------------------------------------------

## Calling the DeleteComments Method

To call the `DeleteComments` method, provide
the required parameters as shown in the following code.

### [C#](#tab/cs-1)
```csharp
if (args is [{ } fileName, { } author])
{
    DeleteComments(fileName, author);
}
else if (args is [{ } fileName2])
{
    DeleteComments(fileName2);
}
```
### [Visual Basic](#tab/vb-1)
```vb
        If (args.Length >= 2) Then
            Dim fileName As String = args(0)
            Dim author As String = args(1)

            DeleteComments(fileName, author)
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            DeleteComments(fileName)
        End If
```
***

--------------------------------------------------------------------------------
## How the Code Works

The following code starts by opening the document, using the
`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A?displayProperty=nameWithType`
method and indicating that the document should be open for read/write access (the
final `true` parameter value). Next, the code retrieves a reference to the comments
part, using the `DocumentFormat.OpenXml.Packaging.MainDocumentPart.WordprocessingCommentsPart`
property of the main document part, after having retrieved a reference to the main
document part from the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart`
property of the word processing document. If the comments part is missing, there is no point
in proceeding, as there cannot be any comments to delete.

### [C#](#tab/cs-2)
```csharp
    // Get an existing Wordprocessing document.
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
    {

        if (document.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Set commentPart to the document WordprocessingCommentsPart, 
        // if it exists.
        WordprocessingCommentsPart? commentPart = document.MainDocumentPart.WordprocessingCommentsPart;

        // If no WordprocessingCommentsPart exists, there can be no 
        // comments. Stop execution and return from the method.
        if (commentPart is null)
        {
            return;
        }
```
### [Visual Basic](#tab/vb-2)
```vb
        ' Get an existing Wordprocessing document.
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Set commentPart to the document 
            ' WordprocessingCommentsPart, if it exists.
            Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.WordprocessingCommentsPart

            ' If no WordprocessingCommentsPart exists, there can be no
            ' comments. Stop execution and return from the method.
            If (commentPart Is Nothing) Then
                Return
            End If
```
***

--------------------------------------------------------------------------------

## Creating the List of Comments

The code next performs two tasks: creating a list of all the comments to
delete, and creating a list of comment IDs that correspond to the
comments to delete. Given these lists, the code can both delete the
comments from the comments part that contains the comments, and delete
the references to the comments from the document part.The following code
starts by retrieving a list of `DocumentFormat.OpenXml.Wordprocessing.Comment`
elements. To retrieve the list, it converts the `DocumentFormat.OpenXml.OpenXmlElement.Elements`
collection exposed by the `commentPart` variable into a list of `Comment` objects.

### [C#](#tab/cs-3)
```csharp
        List<Comment> commentsToDelete = commentPart.Comments.Elements<Comment>().ToList();
```
### [Visual Basic](#tab/vb-3)
```vb
            Dim commentsToDelete As List(Of Comment) = commentPart.Comments.Elements(Of Comment)().ToList()
```
***

So far, the list of comments contains all of the comments. If the author
parameter is not an empty string, the following code limits the list to
only those comments where the `DocumentFormat.OpenXml.Wordprocessing.Comment.Author`
property matches the parameter you supplied.

### [C#](#tab/cs-4)
```csharp
        if (!String.IsNullOrEmpty(author))
        {
            commentsToDelete = commentsToDelete.Where(c => c.Author == author).ToList();
        }
```
### [Visual Basic](#tab/vb-4)
```vb
            If Not String.IsNullOrEmpty(author) Then
                commentsToDelete = commentsToDelete.Where(Function(c) c.Author = author).ToList()
            End If
```
***

Before deleting any comments, the code retrieves a list of comments ID
values, so that it can later delete matching elements from the document
part. The call to the `System.Linq.Enumerable.Select%2A`
method effectively projects the list of comments, retrieving an 
`System.Collections.Generic.IEnumerable%601` of strings that
contain all the comment ID values.

### [C#](#tab/cs-5)
```csharp
        IEnumerable<string?> commentIds = commentsToDelete.Where(r => r.Id is not null && r.Id.HasValue).Select(r => r.Id?.Value);
```
### [Visual Basic](#tab/vb-5)
```vb
            Dim commentIds As IEnumerable(Of String) = commentsToDelete.Where(Function(r) r IsNot Nothing And r.Id.HasValue).Select(Function(r) r.Id.Value)
```
***

--------------------------------------------------------------------------------

## Deleting Comments and Saving the Part

Given the `commentsToDelete` collection, to
the following code loops through all the comments that require deleting
and performs the deletion.

### [C#](#tab/cs-6)
```csharp
        // Delete each comment in commentToDelete from the 
        // Comments collection.
        foreach (Comment c in commentsToDelete)
        {
            if (c is not null)
            {
                c.Remove();
            }
        }
```
### [Visual Basic](#tab/vb-6)
```vb
            ' Delete each comment in commentToDelete from the Comments 
            ' collection.
            For Each c As Comment In commentsToDelete
                If (c IsNot Nothing) Then
                    c.Remove()
                End If
            Next
```
***

--------------------------------------------------------------------------------

## Deleting Comment References in the Document

Although the code has successfully removed all the comments by this
point, that is not enough. The code must also remove references to the
comments from the document part. This action requires three steps
because the comment reference includes the
`DocumentFormat.OpenXml.Wordprocessing.CommentRangeStart`,
`DocumentFormat.OpenXml.Wordprocessing.CommentRangeEnd`,
and `DocumentFormat.OpenXml.Wordprocessing.CommentReference`
elements, and the code must remove all three for each comment.
Before performing any deletions, the code first retrieves a reference
to the root element of the main document part, as shown in the following code.

### [C#](#tab/cs-7)
```csharp
    Document doc = document.MainDocumentPart.Document;
```

### [Visual Basic](#tab/vb-7)
```vb
    Dim doc As Document = document.MainDocumentPart.Document
```
***

Given a reference to the document element, the following code performs
its deletion loop three times, once for each of the different elements
it must delete. In each case, the code looks for all descendants of the
correct type (`CommentRangeStart`, `CommentRangeEnd`, or `CommentReference`)
and limits the list to those whose `DocumentFormat.OpenXml.Wordprocessing.MarkupRangeType.Id`
property value is contained in the list of comment IDs to be deleted.
Given the list of elements to be deleted, the code removes each element in turn.
Finally, the code completes by saving the document.

### [C#](#tab/cs-8)
```csharp
        // Delete CommentRangeStart for each
        // deleted comment in the main document.
        List<CommentRangeStart> commentRangeStartToDelete = doc.Descendants<CommentRangeStart>()
                                                            .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                            .ToList();

        foreach (CommentRangeStart c in commentRangeStartToDelete)
        {
            c.Remove();
        }

        // Delete CommentRangeEnd for each deleted comment in the main document.
        List<CommentRangeEnd> commentRangeEndToDelete = doc.Descendants<CommentRangeEnd>()
                                                        .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                        .ToList();

        foreach (CommentRangeEnd c in commentRangeEndToDelete)
        {
            c.Remove();
        }

        // Delete CommentReference for each deleted comment in the main document.
        List<CommentReference> commentRangeReferenceToDelete = doc.Descendants<CommentReference>()
                                                                .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                                .ToList();

        foreach (CommentReference c in commentRangeReferenceToDelete)
        {
            c.Remove();
        }
```
### [Visual Basic](#tab/vb-8)
```vb
            ' Delete CommentRangeStart for each 
            ' deleted comment in the main document.
            Dim commentRangeStartToDelete As List(Of CommentRangeStart) = doc.Descendants(Of CommentRangeStart).
                                                                            Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                            ToList()

            For Each c As CommentRangeStart In commentRangeStartToDelete
                c.Remove()
            Next

            ' Delete CommentRangeEnd for each deleted comment in main document.
            Dim commentRangeEndToDelete As List(Of CommentRangeEnd) = doc.Descendants(Of CommentRangeEnd).
                                                                        Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                        ToList()

            For Each c As CommentRangeEnd In commentRangeEndToDelete
                c.Remove()
            Next

            ' Delete CommentReference for each deleted comment in the main document.
            Dim commentRangeReferenceToDelete As List(Of CommentReference) = doc.Descendants(Of CommentReference).
                                                                                Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                                ToList()

            For Each c As CommentReference In commentRangeReferenceToDelete
                c.Remove()
            Next
```
***

--------------------------------------------------------------------------------

## Sample Code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
// Delete comments by a specific author. Pass an empty string for the 
// author to delete all comments, by all authors.
static void DeleteComments(string fileName, string author = "")
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        If (args.Length >= 2) Then
            Dim fileName As String = args(0)
            Dim author As String = args(1)

            DeleteComments(fileName, author)
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            DeleteComments(fileName)
        End If
```

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
