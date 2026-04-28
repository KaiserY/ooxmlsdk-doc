# Working with comments

This topic discusses the Open XML SDK for Office `DocumentFormat.OpenXml.Presentation.Comment` class and how it relates to the
Open XML File Format PresentationML schema. For more information about the overall structure of the parts and elements that make up a PresentationML document, see [Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

## Comments in PresentationML

The [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification describes the Comments section of the Open XML PresentationML framework as follows:

A comment is a text note attached to a slide, with the primary purpose
of allowing readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author, and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

The [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification describes the Open XML PresentationML `<cm/>` element used to represent comments in a PresentationML document as follows:

This element specifies a single comment attached to a slide. It contains
the text of the comment, its position on the slide, and attributes
referring to its author and date.

Example:

```xml
<p:cm authorId="0" dt="2006-08-28T17:26:44.129" idx="1">  
   <p:pos x="10" y="10"/>  
   <p:text\>Add diagram to clarify.</p:text>  
</p:cm>
```

The following table lists the child elements of the `<cm/>` element used when working with comments and the Open XML SDK classes that correspond to them.

| **PresentationML Element** |                                                               **Open XML SDK Class**                                                 |
|----------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| `<extLst/>`        | `DocumentFormat.OpenXml.Presentation.ExtensionListWithModification` |
| `<pos/>`           | `DocumentFormat.OpenXml.Presentation.Position`                  |
| `<text/>`          | `DocumentFormat.OpenXml.Presentation.Text`                          |

The following table from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the attributes of the `<cm/>` element.

| **Attributes** |  **Description**   |
|----------------|--------------------|
|    authorId    | This attribute specifies the author of the comment.It refers to the ID of an author in the comment author list for the document.<br/>The possible values for this attribute are defined by the W3C XML Schema `unsignedInt` datatype.  |
|       dt       | This attribute specifies the date and time this comment was last modified.<br/>The possible values for this attribute are defined by the W3C XML Schema `datetime` datatype.           |
|      idx       | This attribute specifies an identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1.<br/>Note: Because the index is unique only for the comment author, a document can contain multiple comments with the same index created by different authors.<br/>The possible values for this attribute are defined by the ST_Index simple type (§19.7.3). |

## Open XML SDK Comment Class

The OXML SDK `Comment` class represents the `<cm/>` element defined in the Open XML File Format schema for PresentationML documents. Use the `Comment`
class to manipulate individual `<cm/>` elements in a PresentationML document.

Classes that represent child elements of the `<cm/>` element and that are
therefore commonly associated with the `Comment` class are shown in the following list.

### ExtensionListWithModification Class

The `ExtensionListWithModification` class corresponds to the `<extLst/>`element. The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification introduces the `<extLst/>` element:

This element specifies the extension list with modification ability within which all future extensions of element type `<ext/>` are defined. The extension list along with corresponding future extensions is used to extend the storage capabilities of the PresentationML framework. This allows for various new kinds of data to be stored natively within the framework.

> **Note**
> Using this `extLst` element allows the generating application to store whether this extension property has been modified. end note

### Position Class

The `Position` class corresponds to the
`<pos/>`element. The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification introduces the `<pos/>` element:

This element specifies the positioning information for the placement of
a comment on a slide surface. In LTR versions of the generating
application, this position information should refer to the upper left
point of the comment shape. In RTL versions of the generating
application, this position information should refer to the upper right
point of the comment shape.

[Note: The anchoring point on the slide surface is unaffected by a
right-to-left or left-to-right layout change. That is the anchoring
point remains the same for all language versions. end note]

[Note: Because there is no specified size or formatting for comments,
this UI widget used to display a comment can be any size and thus the
lower right point of the comment shape is determined by how the viewing
application chooses to display comments. end note]

[Example: \<p:pos x="1426" y="660"/\> end example]

### Text class

The `Text` class corresponds to the
`<text/>` element. The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification introduces the `<text/>` element:

This element specifies the content of a comment. This is the text with
which the author has annotated the slide.

The possible values for this element are defined by the W3C XML Schema
`string` datatype.

## Working with the Comment Class

A comment is a text note attached to a slide, with the primary purpose
of enabling readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

As shown in the Open XML SDK code sample that follows, every instance of
the `Comment` class is associated with an
instance of the `DocumentFormat.OpenXml.Packaging.SlideCommentsPart` class, which represents a
slide comments part, one of the parts of a PresentationML presentation
file package, and a part that is required for each slide in a
presentation file that contains comments. Each `Comment` class instance is also associated with an
instance of the `DocumentFormat.OpenXml.Presentation.CommentAuthor` class, which is in turn
associated with a similarly named presentation part, represented by the
`DocumentFormat.OpenXml.Packaging.CommentAuthorsPart` class. Comment authors
for a presentation are specified in a comment author list, represented
by the `DocumentFormat.OpenXml.Presentation.CommentAuthorList` class, while comments for
each slide are listed in a comments list for that slide, represented by
the `DocumentFormat.OpenXml.Presentation.CommentList` class.

The `Comment` class, which represents the `<cm/>` element, is therefore also associated with other classes that represent the child elements of the `<cm/>` element. Among these classes, as shown in the following code sample, are the `Position` class, which specifies the position of the comment relative to the slide, and the `Text` class, which specifies the text content of the comment.

## Open XML SDK Code Example

The following code segment from the article [How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md) adds a new comments part to an existing slide in a presentation (if the slide does not already contain comments) and creates an instance of an Open XML SDK `Comment` class in the slide comments part. It also adds a comment list to the comments part by creating an instance of the `CommentList` class, if one does not already exist; assigns an ID to the comment; and then adds a comment to the comment list by creating an instance of the `Comment` class, assigning the required attribute values. In addition, it creates instances of the `Position` and `Text` classes associated with the new `Comment` class instance. For the complete code sample, see the aforementioned article.

### [C#](#tab/cs)
```csharp
static void AddCommentToPresentation(string file, string initials, string name, string text)
{
    using (PresentationDocument presentationDocument = PresentationDocument.Open(file, true))
    {
        PresentationPart presentationPart = presentationDocument?.PresentationPart ?? throw new MissingFieldException("PresentationPart");
        // create missing PowerPointAuthorsPart if it is null
        if (presentationDocument.PresentationPart.authorsPart is null)
        {
            presentationDocument.PresentationPart.AddNewPart<PowerPointAuthorsPart>();
        }
        // Add missing AuthorList if it is null
        if (presentationDocument.PresentationPart.authorsPart!.AuthorList is null)
        {
            presentationDocument.PresentationPart.authorsPart.AuthorList = new AuthorList();
        }

        // Get the author or create a new one
        Author? author = presentationDocument.PresentationPart.authorsPart.AuthorList
            .ChildElements.OfType<Author>().Where(a => a.Name?.Value == name).FirstOrDefault();

        if (author is null)
        {
            string authorId = string.Concat("{", Guid.NewGuid(), "}");
            string userId = string.Concat(name.Split(" ").FirstOrDefault() ?? "user", "@example.com::", Guid.NewGuid());
            author = new Author() { Id = authorId, Name = name, Initials = initials, UserId = userId, ProviderId = string.Empty };

            presentationDocument.PresentationPart.authorsPart.AuthorList.AppendChild(author);
        }
        // Get the Id of the slide to add the comment to
        SlideId? slideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements<SlideId>()?.FirstOrDefault();
        
        // If slideId is null, there are no slides, so return
        if (slideId is null) return;
        Random ran = new Random();
        UInt32Value cid = Convert.ToUInt32(ran.Next(100000000, 999999999));
        // Get the relationship id of the slide if it exists
        string? relId = slideId.RelationshipId;

        // Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
        SlidePart slidePart = relId is not null ? (SlidePart)presentationPart.GetPartById(relId) 
            : presentationDocument.PresentationPart.SlideParts.First();

        // If the slide part has comments parts take the first PowerPointCommentsPart
        // otherwise add a new one
        PowerPointCommentPart powerPointCommentPart = slidePart.commentParts.FirstOrDefault() ?? slidePart.AddNewPart<PowerPointCommentPart>();
        // Create the comment using the new modern comment class DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment
        DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment comment = new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                new SlideMonikerList(
                    new DocumentMoniker(),
                    new SlideMoniker()
                    {
                        CId = cid,
                        SldId = slideId.Id,
                    }),
                new TextBodyType(
                    new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text(text)))))
        {
            Id = string.Concat("{", Guid.NewGuid(), "}"),
            AuthorId = author.Id,
            Created = DateTime.Now,
        };

        // If the comment list does not exist, add one.
        powerPointCommentPart.CommentList ??= new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList();
        // Add the comment to the comment list
        powerPointCommentPart.CommentList.AppendChild(comment);
        // Get the presentation extension list if it exists
        SlideExtensionList? presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().FirstOrDefault();
        // Create a boolean that determines if this is the slide's first comment
        bool isFirstComment = false;

        // If the presentation extension list is null, add one and set this as the first comment for the slide
        if (presentationExtensionList is null)
        {
            isFirstComment = true;
            slidePart.Slide.AppendChild(new SlideExtensionList());
            presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().First();
        }

        // Get the slide extension if it exists
        SlideExtension? presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().FirstOrDefault();

        // If the slide extension is null, add it and set this as a new comment
        if (presentationExtension is null)
        {
            isFirstComment = true;
            presentationExtensionList.AddChild(new SlideExtension() { Uri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}" });
            presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().First();
        }

        // If this is the first comment for the slide add the comment relationship
        if (isFirstComment)
        {
            presentationExtension.AddChild(new CommentRelationship()
            { Id = slidePart.GetIdOfPart(powerPointCommentPart) });
        }
    }
}
```
### [Visual Basic](#tab/vb)
```vb
    Sub AddCommentToPresentation(file As String, initials As String, name As String, text As String)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(file, True)
            Dim presentationPart As PresentationPart = presentationDocument?.PresentationPart

            If (presentationPart Is Nothing) Then
                Throw New MissingFieldException("PresentationPart")
            End If
            ' create missing PowerPointAuthorsPart if it is null
            If presentationDocument.PresentationPart.authorsPart Is Nothing Then
                presentationDocument.PresentationPart.AddNewPart(Of PowerPointAuthorsPart)()
            End If
            ' Add missing AuthorList if it is null
            If presentationDocument.PresentationPart.authorsPart Is Nothing Or presentationDocument.PresentationPart.authorsPart.AuthorList Is Nothing Then
                presentationDocument.PresentationPart.authorsPart.AuthorList = New AuthorList()
            End If

            ' Get the author or create a new one
            Dim author As Author = presentationDocument.PresentationPart.authorsPart.AuthorList _
                .ChildElements.OfType(Of Author)().Where(Function(a) a.Name?.Value = name).FirstOrDefault()

            If author Is Nothing Then
                Dim authorId As String = String.Concat("{", Guid.NewGuid(), "}")
                Dim userId As String = String.Concat(If(name.Split(" "c).FirstOrDefault(), "user"), "@example.com::", Guid.NewGuid())
                author = New Author() With {.Id = authorId, .Name = name, .Initials = initials, .UserId = userId, .ProviderId = String.Empty}

                presentationDocument.PresentationPart.authorsPart.AuthorList.AppendChild(author)
            End If
            ' Get the Id of the slide to add the comment to
            Dim slideId As SlideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements(Of SlideId)()?.FirstOrDefault()

            ' If slideId is null, there are no slides, so return
            If slideId Is Nothing Then Return
            Dim ran As New Random()
            Dim cid As UInt32Value = Convert.ToUInt32(ran.Next(100000000, 999999999))
            ' Get the relationship id of the slide if it exists
            Dim relId As String = slideId.RelationshipId

            ' Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
            Dim slidePart As SlidePart = If(relId IsNot Nothing, CType(presentationPart.GetPartById(relId), SlidePart), presentationDocument.PresentationPart.SlideParts.First())

            ' If the slide part has comments parts take the first PowerPointCommentsPart
            ' otherwise add a new one
            Dim powerPointCommentPart As PowerPointCommentPart = slidePart.commentParts.FirstOrDefault()

            If (powerPointCommentPart Is Nothing) Then
                powerPointCommentPart = slidePart.AddNewPart(Of PowerPointCommentPart)()
            End If
            ' Create the comment using the new modern comment class DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment
            Dim comment As New DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                New SlideMonikerList(
                    New DocumentMoniker(),
                    New SlideMoniker() With {
                        .CId = cid,
                        .SldId = slideId.Id
                    }),
                New TextBodyType(
                    New BodyProperties(),
                    New ListStyle(),
                    New Paragraph(New Run(New DocumentFormat.OpenXml.Drawing.Text(text))))) With {
                .Id = String.Concat("{", Guid.NewGuid(), "}"),
                .AuthorId = author.Id,
                .Created = DateTime.Now
            }

            ' If the comment list does not exist, add one.
            If (powerPointCommentPart.CommentList Is Nothing) Then
                powerPointCommentPart.CommentList = New DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList()
            End If
            ' Add the comment to the comment list
            powerPointCommentPart.CommentList.AppendChild(comment)
            ' Get the presentation extension list if it exists
            Dim presentationExtensionList As SlideExtensionList = slidePart.Slide.ChildElements.OfType(Of SlideExtensionList)().FirstOrDefault()
            ' Create a boolean that determines if this is the slide's first comment
            Dim isFirstComment As Boolean = False

            ' If the presentation extension list is null, add one and set this as the first comment for the slide
            If presentationExtensionList Is Nothing Then
                isFirstComment = True
                slidePart.Slide.AppendChild(New SlideExtensionList())
                presentationExtensionList = slidePart.Slide.ChildElements.OfType(Of SlideExtensionList)().First()
            End If

            ' Get the slide extension if it exists
            Dim presentationExtension As SlideExtension = presentationExtensionList.ChildElements.OfType(Of SlideExtension)().FirstOrDefault()

            ' If the slide extension is null, add it and set this as a new comment
            If presentationExtension Is Nothing Then
                isFirstComment = True
                presentationExtensionList.AddChild(New SlideExtension() With {.Uri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"})
                presentationExtension = presentationExtensionList.ChildElements.OfType(Of SlideExtension)().First()
            End If

            ' If this is the first comment for the slide add the comment relationship
            If isFirstComment Then
                presentationExtension.AddChild(New CommentRelationship() With {.Id = slidePart.GetIdOfPart(powerPointCommentPart)})
            End If
        End Using
    End Sub
```

## Generated PresentationML

When the Open XML SDK code in [How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md) is run, including
the segment shown in this article, the following XML is written to a new CommentAuthors.xml part in the existing PresentationML document referenced in the code, assuming that the document contained no comments
or comment authors before the code was run.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cmAuthor id="1"
                  name="userName"
                  initials="userInitials"
                  lastIdx="1"
                  clrIdx="0" />
    </p:cmAuthorLst>
```

In addition, the following XML is written to a new Comments.xml part in
the existing PresentationML document referenced in the code in the
article.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cm authorId="1"
            dt="2010-09-07T16:01:18.5351166-07:00"
            idx="1">
        <p:pos x="100"
               y="200" />
        <p:text>commentText</p:text>
      </p:cm>
    </p:cmLst>
```

## See also

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)
[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)
[How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md)
[How to: Delete all the comments by an author from all the slides in a presentation](how-to-delete-all-the-comments-by-an-author-from-all-the-slides-in-a-presentation.md)
