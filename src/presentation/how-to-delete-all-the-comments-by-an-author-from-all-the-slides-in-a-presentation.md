# Delete all the comments by an author from all the slides in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to delete all of the comments by a specific author in a
presentation programmatically.

> **Note**
> This sample is for PowerPoint modern comments. For classic comments view
> the [archived sample on GitHub](https://github.com/OfficeDev/open-xml-docs/blob/7002d692ab4abc629d617ef6a0214fc2bf2910c8/docs/how-to-delete-all-the-comments-by-an-author-from-all-the-slides-in-a-presentatio.md).

## Getting a PresentationDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)` method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value `true` for this parameter
as shown in the following `using` statement.
In this code, the *fileName* parameter is a string that represents the
path for the file from which you want to open the document, and the
author is the user name displayed in the General tab of the PowerPoint
Options.

### [C#](#tab/cs-1)
```csharp
    using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, True)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
This ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `doc`.

## Basic Presentation Document Structure 

The basic document structure of a `PresentationML` document consists of a number of
parts, among which is the main part that contains the presentation
definition. The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the overall form of a `PresentationML` package.

> The main part of a `PresentationML` package
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a *slide* list, a *slide master* list, a *notes
> master* list, and a *handout master* list. The slide list refers to
> all of the slides in the presentation; the slide master list refers to
> the entire slide masters used in the presentation; the notes master
> contains information about the formatting of notes pages; and the
> handout master describes how a handout looks.
> 
> A *handout* is a printed set of slides that can be provided to an
> *audience*.
> 
> As well as text and graphics, each slide can contain *comments* and
> *notes*, can have a *layout*, and can be part of one or more *custom
> presentations*. A comment is an annotation intended for the person
> maintaining the presentation slide deck. A note is a reminder or piece
> of text intended for the presenter or the audience.
> 
> Other features that a `PresentationML`
> document can include the following: *animation*, *audio*, *video*, and
> *transitions* between slides.
> 
> A `PresentationML` document is not stored
> as one large body in a single part. Instead, the elements that
> implement certain groupings of functionality are stored in separate
> parts. For example, all authors in a document are stored in one
> authors part while each slide has its own part.
> 
> ISO/IEC 29500: 2016

The following XML code example represents a presentation that contains
two slides denoted by the IDs 267 and 256.

```xml
    <p:presentation xmlns:p="…" … > 
       <p:sldMasterIdLst>
          <p:sldMasterId
             xmlns:rel="https://…/relationships" rel:id="rId1"/>
       </p:sldMasterIdLst>
       <p:notesMasterIdLst>
          <p:notesMasterId
             xmlns:rel="https://…/relationships" rel:id="rId4"/>
       </p:notesMasterIdLst>
       <p:handoutMasterIdLst>
          <p:handoutMasterId
             xmlns:rel="https://…/relationships" rel:id="rId5"/>
       </p:handoutMasterIdLst>
       <p:sldIdLst>
          <p:sldId id="267"
             xmlns:rel="https://…/relationships" rel:id="rId2"/>
          <p:sldId id="256"
             xmlns:rel="https://…/relationships" rel:id="rId3"/>
       </p:sldIdLst>
           <p:sldSz cx="9144000" cy="6858000"/>
       <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

Using the Open XML SDK, you can create document structure and
content using strongly-typed classes that correspond to PresentationML
elements. You can find these classes in the `DocumentFormat.OpenXml.Presentation`
namespace. The following table lists the class names of the classes that
correspond to the `sld`, `sldLayout`, `sldMaster`, and `notesMaster` elements.

| **PresentationML Element** | **Open XML SDK Class** | **Description** |
|---|---|---|
| `<sld/>` | `DocumentFormat.OpenXml.Presentation.Slide` | Presentation Slide. It is the root element of SlidePart. |
| `<sldLayout/>` | `DocumentFormat.OpenXml.Presentation.SlideLayout` | Slide Layout. It is the root element of SlideLayoutPart. |
| `<sldMaster/>` | `DocumentFormat.OpenXml.Presentation.SlideMaster` | Slide Master. It is the root element of SlideMasterPart. |
| `<notesMaster/>` | `DocumentFormat.OpenXml.Presentation.NotesMaster` | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |

## The Structure of the Comment Element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces comments in a presentation package.

> A comment is a text note attached to a slide, with the primary purpose
> of allowing readers of a presentation to provide feedback to the
> presentation author. Each comment contains an unformatted text string
> and information about its author, and is attached to a particular
> location on a slide. Comments can be visible while editing the
> presentation, but do not appear when a slide show is given. The
> displaying application decides when to display comments and determines
> their visual appearance.
> 
> &copy; ISO/IEC 29500: 2016

## The Structure of the Modern Comment Element

The following XML element specifies a single comment. 
It contains the text of the comment (`t`) and attributes referring to its author
(`authorId`), date time created (`created`), and comment id (`id`).

```xml
<p188:cm id="{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}" authorId="{CD37207E-7903-4ED4-8AE8-017538D2DF7E}" created="2024-12-30T20:26:06.503">
  <p188:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
      <a:r>
          <a:t>Needs more cowbell</a:t>
      </a:r>
      </a:p>
  </p188:txBody>
</p188:cm>
```

The following tables list the definitions of the possible child elements and attributes
of the `cm` (comment) element. For the complete definition see [MS-PPTX 2.16.3.3 CT_Comment](https://learn.microsoft.com/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54)

| Attribute | Definition |
|---|---|
| id | Specifies the ID of a comment or a comment reply. |
| authorId | Specifies the author ID of a comment or a comment reply. |
| status | Specifies the status of a comment or a comment reply. |
| created | Specifies the date time when the comment or comment reply is created. |
| startDate | Specifies start date of the comment. |
| dueDate | Specifies due date of the comment. |
| assignedTo | Specifies a list of authors to whom the comment is assigned. |
| complete | Specifies the completion percentage of the comment. |
| title | Specifies the title for a comment. |

| Child Element | Definition |
|------------|---------------|
| pc:sldMkLst | Specifies a content moniker that identifies the slide to which the comment is anchored. |
| ac:deMkLst | Specifies a content moniker that identifies the drawing element to which the comment is anchored. |
| ac:txMkLst | Specifies a content moniker that identifies the text character range to which the comment is anchored. |
| unknownAnchor | Specifies an unknown anchor to which the comment is anchored. |
| pos | Specifies the position of the comment, relative to the top-left corner of the first object to which the comment is anchored. |
| replyLst | Specifies the list of replies to the comment. |
| txBody | Specifies the text of a comment or a comment reply. |
| extLst | Specifies a list of extensions for a comment or a comment reply. |

The following XML schema example defines the members of the `cm` element in addition to the required and
optional attributes.

```xml
 <xsd:complexType name="CT_Comment">
   <xsd:sequence>
     <xsd:group ref="EG_CommentAnchor" minOccurs="1" maxOccurs="1"/>
     <xsd:element name="pos" type="a:CT_Point2D" minOccurs="0" maxOccurs="1"/>
     <xsd:element name="replyLst" type="CT_CommentReplyList" minOccurs="0" maxOccurs="1"/>
     <xsd:group ref="EG_CommentProperties" minOccurs="1" maxOccurs="1"/>
   </xsd:sequence>
   <xsd:attributeGroup ref="AG_CommentProperties"/>
   <xsd:attribute name="startDate" type="xsd:dateTime" use="optional"/>
   <xsd:attribute name="dueDate" type="xsd:dateTime" use="optional"/>
   <xsd:attribute name="assignedTo" type="ST_AuthorIdList" use="optional" default=""/>
   <xsd:attribute name="complete" type="s:ST_PositiveFixedPercentage" default="0%" use="optional"/>
   <xsd:attribute name="title" type="xsd:string" use="optional" default=""/>
 </xsd:complexType>
```

## How the Sample Code Works

After opening the presentation document for read/write access and
instantiating the `PresentationDocument`
class, the code gets the specified comment author from the list of
comment authors.

### [C#](#tab/cs-2)
```csharp
        // Get the modern comments.
        IEnumerable<Author>? commentAuthors = doc.PresentationPart?.authorsPart?.AuthorList.Elements<Author>()
            .Where(x => x.Name is not null && x.Name.HasValue && x.Name.Value!.Equals(author));
```

### [Visual Basic](#tab/vb-2)
```vb
            ' Get the modern comments.
            Dim commentAuthors As IEnumerable(Of Author) = doc.PresentationPart?.authorsPart?.AuthorList.Elements(Of Author)().Where(Function(x) x.Name IsNot Nothing AndAlso x.Name.HasValue AndAlso x.Name.Value.Equals(author))
```
***

By iterating through the matching authors and all the slides in the
presentation the code gets all the slide parts, and the comments part of
each slide part. It then gets the list of comments by the specified
author and deletes each one. It also verifies that the comment part has
no existing comment, in which case it deletes that part. It also deletes
the comment author from the comment authors part.

### [C#](#tab/cs-3)
```csharp
        // Iterate through all the matching authors.
        foreach (Author commentAuthor in commentAuthors)
        {
            string? authorId = commentAuthor.Id;
            IEnumerable<SlidePart>? slideParts = doc.PresentationPart?.SlideParts;

            // If there's no author ID or slide parts or slide parts, return.
            if (authorId is null || slideParts is null)
            {
                return;
            }

            // Iterate through all the slides and get the slide parts.
            foreach (SlidePart slide in slideParts)
            {
                IEnumerable<PowerPointCommentPart>? slideCommentsParts = slide.commentParts;

                // Get the list of comments.
                if (slideCommentsParts is not null)
                {
                    IEnumerable<Tuple<PowerPointCommentPart, Comment>> commentsTup = slideCommentsParts
                        .SelectMany(scp => scp.CommentList.Elements<Comment>()
                        .Where(comment => comment.AuthorId is not null && comment.AuthorId == authorId)
                        .Select(c => new Tuple<PowerPointCommentPart, Comment>(scp, c)));

                    foreach (Tuple<PowerPointCommentPart, Comment> comment in commentsTup)
                    {
                        // Delete all the comments by the specified author.
                        comment.Item1.CommentList.RemoveChild(comment.Item2);

                        // If the commentPart has no existing comment.
                        if (comment.Item1.CommentList.ChildElements.Count == 0)
                        {
                            // Delete this part.
                            slide.DeletePart(comment.Item1);
                        }
                    }

                }
            }

            // Delete the comment author from the authors part.
            doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor);
        }
```

### [Visual Basic](#tab/vb-3)
```vb
            ' Iterate through all the matching authors.
            For Each commentAuthor As Author In commentAuthors
                Dim authorId As String = commentAuthor.Id
                Dim slideParts As IEnumerable(Of SlidePart) = doc.PresentationPart?.SlideParts

                ' If there's no author ID or slide parts, return.
                If authorId Is Nothing OrElse slideParts Is Nothing Then
                    Return
                End If

                ' Iterate through all the slides and get the slide parts.
                For Each slide As SlidePart In slideParts
                    Dim slideCommentsParts As IEnumerable(Of PowerPointCommentPart) = slide.commentParts

                    ' Get the list of comments.
                    If slideCommentsParts IsNot Nothing Then
                        Dim commentsTup = slideCommentsParts.SelectMany(Function(scp) scp.CommentList.Elements(Of Comment)().Where(Function(comment) comment.AuthorId IsNot Nothing AndAlso comment.AuthorId = authorId).Select(Function(c) New Tuple(Of PowerPointCommentPart, Comment)(scp, c)))

                        For Each comment As Tuple(Of PowerPointCommentPart, Comment) In commentsTup
                            ' Delete all the comments by the specified author.
                            comment.Item1.CommentList.RemoveChild(comment.Item2)

                            ' If the commentPart has no existing comment.
                            If comment.Item1.CommentList.ChildElements.Count = 0 Then
                                ' Delete this part.
                                slide.DeletePart(comment.Item1)
                            End If
                        Next
                    End If
                Next

                ' Delete the comment author from the authors part.
                doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor)
            Next
```
***

## Sample Code

The following method takes as parameters the source presentation file
name and path and the name of the comment author whose comments are to
be deleted. It finds all the comments by the specified author in the
presentation and deletes them. It then deletes the comment author from
the list of comment authors.

> **Note**
> To get the exact author's name, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPoint Options** window opens and the content of the **General** tab is displayed. The author's name must match the **User name** in this tab.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// Remove all the comments in the slides by a certain x.
static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
{
    using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
    {
        // Get the modern comments.
        IEnumerable<Author>? commentAuthors = doc.PresentationPart?.authorsPart?.AuthorList.Elements<Author>()
            .Where(x => x.Name is not null && x.Name.HasValue && x.Name.Value!.Equals(author));
        if (commentAuthors is null)
        {
            return;
        }
        // Iterate through all the matching authors.
        foreach (Author commentAuthor in commentAuthors)
        {
            string? authorId = commentAuthor.Id;
            IEnumerable<SlidePart>? slideParts = doc.PresentationPart?.SlideParts;

            // If there's no author ID or slide parts or slide parts, return.
            if (authorId is null || slideParts is null)
            {
                return;
            }

            // Iterate through all the slides and get the slide parts.
            foreach (SlidePart slide in slideParts)
            {
                IEnumerable<PowerPointCommentPart>? slideCommentsParts = slide.commentParts;

                // Get the list of comments.
                if (slideCommentsParts is not null)
                {
                    IEnumerable<Tuple<PowerPointCommentPart, Comment>> commentsTup = slideCommentsParts
                        .SelectMany(scp => scp.CommentList.Elements<Comment>()
                        .Where(comment => comment.AuthorId is not null && comment.AuthorId == authorId)
                        .Select(c => new Tuple<PowerPointCommentPart, Comment>(scp, c)));

                    foreach (Tuple<PowerPointCommentPart, Comment> comment in commentsTup)
                    {
                        // Delete all the comments by the specified author.
                        comment.Item1.CommentList.RemoveChild(comment.Item2);

                        // If the commentPart has no existing comment.
                        if (comment.Item1.CommentList.ChildElements.Count == 0)
                        {
                            // Delete this part.
                            slide.DeletePart(comment.Item1);
                        }
                    }

                }
            }

            // Delete the comment author from the authors part.
            doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor);
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' Remove all the comments in the slides by a certain author.
    Sub DeleteCommentsByAuthorInPresentation(fileName As String, author As String)
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, True)
            ' Get the modern comments.
            Dim commentAuthors As IEnumerable(Of Author) = doc.PresentationPart?.authorsPart?.AuthorList.Elements(Of Author)().Where(Function(x) x.Name IsNot Nothing AndAlso x.Name.HasValue AndAlso x.Name.Value.Equals(author))
            If commentAuthors Is Nothing Then
                Return
            End If
            ' Iterate through all the matching authors.
            For Each commentAuthor As Author In commentAuthors
                Dim authorId As String = commentAuthor.Id
                Dim slideParts As IEnumerable(Of SlidePart) = doc.PresentationPart?.SlideParts

                ' If there's no author ID or slide parts, return.
                If authorId Is Nothing OrElse slideParts Is Nothing Then
                    Return
                End If

                ' Iterate through all the slides and get the slide parts.
                For Each slide As SlidePart In slideParts
                    Dim slideCommentsParts As IEnumerable(Of PowerPointCommentPart) = slide.commentParts

                    ' Get the list of comments.
                    If slideCommentsParts IsNot Nothing Then
                        Dim commentsTup = slideCommentsParts.SelectMany(Function(scp) scp.CommentList.Elements(Of Comment)().Where(Function(comment) comment.AuthorId IsNot Nothing AndAlso comment.AuthorId = authorId).Select(Function(c) New Tuple(Of PowerPointCommentPart, Comment)(scp, c)))

                        For Each comment As Tuple(Of PowerPointCommentPart, Comment) In commentsTup
                            ' Delete all the comments by the specified author.
                            comment.Item1.CommentList.RemoveChild(comment.Item2)

                            ' If the commentPart has no existing comment.
                            If comment.Item1.CommentList.ChildElements.Count = 0 Then
                                ' Delete this part.
                                slide.DeletePart(comment.Item1)
                            End If
                        Next
                    End If
                Next

                ' Delete the comment author from the authors part.
                doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor)
            Next
        End Using
    End Sub
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
