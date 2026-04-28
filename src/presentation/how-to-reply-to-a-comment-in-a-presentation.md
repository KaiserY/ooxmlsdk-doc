# Reply to a comment in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to reply to existing comments in a presentation
programmatically.

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

The sample code opens the presentation document in the using statement. Then it gets or creates the CommentAuthorsPart, and verifies that there is an existing comment authors part. If there is not, it adds one.

### [C#](#tab/cs-1)
```csharp
// Open the PowerPoint presentation for editing
using (PresentationDocument presentationDocument = PresentationDocument.Open(path, true))
{
    // Check if the presentation part exists
    if (presentationDocument.PresentationPart is null)
    {
        Console.WriteLine("No presentation part found in the presentation");
        return;
    }
    else
    {
        // Prompt the user for the author's name
        Console.WriteLine("Please enter the author's name");

        string? authorName = Console.ReadLine();

        // Ensure the author name is provided
        while (authorName is null)
        {
            Console.WriteLine("Author's name is required. Please enter author's name below");
            authorName = Console.ReadLine();
        }

        // Generate initials from the author's name
        string[] splitName = authorName.Split(" ");
        string authorInitials = string.Concat(splitName[0].Substring(0, 1), splitName[splitName.Length - 1].Substring(0, 1));

        // Get or create the authors part for comment authorship
        PowerPointAuthorsPart authorsPart = presentationDocument.PresentationPart.authorsPart ?? presentationDocument.AddNewPart<PowerPointAuthorsPart>();
        authorsPart.AuthorList ??= new AuthorList();
```

### [Visual Basic](#tab/vb-1)
```vb
        ' Open the PowerPoint presentation for editing
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(path, True)
            ' Check if the presentation part exists
            If presentationDocument.PresentationPart Is Nothing Then
                Console.WriteLine("No presentation part found in the presentation")
                Return
            Else
                ' Prompt the user for the author's name
                Console.WriteLine("Please enter the author's name")

                Dim authorName As String = Console.ReadLine()

                ' Ensure the author name is provided
                While authorName Is Nothing
                    Console.WriteLine("Author's name is required. Please enter author's name below")
                    authorName = Console.ReadLine()
                End While

                ' Generate initials from the author's name
                Dim splitName As String() = authorName.Split(" "c)
                Dim authorInitials As String = splitName(0).Substring(0, 1) & splitName(splitName.Length - 1).Substring(0, 1)

                ' Get or create the authors part for comment authorship
                Dim authorsPart As PowerPointAuthorsPart = If(presentationDocument.PresentationPart.authorsPart, presentationDocument.AddNewPart(Of PowerPointAuthorsPart)())
                If authorsPart.AuthorList Is Nothing Then
                    authorsPart.AuthorList = New AuthorList()
                End If
```
***

Next the code determines if the author that is passed in is on the list of existing authors; if so, it assigns the existing author ID. If not, it adds a new author to the list of authors and assigns an author ID and the parameter values.

### [C#](#tab/cs-2)
```csharp
        // Try to find an existing author by name, otherwise create a new author
        string? authorId = authorsPart.AuthorList.Descendants<Author>().Where(author => author.Name == authorName).FirstOrDefault()?.UserId;

        if (authorId is null)
        {
            authorId = Guid.NewGuid().ToString("B");
            authorsPart.AuthorList.AppendChild(new Author() { Id = authorId, Name = authorName, Initials = authorInitials, UserId = authorId, ProviderId = "Me" });
        }
```

### [Visual Basic](#tab/vb-2)
```vb
                ' Try to find an existing author by name, otherwise create a new author
                Dim authorId As String = Nothing
                Dim existingAuthor = authorsPart.AuthorList.Descendants(Of Author)().Where(Function(author) author.Name = authorName).FirstOrDefault()
                If existingAuthor IsNot Nothing Then
                    authorId = existingAuthor.UserId
                End If

                If authorId Is Nothing Then
                    authorId = Guid.NewGuid().ToString("B")
                    authorsPart.AuthorList.AppendChild(New Author() With {.Id = authorId, .Name = authorName, .Initials = authorInitials, .UserId = authorId, .ProviderId = "Me"})
                End If
```
***

Next the code gets the first slide part and verifies that it exists, then checks if there are any comment parts associated with the slide.

### [C#](#tab/cs-3)
```csharp
        // Get the first slide part in the presentation
        SlidePart? slidePart = presentationDocument.PresentationPart.SlideParts?.FirstOrDefault();

        if (slidePart is null)
        {
            Console.WriteLine("No slide part found in the presentation.");
            return;
        }
        else
        {
            // Check if the slide has any comment parts
            if (slidePart.commentParts is null || slidePart.commentParts.Count() == 0)
            {
                Console.WriteLine("No comments part found for slide 1");
                return;
            }
```

### [Visual Basic](#tab/vb-3)
```vb
                ' Get the first slide part in the presentation
                Dim slidePart As SlidePart = presentationDocument.PresentationPart.SlideParts?.FirstOrDefault()

                If slidePart Is Nothing Then
                    Console.WriteLine("No slide part found in the presentation.")
                    Return
                Else
                    ' Check if the slide has any comment parts
                    If slidePart.commentParts Is Nothing OrElse Not slidePart.commentParts.Any() Then
                        Console.WriteLine("No comments part found for slide 1")
                        Return
                    End If
```
***

The code then retrieves the comment list and then iterates through each comment in the comment list, displays the comment text to the user, and prompts whether they want to reply to each comment.

### [C#](#tab/cs-5)
```csharp
            // Get the comment list
            DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList? commentList =
                (DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList?)(slidePart.commentParts.FirstOrDefault()?.CommentList);

            if (commentList is null)
            {
                Console.WriteLine("No comments found for slide 1");
                return;
            }
            else
            {
                // Iterate through each comment in the comment list
                foreach (DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment comment in commentList)
                {
                    // Display the comment text to the user
                    Console.WriteLine("Comment:");
                    Console.WriteLine(comment.ChildElements.Where(c => c is TextBodyType).FirstOrDefault()?.InnerText);
                    Console.WriteLine("Do you want to reply Y/N");
                    string? leaveReply = Console.ReadLine();
```

### [Visual Basic](#tab/vb-5)
```vb
                    ' Get the comment list
                    Dim commentList As CommentList = CType(slidePart.commentParts.FirstOrDefault()?.CommentList, CommentList)

                    If commentList Is Nothing Then
                        Console.WriteLine("No comments found for slide 1")
                        Return
                    Else
                        ' Iterate through each comment in the comment list
                        For Each comment As Comment In commentList
                            ' Display the comment text to the user
                            Console.WriteLine("Comment:")
                            Console.WriteLine(comment.ChildElements.Where(Function(c) TypeOf c Is TextBodyType).FirstOrDefault()?.InnerText)
                            Console.WriteLine("Do you want to reply Y/N")
                            Dim leaveReply As String = Console.ReadLine()
```
***

When the user chooses to reply to a comment, the code prompts for the reply text, then gets or creates a `CommentReplyList` for the comment and adds the new reply with the appropriate author information and timestamp.

### [C#](#tab/cs-6)
```csharp
                    // If the user wants to reply, prompt for the reply text
                    if (leaveReply is not null && leaveReply.ToUpper() == "Y")
                    {
                        Console.WriteLine("What is your reply?");
                        string? reply = Console.ReadLine();

                        if (reply is not null)
                        {
                            // Get or create the reply list for the comment
                            CommentReplyList? commentReplyList = comment.Descendants<CommentReplyList>()?.FirstOrDefault();

                            if (commentReplyList is null)
                            {
                                commentReplyList = new CommentReplyList();
                                comment.AddChild(commentReplyList);
                            }

                            // Add the user's reply to the comment
                            commentReplyList.AppendChild(new CommentReply(
                                new TextBodyType(
                                    new BodyProperties(),
                                    new Paragraph(
                                        new Run(
                                            new DocumentFormat.OpenXml.Drawing.Text(reply)))))
                            {
                                Id = Guid.NewGuid().ToString("B"),
                                AuthorId = authorId,
                                Created = new DateTimeValue(DateTime.Now)
                            });
                        }
                    }
```

### [Visual Basic](#tab/vb-6)
```vb
                            ' If the user wants to reply, prompt for the reply text
                            If leaveReply IsNot Nothing AndAlso leaveReply.ToUpper() = "Y" Then
                                Console.WriteLine("What is your reply?")
                                Dim reply As String = Console.ReadLine()

                                If reply IsNot Nothing Then
                                    ' Get or create the reply list for the comment
                                    Dim commentReplyList As CommentReplyList = comment.Descendants(Of CommentReplyList)()?.FirstOrDefault()

                                    If commentReplyList Is Nothing Then
                                        commentReplyList = New CommentReplyList()
                                        comment.AddChild(commentReplyList)
                                    End If

                                    ' Add the user's reply to the comment
                                    commentReplyList.AppendChild(New CommentReply(
                                        New TextBodyType(
                                            New BodyProperties(),
                                            New Paragraph(
                                                New Run(
                                                    New DocumentFormat.OpenXml.Drawing.Text(reply))))) With {
                                        .Id = Guid.NewGuid().ToString("B"),
                                        .AuthorId = authorId,
                                        .Created = New DateTimeValue(DateTime.Now)
                                    })
                                End If
                            End If
```
***

## Sample Code

Following is the complete code sample showing how to reply to existing comments
in a presentation slide with modern PowerPoint comments.

### [C#](#tab/cs)
```csharp
// Open the PowerPoint presentation for editing
using (PresentationDocument presentationDocument = PresentationDocument.Open(path, true))
{
    // Check if the presentation part exists
    if (presentationDocument.PresentationPart is null)
    {
        Console.WriteLine("No presentation part found in the presentation");
        return;
    }
    else
    {
        // Prompt the user for the author's name
        Console.WriteLine("Please enter the author's name");

        string? authorName = Console.ReadLine();

        // Ensure the author name is provided
        while (authorName is null)
        {
            Console.WriteLine("Author's name is required. Please enter author's name below");
            authorName = Console.ReadLine();
        }

        // Generate initials from the author's name
        string[] splitName = authorName.Split(" ");
        string authorInitials = string.Concat(splitName[0].Substring(0, 1), splitName[splitName.Length - 1].Substring(0, 1));

        // Get or create the authors part for comment authorship
        PowerPointAuthorsPart authorsPart = presentationDocument.PresentationPart.authorsPart ?? presentationDocument.AddNewPart<PowerPointAuthorsPart>();
        authorsPart.AuthorList ??= new AuthorList();
        // Try to find an existing author by name, otherwise create a new author
        string? authorId = authorsPart.AuthorList.Descendants<Author>().Where(author => author.Name == authorName).FirstOrDefault()?.UserId;

        if (authorId is null)
        {
            authorId = Guid.NewGuid().ToString("B");
            authorsPart.AuthorList.AppendChild(new Author() { Id = authorId, Name = authorName, Initials = authorInitials, UserId = authorId, ProviderId = "Me" });
        }
        // Get the first slide part in the presentation
        SlidePart? slidePart = presentationDocument.PresentationPart.SlideParts?.FirstOrDefault();

        if (slidePart is null)
        {
            Console.WriteLine("No slide part found in the presentation.");
            return;
        }
        else
        {
            // Check if the slide has any comment parts
            if (slidePart.commentParts is null || slidePart.commentParts.Count() == 0)
            {
                Console.WriteLine("No comments part found for slide 1");
                return;
            }
            // Get the comment list
            DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList? commentList =
                (DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList?)(slidePart.commentParts.FirstOrDefault()?.CommentList);

            if (commentList is null)
            {
                Console.WriteLine("No comments found for slide 1");
                return;
            }
            else
            {
                // Iterate through each comment in the comment list
                foreach (DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment comment in commentList)
                {
                    // Display the comment text to the user
                    Console.WriteLine("Comment:");
                    Console.WriteLine(comment.ChildElements.Where(c => c is TextBodyType).FirstOrDefault()?.InnerText);
                    Console.WriteLine("Do you want to reply Y/N");
                    string? leaveReply = Console.ReadLine();
                    // If the user wants to reply, prompt for the reply text
                    if (leaveReply is not null && leaveReply.ToUpper() == "Y")
                    {
                        Console.WriteLine("What is your reply?");
                        string? reply = Console.ReadLine();

                        if (reply is not null)
                        {
                            // Get or create the reply list for the comment
                            CommentReplyList? commentReplyList = comment.Descendants<CommentReplyList>()?.FirstOrDefault();

                            if (commentReplyList is null)
                            {
                                commentReplyList = new CommentReplyList();
                                comment.AddChild(commentReplyList);
                            }

                            // Add the user's reply to the comment
                            commentReplyList.AppendChild(new CommentReply(
                                new TextBodyType(
                                    new BodyProperties(),
                                    new Paragraph(
                                        new Run(
                                            new DocumentFormat.OpenXml.Drawing.Text(reply)))))
                            {
                                Id = Guid.NewGuid().ToString("B"),
                                AuthorId = authorId,
                                Created = new DateTimeValue(DateTime.Now)
                            });
                        }
                    }
                }
            }
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
        ' Open the PowerPoint presentation for editing
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(path, True)
            ' Check if the presentation part exists
            If presentationDocument.PresentationPart Is Nothing Then
                Console.WriteLine("No presentation part found in the presentation")
                Return
            Else
                ' Prompt the user for the author's name
                Console.WriteLine("Please enter the author's name")

                Dim authorName As String = Console.ReadLine()

                ' Ensure the author name is provided
                While authorName Is Nothing
                    Console.WriteLine("Author's name is required. Please enter author's name below")
                    authorName = Console.ReadLine()
                End While

                ' Generate initials from the author's name
                Dim splitName As String() = authorName.Split(" "c)
                Dim authorInitials As String = splitName(0).Substring(0, 1) & splitName(splitName.Length - 1).Substring(0, 1)

                ' Get or create the authors part for comment authorship
                Dim authorsPart As PowerPointAuthorsPart = If(presentationDocument.PresentationPart.authorsPart, presentationDocument.AddNewPart(Of PowerPointAuthorsPart)())
                If authorsPart.AuthorList Is Nothing Then
                    authorsPart.AuthorList = New AuthorList()
                End If
                ' Try to find an existing author by name, otherwise create a new author
                Dim authorId As String = Nothing
                Dim existingAuthor = authorsPart.AuthorList.Descendants(Of Author)().Where(Function(author) author.Name = authorName).FirstOrDefault()
                If existingAuthor IsNot Nothing Then
                    authorId = existingAuthor.UserId
                End If

                If authorId Is Nothing Then
                    authorId = Guid.NewGuid().ToString("B")
                    authorsPart.AuthorList.AppendChild(New Author() With {.Id = authorId, .Name = authorName, .Initials = authorInitials, .UserId = authorId, .ProviderId = "Me"})
                End If
                ' Get the first slide part in the presentation
                Dim slidePart As SlidePart = presentationDocument.PresentationPart.SlideParts?.FirstOrDefault()

                If slidePart Is Nothing Then
                    Console.WriteLine("No slide part found in the presentation.")
                    Return
                Else
                    ' Check if the slide has any comment parts
                    If slidePart.commentParts Is Nothing OrElse Not slidePart.commentParts.Any() Then
                        Console.WriteLine("No comments part found for slide 1")
                        Return
                    End If
                    ' Get the comment list
                    Dim commentList As CommentList = CType(slidePart.commentParts.FirstOrDefault()?.CommentList, CommentList)

                    If commentList Is Nothing Then
                        Console.WriteLine("No comments found for slide 1")
                        Return
                    Else
                        ' Iterate through each comment in the comment list
                        For Each comment As Comment In commentList
                            ' Display the comment text to the user
                            Console.WriteLine("Comment:")
                            Console.WriteLine(comment.ChildElements.Where(Function(c) TypeOf c Is TextBodyType).FirstOrDefault()?.InnerText)
                            Console.WriteLine("Do you want to reply Y/N")
                            Dim leaveReply As String = Console.ReadLine()
                            ' If the user wants to reply, prompt for the reply text
                            If leaveReply IsNot Nothing AndAlso leaveReply.ToUpper() = "Y" Then
                                Console.WriteLine("What is your reply?")
                                Dim reply As String = Console.ReadLine()

                                If reply IsNot Nothing Then
                                    ' Get or create the reply list for the comment
                                    Dim commentReplyList As CommentReplyList = comment.Descendants(Of CommentReplyList)()?.FirstOrDefault()

                                    If commentReplyList Is Nothing Then
                                        commentReplyList = New CommentReplyList()
                                        comment.AddChild(commentReplyList)
                                    End If

                                    ' Add the user's reply to the comment
                                    commentReplyList.AppendChild(New CommentReply(
                                        New TextBodyType(
                                            New BodyProperties(),
                                            New Paragraph(
                                                New Run(
                                                    New DocumentFormat.OpenXml.Drawing.Text(reply))))) With {
                                        .Id = Guid.NewGuid().ToString("B"),
                                        .AuthorId = authorId,
                                        .Created = New DateTimeValue(DateTime.Now)
                                    })
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End Using
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
