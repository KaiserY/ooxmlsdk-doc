# Add a comment to a slide in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to add a comment to the first slide in a presentation
programmatically.

> **Note**
> This sample is for PowerPoint modern comments. For classic comments view
> the [archived sample on GitHub](https://github.com/OfficeDev/open-xml-docs/blob/7002d692ab4abc629d617ef6a0214fc2bf2910c8/docs/how-to-add-a-comment-to-a-slide-in-a-presentation.md).

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

The sample code opens the presentation document in the using statement. Then it instantiates the CommentAuthorsPart, and verifies that there is an existing comment authors part. If there is not, it adds one.

### [C#](#tab/cs-1)
```csharp
        // create missing PowerPointAuthorsPart if it is null
        if (presentationDocument.PresentationPart.authorsPart is null)
        {
            presentationDocument.PresentationPart.AddNewPart<PowerPointAuthorsPart>();
        }
```

### [Visual Basic](#tab/vb-1)
```vb
            ' create missing PowerPointAuthorsPart if it is null
            If presentationDocument.PresentationPart.authorsPart Is Nothing Then
                presentationDocument.PresentationPart.AddNewPart(Of PowerPointAuthorsPart)()
            End If
```
***

The code determines whether there is an existing PowerPoint author part in the presentation part; if not, it adds one, then checks if there is an authors list 
and adds one if it is missing. It also verifies that the author that is passed in is on the list of existing authors; if so, it assigns the existing author ID. If not, it adds a new author to the list of authors and assigns an author ID and the parameter values.

### [C#](#tab/cs-2)
```csharp
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
```

### [Visual Basic](#tab/vb-2)
```vb
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
```
***

Next the code determines if there is a slide id and returns if one does not exist

### [C#](#tab/cs-3)
```csharp
        // Get the Id of the slide to add the comment to
        SlideId? slideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements<SlideId>()?.FirstOrDefault();
        
        // If slideId is null, there are no slides, so return
        if (slideId is null) return;
```

### [Visual Basic](#tab/vb-3)
```vb
            ' Get the Id of the slide to add the comment to
            Dim slideId As SlideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements(Of SlideId)()?.FirstOrDefault()

            ' If slideId is null, there are no slides, so return
            If slideId Is Nothing Then Return
```
***

In the segment below, the code gets the relationship ID. If it exists, it is used to find the slide part
otherwise the first slide in the slide parts enumerable is taken. Then it verifies that there is 
a PowerPoint comments part for the slide and if not adds one.

### [C#](#tab/cs-4)
```csharp
        // Get the relationship id of the slide if it exists
        string? relId = slideId.RelationshipId;

        // Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
        SlidePart slidePart = relId is not null ? (SlidePart)presentationPart.GetPartById(relId) 
            : presentationDocument.PresentationPart.SlideParts.First();

        // If the slide part has comments parts take the first PowerPointCommentsPart
        // otherwise add a new one
        PowerPointCommentPart powerPointCommentPart = slidePart.commentParts.FirstOrDefault() ?? slidePart.AddNewPart<PowerPointCommentPart>();
```

### [Visual Basic](#tab/vb-4)
```vb
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
```
***

Below the code creates a new modern comment then adds a comment list to the PowerPoint comment part
if one does not exist and adds the comment to that comment list.

### [C#](#tab/cs-5)
```csharp
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
```

### [Visual Basic](#tab/vb-5)
```vb
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
```
***

With modern comments the slide needs to have the correct extension list and extension.
The following code determines if the slide already has a SlideExtensionList and
SlideExtension and adds them to the slide if they are not present.

### [C#](#tab/cs-6)
```csharp
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
```

### [Visual Basic](#tab/vb-6)
```vb
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
```
***

## Sample Code

Following is the complete code sample showing how to add a new comment with
a new or existing author to a slide with or without existing comments.

> **Note**
> To get the exact author name and initials, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPointOptions** window opens and the content of the **General** tab is displayed. The author name and initials must match the **User name** and **Initials** in this tab.

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
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
