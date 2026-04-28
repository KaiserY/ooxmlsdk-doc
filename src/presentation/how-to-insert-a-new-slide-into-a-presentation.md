# Insert a new slide into a presentation

This topic shows how to use the classes in the Open XML SDK to
insert a new slide into a presentation programmatically.

## Getting a PresentationDocument Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)` method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value `true` for this parameter
as shown in the following `using` statement.
In this code segment, the `presentationFile` parameter is a string that
represents the full path for the file from which you want to open the
document.

### [C#](#tab/cs-1)
```csharp
            // Open the source document as read/write. 
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
```

### [Visual Basic](#tab/vb-1)
```vb
            ' Open the source document as read/write. 
            Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
This ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `presentationDocument`.

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

## How the Sample Code Works 

The sample code consists of two overloads of the `InsertNewSlide` method. The first overloaded
method takes three parameters: the full path to the presentation file to
which to add a slide, an integer that represents the zero-based slide
index position in the presentation where to add the slide, and the
string that represents the title of the new slide. It opens the
presentation file as read/write, gets a `PresentationDocument` object, and then passes that
object to the second overloaded `InsertNewSlide` method, which performs the
insertion.

### [C#](#tab/cs-10)
```csharp
        // Insert a slide into the specified presentation.
        public static void InsertNewSlide(string presentationFile, int position, string slideTitle)
        {
            // Open the source document as read/write. 
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                // Pass the source document and the position and title of the slide to be inserted to the next method.
                InsertNewSlide(presentationDocument, position, slideTitle);
            }
        }
```

### [Visual Basic](#tab/vb-10)
```vb
        ' Insert a slide into the specified presentation.
        Public Shared Sub InsertNewSlide(presentationFile As String, position As Integer, slideTitle As String)
            ' Open the source document as read/write. 
            Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
                ' Pass the source document and the position and title of the slide to be inserted to the next method.
                InsertNewSlide(presentationDocument, position, slideTitle)
            End Using
        End Sub
```
***

The second overloaded `InsertNewSlide` method
creates a new `Slide` object, sets its
properties, and then inserts it into the slide order in the
presentation. The first section of the method creates the slide and sets
its properties.

### [C#](#tab/cs-11)
```csharp
        // Insert the specified slide into the presentation at the specified position.
        public static SlidePart InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart is null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            CommonSlideData commonSlideData = slide.CommonSlideData ?? slide.AppendChild(new CommonSlideData());
            ShapeTree shapeTree = commonSlideData.ShapeTree ?? commonSlideData.AppendChild(new ShapeTree());
            NonVisualGroupShapeProperties nonVisualProperties = shapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            shapeTree.AppendChild(new GroupShapeProperties());
```

### [Visual Basic](#tab/vb-11)
```vb
        ' Insert the specified slide into the presentation at the specified position.
        Public Shared Function InsertNewSlide(presentationDocument As PresentationDocument, position As Integer, slideTitle As String) As SlidePart
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Verify that the presentation is not empty.
            If presentationPart Is Nothing Then
                Throw New InvalidOperationException("The presentation document is empty.")
            End If

            ' Declare and instantiate a new slide.
            Dim slide As New Slide(New CommonSlideData(New ShapeTree()))
            Dim drawingObjectId As UInteger = 1

            ' Construct the slide content.            
            ' Specify the non-visual properties of the new slide.
            Dim commonSlideData As CommonSlideData = If(slide.CommonSlideData, slide.AppendChild(New CommonSlideData()))
            Dim shapeTree As ShapeTree = If(commonSlideData.ShapeTree, commonSlideData.AppendChild(New ShapeTree()))
            Dim nonVisualProperties As NonVisualGroupShapeProperties = shapeTree.AppendChild(New NonVisualGroupShapeProperties())
            nonVisualProperties.NonVisualDrawingProperties = New NonVisualDrawingProperties() With {.Id = 1, .Name = ""}
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = New NonVisualGroupShapeDrawingProperties()
            nonVisualProperties.ApplicationNonVisualDrawingProperties = New ApplicationNonVisualDrawingProperties()

            ' Specify the group shape properties of the new slide.
            shapeTree.AppendChild(New GroupShapeProperties())
```
***

The next section of the second overloaded `InsertNewSlide` method adds a title shape to the
slide and sets its properties, including its text.

### [C#](#tab/cs-12)
```csharp
            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = shapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));
```

### [Visual Basic](#tab/vb-12)
```vb
            ' Declare and instantiate the title shape of the new slide.
            Dim titleShape As Shape = shapeTree.AppendChild(New Shape())

            drawingObjectId += 1

            ' Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = New NonVisualShapeProperties(
                New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Title"},
                New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}),
                New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title}))
            titleShape.ShapeProperties = New ShapeProperties()

            ' Specify the text of the title shape.
            titleShape.TextBody = New TextBody(New Drawing.BodyProperties(),
                    New Drawing.ListStyle(),
                    New Drawing.Paragraph(New Drawing.Run(New Drawing.Text() With {.Text = slideTitle})))
```
***

The next section of the second overloaded `InsertNewSlide` method adds a body shape to the
slide and sets its properties, including its text.

### [C#](#tab/cs-13)
```csharp
            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = shapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph());
```

### [Visual Basic](#tab/vb-13)
```vb
            ' Declare and instantiate the body shape of the new slide.
            Dim bodyShape As Shape = shapeTree.AppendChild(New Shape())
            drawingObjectId += 1

            ' Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = New NonVisualShapeProperties(New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Content Placeholder"},
                    New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}),
                    New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Index = 1}))
            bodyShape.ShapeProperties = New ShapeProperties()

            ' Specify the text of the body shape.
            bodyShape.TextBody = New TextBody(New Drawing.BodyProperties(),
                    New Drawing.ListStyle(),
                    New Drawing.Paragraph())
```
***

The final section of the second overloaded `InsertNewSlide` method creates a new slide part,
finds the specified index position where to insert the slide, and then
inserts it and assigns the new slide to the new slide part.

### [C#](#tab/cs-14)
```csharp
            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            
            // Assign the new slide to the new slide part
            slidePart.Slide = slide;

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId? prevSlideId = null;

            OpenXmlElementList slideIds = slideIdList?.ChildElements ?? default;

            foreach (SlideId slideId in slideIds)
            {
                if (slideId.Id is not null && slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId is not null && prevSlideId.RelationshipId is not null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
            }
            else
            {
                string? firstRelId = ((SlideId)slideIds[0]).RelationshipId;
                // If the first slide does not contain a relationship ID, throw an exception.
                if (firstRelId is null)
                {
                    throw new ArgumentNullException(nameof(firstRelId));
                }

                lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
            }

            // Use the same slide layout as that of the previous slide.
            if (lastSlidePart.SlideLayoutPart is not null)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
```

### [Visual Basic](#tab/vb-14)
```vb
            ' Create the slide part for the new slide.
            Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)()

            ' Assign the new slide to the new slide part
            slidePart.Slide = slide

            ' Modify the slide ID list in the presentation part.
            ' The slide ID list should not be null.
            Dim slideIdList As SlideIdList = presentationPart.Presentation.SlideIdList

            ' Find the highest slide ID in the current list.
            Dim maxSlideId As UInteger = 1
            Dim prevSlideId As SlideId = Nothing

            Dim slideIds As OpenXmlElementList = slideIdList?.ChildElements

            For Each slideId As SlideId In slideIds
                If slideId.Id IsNot Nothing AndAlso slideId.Id.Value > maxSlideId Then
                    maxSlideId = slideId.Id
                End If

                position -= 1
                If position = 0 Then
                    prevSlideId = slideId
                End If
            Next

            maxSlideId += 1

            ' Get the ID of the previous slide.
            Dim lastSlidePart As SlidePart

            If prevSlideId IsNot Nothing AndAlso prevSlideId.RelationshipId IsNot Nothing Then
                lastSlidePart = CType(presentationPart.GetPartById(prevSlideId.RelationshipId), SlidePart)
            Else
                Dim firstRelId As String = CType(slideIds(0), SlideId).RelationshipId
                ' If the first slide does not contain a relationship ID, throw an exception.
                If firstRelId Is Nothing Then
                    Throw New ArgumentNullException(NameOf(firstRelId))
                End If

                lastSlidePart = CType(presentationPart.GetPartById(firstRelId), SlidePart)
            End If

            ' Use the same slide layout as that of the previous slide.
            If lastSlidePart.SlideLayoutPart IsNot Nothing Then
                slidePart.AddPart(lastSlidePart.SlideLayoutPart)
            End If

            ' Insert the new slide into the slide list after the previous slide.
            Dim newSlideId As SlideId = slideIdList.InsertAfter(New SlideId(), prevSlideId)
            newSlideId.Id = maxSlideId
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart)
```
***

## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
        // Insert a slide into the specified presentation.
        public static void InsertNewSlide(string presentationFile, int position, string slideTitle)
        {
            // Open the source document as read/write. 
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                // Pass the source document and the position and title of the slide to be inserted to the next method.
                InsertNewSlide(presentationDocument, position, slideTitle);
            }
        }
        // Insert the specified slide into the presentation at the specified position.
        public static SlidePart InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart is null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            CommonSlideData commonSlideData = slide.CommonSlideData ?? slide.AppendChild(new CommonSlideData());
            ShapeTree shapeTree = commonSlideData.ShapeTree ?? commonSlideData.AppendChild(new ShapeTree());
            NonVisualGroupShapeProperties nonVisualProperties = shapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            shapeTree.AppendChild(new GroupShapeProperties());
            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = shapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));
            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = shapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph());
            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            
            // Assign the new slide to the new slide part
            slidePart.Slide = slide;

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId? prevSlideId = null;

            OpenXmlElementList slideIds = slideIdList?.ChildElements ?? default;

            foreach (SlideId slideId in slideIds)
            {
                if (slideId.Id is not null && slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId is not null && prevSlideId.RelationshipId is not null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
            }
            else
            {
                string? firstRelId = ((SlideId)slideIds[0]).RelationshipId;
                // If the first slide does not contain a relationship ID, throw an exception.
                if (firstRelId is null)
                {
                    throw new ArgumentNullException(nameof(firstRelId));
                }

                lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
            }

            // Use the same slide layout as that of the previous slide.
            if (lastSlidePart.SlideLayoutPart is not null)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            return slidePart;
        }
```

### [Visual Basic](#tab/vb)
```vb
        ' Insert a slide into the specified presentation.
        Public Shared Sub InsertNewSlide(presentationFile As String, position As Integer, slideTitle As String)
            ' Open the source document as read/write. 
            Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
                ' Pass the source document and the position and title of the slide to be inserted to the next method.
                InsertNewSlide(presentationDocument, position, slideTitle)
            End Using
        End Sub
        ' Insert the specified slide into the presentation at the specified position.
        Public Shared Function InsertNewSlide(presentationDocument As PresentationDocument, position As Integer, slideTitle As String) As SlidePart
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Verify that the presentation is not empty.
            If presentationPart Is Nothing Then
                Throw New InvalidOperationException("The presentation document is empty.")
            End If

            ' Declare and instantiate a new slide.
            Dim slide As New Slide(New CommonSlideData(New ShapeTree()))
            Dim drawingObjectId As UInteger = 1

            ' Construct the slide content.            
            ' Specify the non-visual properties of the new slide.
            Dim commonSlideData As CommonSlideData = If(slide.CommonSlideData, slide.AppendChild(New CommonSlideData()))
            Dim shapeTree As ShapeTree = If(commonSlideData.ShapeTree, commonSlideData.AppendChild(New ShapeTree()))
            Dim nonVisualProperties As NonVisualGroupShapeProperties = shapeTree.AppendChild(New NonVisualGroupShapeProperties())
            nonVisualProperties.NonVisualDrawingProperties = New NonVisualDrawingProperties() With {.Id = 1, .Name = ""}
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = New NonVisualGroupShapeDrawingProperties()
            nonVisualProperties.ApplicationNonVisualDrawingProperties = New ApplicationNonVisualDrawingProperties()

            ' Specify the group shape properties of the new slide.
            shapeTree.AppendChild(New GroupShapeProperties())
            ' Declare and instantiate the title shape of the new slide.
            Dim titleShape As Shape = shapeTree.AppendChild(New Shape())

            drawingObjectId += 1

            ' Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = New NonVisualShapeProperties(
                New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Title"},
                New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}),
                New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title}))
            titleShape.ShapeProperties = New ShapeProperties()

            ' Specify the text of the title shape.
            titleShape.TextBody = New TextBody(New Drawing.BodyProperties(),
                    New Drawing.ListStyle(),
                    New Drawing.Paragraph(New Drawing.Run(New Drawing.Text() With {.Text = slideTitle})))
            ' Declare and instantiate the body shape of the new slide.
            Dim bodyShape As Shape = shapeTree.AppendChild(New Shape())
            drawingObjectId += 1

            ' Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = New NonVisualShapeProperties(New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Content Placeholder"},
                    New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}),
                    New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Index = 1}))
            bodyShape.ShapeProperties = New ShapeProperties()

            ' Specify the text of the body shape.
            bodyShape.TextBody = New TextBody(New Drawing.BodyProperties(),
                    New Drawing.ListStyle(),
                    New Drawing.Paragraph())
            ' Create the slide part for the new slide.
            Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)()

            ' Assign the new slide to the new slide part
            slidePart.Slide = slide

            ' Modify the slide ID list in the presentation part.
            ' The slide ID list should not be null.
            Dim slideIdList As SlideIdList = presentationPart.Presentation.SlideIdList

            ' Find the highest slide ID in the current list.
            Dim maxSlideId As UInteger = 1
            Dim prevSlideId As SlideId = Nothing

            Dim slideIds As OpenXmlElementList = slideIdList?.ChildElements

            For Each slideId As SlideId In slideIds
                If slideId.Id IsNot Nothing AndAlso slideId.Id.Value > maxSlideId Then
                    maxSlideId = slideId.Id
                End If

                position -= 1
                If position = 0 Then
                    prevSlideId = slideId
                End If
            Next

            maxSlideId += 1

            ' Get the ID of the previous slide.
            Dim lastSlidePart As SlidePart

            If prevSlideId IsNot Nothing AndAlso prevSlideId.RelationshipId IsNot Nothing Then
                lastSlidePart = CType(presentationPart.GetPartById(prevSlideId.RelationshipId), SlidePart)
            Else
                Dim firstRelId As String = CType(slideIds(0), SlideId).RelationshipId
                ' If the first slide does not contain a relationship ID, throw an exception.
                If firstRelId Is Nothing Then
                    Throw New ArgumentNullException(NameOf(firstRelId))
                End If

                lastSlidePart = CType(presentationPart.GetPartById(firstRelId), SlidePart)
            End If

            ' Use the same slide layout as that of the previous slide.
            If lastSlidePart.SlideLayoutPart IsNot Nothing Then
                slidePart.AddPart(lastSlidePart.SlideLayoutPart)
            End If

            ' Insert the new slide into the slide list after the previous slide.
            Dim newSlideId As SlideId = slideIdList.InsertAfter(New SlideId(), prevSlideId)
            newSlideId.Id = maxSlideId
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart)
            Return slidePart
        End Function
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
