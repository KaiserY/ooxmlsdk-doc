# Delete a slide from a presentation

This topic shows how to use the Open XML SDK for Office to delete a
slide from a presentation programmatically. It also shows how to delete
all references to the slide from any custom shows that may exist. To
delete a specific slide in a presentation file you need to know first
the number of slides in the presentation. Therefore the code in this
how-to is divided into two parts. The first is counting the number of
slides, and the second is deleting a slide at a specific index.

> **Note**
> Deleting a slide from more complex presentations, such as those that contain outline view settings, for example, may require additional steps.

--------------------------------------------------------------------------------
## Getting a Presentation Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a presentation document package. To work with a presentation document, first create an instance of the `PresentationDocument` class, and then work with that instance. To create the class instance from the document call one of the `Open` method overloads. The code in this topic uses the `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open` method, which takes a file path as the first parameter to specify the file to open, and a Boolean value as the second parameter to specify whether a document is editable. Set this second parameter to `false` to open the file for read-only access, or `true` if you want to open the file for read/write access. The code in this topic opens the file twice, once to count the number of slides and once to delete a specific slide. When you count the number of slides in a presentation, it is best to open the file for read-only access to protect the file against accidental writing. The following `using` statement opens the file for read-only access. In this code example, the `presentationFile` parameter is a string that represents the path for the file from which you want to open the document.

### [C#](#tab/cs-1)
```csharp
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
```

### [Visual Basic](#tab/vb-1)
```vb
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
```
***

To delete a slide from the presentation file, open it for read/write
access as shown in the following `using`
statement.

### [C#](#tab/cs-2)
```csharp
        // Open the source document as read/write.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
```

### [Visual Basic](#tab/vb-2)
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

## Counting the Number of Slides

The sample code consists of two overloads of the `CountSlides` method. The first overload uses a `string` parameter and the second overload uses a `PresentationDocument` parameter. In the first `CountSlides` method, the sample code opens the presentation document in the `using` statement. Then it passes the `PresentationDocument` object to the second `CountSlides` method, which returns an integer number that represents the number of slides in the presentation.

### [C#](#tab/cs-3)
```csharp
            // Pass the presentation to the next CountSlide method
            // and return the slide count.
            return CountSlides(presentationDocument);
```

### [Visual Basic](#tab/vb-3)
```vb
            ' Pass the presentation to the next CountSlide method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
```
***

In the second `CountSlides` method, the code
verifies that the `PresentationDocument`
object passed in is not `null`, and if it is
not, it gets a `PresentationPart` object from
the `PresentationDocument` object. By using
the `SlideParts` the code gets the slideCount
and returns it.

### [C#](#tab/cs-4)
```csharp
        if (presentationDocument is null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        int slidesCount = 0;

        // Get the presentation part of document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // Get the slide count from the SlideParts.
        if (presentationPart is not null)
        {
            slidesCount = presentationPart.SlideParts.Count();
        }

        // Return the slide count to the previous method.
        return slidesCount;
```

### [Visual Basic](#tab/vb-4)
```vb
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the slide count from the SlideParts.
        If presentationPart IsNot Nothing Then
            slidesCount = presentationPart.SlideParts.Count()
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
```
***

--------------------------------------------------------------------------------
## Deleting a Specific Slide

The code for deleting a slide uses two overloads of the `DeleteSlide` method. The first overloaded `DeleteSlide` method takes two parameters: a string
that represents the presentation file name and path, and an integer that
represents the zero-based index position of the slide to delete. It
opens the presentation file for read/write access, gets a `PresentationDocument` object, and then passes that
object and the index number to the next overloaded `DeleteSlide` method, which performs the deletion.

### [C#](#tab/cs-5)
```csharp
    // Get the presentation object and pass it to the next DeleteSlide method.
    static void DeleteSlide(string presentationFile, int slideIndex)
    {
        // Open the source document as read/write.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex);
        }
    }
```

### [Visual Basic](#tab/vb-5)
```vb
    ' Get the presentation object and pass it to the next DeleteSlide method.
    Private Sub DeleteSlide(presentationFile As String, slideIndex As Integer)
        ' Open the source document as read/write.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            ' Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex)
        End Using
    End Sub
```
***

The first section of the second overloaded `DeleteSlide` method uses the `CountSlides` method to get the number of slides in
the presentation. Then, it gets the list of slide IDs in the
presentation, identifies the specified slide in the slide list, and
removes the slide from the slide list.

### [C#](#tab/cs-6)
```csharp
    // Delete the specified slide from the presentation.
    static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException(nameof(presentationDocument));
        }

        // Use the CountSlides sample to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        if (slideIndex < 0 || slideIndex >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part from the presentation document. 
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // Get the presentation from the presentation part.
        Presentation? presentation = presentationPart?.Presentation;

        // Get the list of slide IDs in the presentation.
        SlideIdList? slideIdList = presentation?.SlideIdList;

        // Get the slide ID of the specified slide
        SlideId? slideId = slideIdList?.ChildElements[slideIndex] as SlideId;

        // Get the relationship ID of the slide.
        string? slideRelId = slideId?.RelationshipId;

        // If there's no relationship ID, there's no slide to delete.
        if (slideRelId is null)
        {
            return;
        }

        // Remove the slide from the slide list.
        slideIdList!.RemoveChild(slideId);
```

### [Visual Basic](#tab/vb-6)
```vb
    ' Delete the specified slide from the presentation.
    Private Sub DeleteSlide(presentationDocument As PresentationDocument, slideIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        If slideIndex < 0 OrElse slideIndex >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(slideIndex))
        End If

        ' Get the presentation part from the presentation document. 
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part.
        Dim presentation As Presentation = presentationPart?.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation?.SlideIdList

        ' Get the slide ID of the specified slide
        Dim slideId As SlideId = TryCast(slideIdList?.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the slide.
        Dim slideRelId As String = slideId?.RelationshipId

        ' If there's no relationship ID, there's no slide to delete.
        If slideRelId Is Nothing Then
            Return
        End If

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
```
***

The next section of the second overloaded `DeleteSlide` method removes all references to the
deleted slide from custom shows. It does that by iterating through the
list of custom shows and through the list of slides in each custom show.
It then declares and instantiates a linked list of slide list entries,
and finds references to the deleted slide by using the relationship ID
of that slide. It adds those references to the list of slide list
entries, and then removes each such reference from the slide list of its
respective custom show.

### [C#](#tab/cs-7)
```csharp
        // Remove references to the slide from all custom shows.
        if (presentation!.CustomShowList is not null)
        {
            // Iterate through the list of custom shows.
            foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
            {
                if (customShow.SlideList is not null)
                {
                    // Declare a link list of slide list entries.
                    LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                    {
                        // Find the slide reference to remove from the custom show.
                        if (slideListEntry.Id is not null && slideListEntry.Id == slideRelId)
                        {
                            slideListEntries.AddLast(slideListEntry);
                        }
                    }

                    // Remove all references to the slide from the custom show.
                    foreach (SlideListEntry slideListEntry in slideListEntries)
                    {
                        customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }
        }
```

### [Visual Basic](#tab/vb-7)
```vb
        ' Remove references to the slide from all custom shows.
        If presentation.CustomShowList IsNot Nothing Then
            ' Iterate through the list of custom shows.
            For Each customShow In presentation.CustomShowList.Elements(Of CustomShow)()
                If customShow.SlideList IsNot Nothing Then
                    ' Declare a link list of slide list entries.
                    Dim slideListEntries As New LinkedList(Of SlideListEntry)()
                    For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements()
                        ' Find the slide reference to remove from the custom show.
                        If slideListEntry.Id IsNot Nothing AndAlso slideListEntry.Id = slideRelId Then
                            slideListEntries.AddLast(slideListEntry)
                        End If
                    Next

                    ' Remove all references to the slide from the custom show.
                    For Each slideListEntry As SlideListEntry In slideListEntries
                        customShow.SlideList.RemoveChild(slideListEntry)
                    Next
                End If
            Next
        End If
```
***

Finally, the code deletes the slide part for the deleted slide.

### [C#](#tab/cs-8)
```csharp
        // Get the slide part for the specified slide.
        SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelId);

        // Remove the slide part.
        presentationPart.DeletePart(slidePart);
```

### [Visual Basic](#tab/vb-8)
```vb
        ' Get the slide part for the specified slide.
        Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideRelId), SlidePart)

        ' Remove the slide part.
        presentationPart.DeletePart(slidePart)
```
***

--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
    // Get the presentation object and pass it to the next CountSlides method.
    static int CountSlides(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation to the next CountSlide method
            // and return the slide count.
            return CountSlides(presentationDocument);
        }
    }

    // Count the slides in the presentation.
    static int CountSlides(PresentationDocument presentationDocument)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        int slidesCount = 0;

        // Get the presentation part of document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // Get the slide count from the SlideParts.
        if (presentationPart is not null)
        {
            slidesCount = presentationPart.SlideParts.Count();
        }

        // Return the slide count to the previous method.
        return slidesCount;
    }
    // Get the presentation object and pass it to the next DeleteSlide method.
    static void DeleteSlide(string presentationFile, int slideIndex)
    {
        // Open the source document as read/write.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex);
        }
    }
    // Delete the specified slide from the presentation.
    static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException(nameof(presentationDocument));
        }

        // Use the CountSlides sample to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        if (slideIndex < 0 || slideIndex >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part from the presentation document. 
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // Get the presentation from the presentation part.
        Presentation? presentation = presentationPart?.Presentation;

        // Get the list of slide IDs in the presentation.
        SlideIdList? slideIdList = presentation?.SlideIdList;

        // Get the slide ID of the specified slide
        SlideId? slideId = slideIdList?.ChildElements[slideIndex] as SlideId;

        // Get the relationship ID of the slide.
        string? slideRelId = slideId?.RelationshipId;

        // If there's no relationship ID, there's no slide to delete.
        if (slideRelId is null)
        {
            return;
        }

        // Remove the slide from the slide list.
        slideIdList!.RemoveChild(slideId);
        // Remove references to the slide from all custom shows.
        if (presentation!.CustomShowList is not null)
        {
            // Iterate through the list of custom shows.
            foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
            {
                if (customShow.SlideList is not null)
                {
                    // Declare a link list of slide list entries.
                    LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                    {
                        // Find the slide reference to remove from the custom show.
                        if (slideListEntry.Id is not null && slideListEntry.Id == slideRelId)
                        {
                            slideListEntries.AddLast(slideListEntry);
                        }
                    }

                    // Remove all references to the slide from the custom show.
                    foreach (SlideListEntry slideListEntry in slideListEntries)
                    {
                        customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }
        }
        // Get the slide part for the specified slide.
        SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelId);

        // Remove the slide part.
        presentationPart.DeletePart(slidePart);
    }
```

### [Visual Basic](#tab/vb)
```vb
    ' Get the presentation object and pass it to the next CountSlides method.
    Private Function CountSlides(presentationFile As String) As Integer
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' Pass the presentation to the next CountSlide method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
        End Using
    End Function

    ' Count the slides in the presentation.
    Private Function CountSlides(presentationDocument As PresentationDocument) As Integer
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the slide count from the SlideParts.
        If presentationPart IsNot Nothing Then
            slidesCount = presentationPart.SlideParts.Count()
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
    End Function
    ' Get the presentation object and pass it to the next DeleteSlide method.
    Private Sub DeleteSlide(presentationFile As String, slideIndex As Integer)
        ' Open the source document as read/write.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            ' Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex)
        End Using
    End Sub
    ' Delete the specified slide from the presentation.
    Private Sub DeleteSlide(presentationDocument As PresentationDocument, slideIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        If slideIndex < 0 OrElse slideIndex >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(slideIndex))
        End If

        ' Get the presentation part from the presentation document. 
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part.
        Dim presentation As Presentation = presentationPart?.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation?.SlideIdList

        ' Get the slide ID of the specified slide
        Dim slideId As SlideId = TryCast(slideIdList?.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the slide.
        Dim slideRelId As String = slideId?.RelationshipId

        ' If there's no relationship ID, there's no slide to delete.
        If slideRelId Is Nothing Then
            Return
        End If

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
        ' Remove references to the slide from all custom shows.
        If presentation.CustomShowList IsNot Nothing Then
            ' Iterate through the list of custom shows.
            For Each customShow In presentation.CustomShowList.Elements(Of CustomShow)()
                If customShow.SlideList IsNot Nothing Then
                    ' Declare a link list of slide list entries.
                    Dim slideListEntries As New LinkedList(Of SlideListEntry)()
                    For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements()
                        ' Find the slide reference to remove from the custom show.
                        If slideListEntry.Id IsNot Nothing AndAlso slideListEntry.Id = slideRelId Then
                            slideListEntries.AddLast(slideListEntry)
                        End If
                    Next

                    ' Remove all references to the slide from the custom show.
                    For Each slideListEntry As SlideListEntry In slideListEntries
                        customShow.SlideList.RemoveChild(slideListEntry)
                    Next
                End If
            Next
        End If
        ' Get the slide part for the specified slide.
        Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideRelId), SlidePart)

        ' Remove the slide part.
        presentationPart.DeletePart(slidePart)
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
