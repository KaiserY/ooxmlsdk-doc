# Move a slide to a new position in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to move a slide to a new position in a presentation
programmatically.

--------------------------------------------------------------------------------
## Getting a Presentation Object

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)` method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. In order to count the number of slides
in a presentation, it is best to open the file for read-only access in
order to avoid accidental writing to the file. To do that, specify the
value `false` for the Boolean parameter as
shown in the following `using` statement. In
this code, the `presentationFile` parameter
is a string that represents the path for the file from which you want to
open the document.

### [C#](#tab/cs-0)
```csharp
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
```

### [Visual Basic](#tab/vb-0)
```vb
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
This ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `presentationDocument`.

--------------------------------------------------------------------------------

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

In order to move a specific slide in a presentation file to a new
position, you need to know first the number of slides in the
presentation. Therefore, the code in this topic is divided into two
parts. The first is counting the number of slides, and the second is
moving a slide to a new position.

--------------------------------------------------------------------------------

## Counting the Number of Slides

The sample code for counting the number of slides consists of two
overloads of the method `CountSlides`. The
first overload uses a `string` parameter and
the second overload uses a `PresentationDocument` parameter. In the first
`CountSlides` method, the sample code opens
the presentation document in the `using`
statement. Then it passes the `PresentationDocument` object to the second `CountSlides` method, which returns an integer
number that represents the number of slides in the presentation.

### [C#](#tab/cs-1)
```csharp
            // Pass the presentation to the next CountSlides method
            // and return the slide count.
            return CountSlides(presentationDocument);
```

### [Visual Basic](#tab/vb-1)
```vb
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
```
***

In the second `CountSlides` method, the code
verifies that the `PresentationDocument`
object passed in is not `null`, and if it is
not, it gets a `PresentationPart` object from
the `PresentationDocument` object. By using
the `SlideParts` the code gets the `slideCount` and returns it.

### [C#](#tab/cs-2)
```csharp
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

### [Visual Basic](#tab/vb-2)
```vb
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

## Moving a Slide from one Position to Another

Moving a slide to a new position requires opening the file for
read/write access by specifying the value `true` to the Boolean parameter as shown in the
following `using` statement. The code for
moving a slide consists of two overloads of the `MoveSlide` method. The first overloaded `MoveSlide` method takes three parameters: a string
that represents the presentation file name and path and two integers
that represent the current index position of the slide and the index
position to which to move the slide respectively. It opens the
presentation file, gets a `PresentationDocument` object, and then passes that
object and the two integers, `from` and `to`, to the second overloaded
`MoveSlide` method, which performs the actual
move.

### [C#](#tab/cs-3)
```csharp
    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(string presentationFile, int from, int to)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            MoveSlide(presentationDocument, from, to);
        }
    }
```

### [Visual Basic](#tab/vb-3)
```vb
    ' Move a slide to a different position in the slide order in the presentation.
    Public Shared Sub MoveSlide(presentationFile As String, from As Integer, toIndex As Integer)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            MoveSlide(presentationDocument, from, toIndex)
        End Using
    End Sub
```
***

In the second overloaded `MoveSlide` method,
the `CountSlides` method is called to get the
number of slides in the presentation. The code then checks if the
zero-based indexes, `from` and `to`, are within the range and different
from one another.

### [C#](#tab/cs-4)
```csharp
    // Move a slide to a different position in the slide order in the presentation.
    static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("from");
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException("to");
        }
```

### [Visual Basic](#tab/vb-4)
```vb
    ' Move a slide to a different position in the slide order in the presentation.
    Private Shared Sub MoveSlide(presentationDocument As PresentationDocument, from As Integer, toIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Call the CountSlides method to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If from < 0 OrElse from >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(from))
        End If

        If toIndex < 0 OrElse from >= slidesCount OrElse toIndex = from Then
            Throw New ArgumentOutOfRangeException(NameOf(toIndex))
        End If
```
***

A `PresentationPart` object is declared and
set equal to the presentation part of the `PresentationDocument` object passed in. The `PresentationPart` object is used to create a `Presentation` object, and then create a `SlideIdList` object that represents the list of
slides in the presentation from the `Presentation` object. A slide ID of the source
slide (the slide to move) is obtained, and then the position of the
target slide (the slide after which in the slide order to move the
source slide) is identified.

### [C#](#tab/cs-5)
```csharp
        // Get the presentation part from the presentation document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        Presentation? presentation = presentationPart?.Presentation;

        if (presentation is null)
        {
            throw new ArgumentNullException(nameof(presentation));
        }

        SlideIdList? slideIdList = presentation.SlideIdList;

        if (slideIdList is null)
        {
            throw new ArgumentNullException(nameof(slideIdList));
        }

        // Get the slide ID of the source slide.
        SlideId? sourceSlide = slideIdList.ChildElements[from] as SlideId;

        if (sourceSlide is null)
        {
            throw new ArgumentNullException(nameof(sourceSlide));
        }

        SlideId? targetSlide = null;

        // Identify the position of the target slide after which to move the source slide.
        if (to == 0)
        {
            targetSlide = null;
        }
        else if (from < to)
        {
            targetSlide = slideIdList.ChildElements[to] as SlideId;
        }
        else
        {
            targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
        }
```

### [Visual Basic](#tab/vb-5)
```vb
        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' The slide count is not zero, so the presentation must contain slides.            
        Dim presentation As Presentation = presentationPart?.Presentation

        If presentation Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentation))
        End If

        Dim slideIdList As SlideIdList = presentation.SlideIdList

        If slideIdList Is Nothing Then
            Throw New ArgumentNullException(NameOf(slideIdList))
        End If

        ' Get the slide ID of the source slide.
        Dim sourceSlide As SlideId = TryCast(slideIdList.ChildElements(from), SlideId)

        If sourceSlide Is Nothing Then
            Throw New ArgumentNullException(NameOf(sourceSlide))
        End If

        Dim targetSlide As SlideId = Nothing

        ' Identify the position of the target slide after which toIndex move the source slide.
        If toIndex = 0 Then
            targetSlide = Nothing
        ElseIf from < toIndex Then
            targetSlide = TryCast(slideIdList.ChildElements(toIndex), SlideId)
        Else
            targetSlide = TryCast(slideIdList.ChildElements(toIndex - 1), SlideId)
        End If
```
***

The `Remove` method of the `SlideID` object is used to remove the source slide
from its current position, and then the `InsertAfter` method of the `SlideIdList` object is used to insert the source
slide in the index position after the target slide. Finally, the
modified presentation is saved.

### [C#](#tab/cs-6)
```csharp
        // Remove the source slide from its current position.
        sourceSlide.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide);
```

### [Visual Basic](#tab/vb-6)
```vb
        ' Remove the source slide from its current position.
        sourceSlide.Remove()

        ' Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide)
```
***

--------------------------------------------------------------------------------
## Sample Code
Following is the complete sample code that you can use to move a slide
from one position to another in the same presentation file in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
    // Counting the slides in the presentation.
    public static int CountSlides(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation to the next CountSlides method
            // and return the slide count.
            return CountSlides(presentationDocument);
        }
    }

    // Count the slides in the presentation.
    static int CountSlides(PresentationDocument presentationDocument)
    {
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
    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(string presentationFile, int from, int to)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            MoveSlide(presentationDocument, from, to);
        }
    }
    // Move a slide to a different position in the slide order in the presentation.
    static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("from");
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException("to");
        }
        // Get the presentation part from the presentation document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        Presentation? presentation = presentationPart?.Presentation;

        if (presentation is null)
        {
            throw new ArgumentNullException(nameof(presentation));
        }

        SlideIdList? slideIdList = presentation.SlideIdList;

        if (slideIdList is null)
        {
            throw new ArgumentNullException(nameof(slideIdList));
        }

        // Get the slide ID of the source slide.
        SlideId? sourceSlide = slideIdList.ChildElements[from] as SlideId;

        if (sourceSlide is null)
        {
            throw new ArgumentNullException(nameof(sourceSlide));
        }

        SlideId? targetSlide = null;

        // Identify the position of the target slide after which to move the source slide.
        if (to == 0)
        {
            targetSlide = null;
        }
        else if (from < to)
        {
            targetSlide = slideIdList.ChildElements[to] as SlideId;
        }
        else
        {
            targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
        }
        // Remove the source slide from its current position.
        sourceSlide.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide);
    }
```

### [Visual Basic](#tab/vb)
```vb
    ' Counting the slides in the presentation.
    Public Shared Function CountSlides(presentationFile As String) As Integer
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
        End Using
    End Function

    ' Count the slides in the presentation.
    Private Shared Function CountSlides(presentationDocument As PresentationDocument) As Integer
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
    ' Move a slide to a different position in the slide order in the presentation.
    Public Shared Sub MoveSlide(presentationFile As String, from As Integer, toIndex As Integer)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            MoveSlide(presentationDocument, from, toIndex)
        End Using
    End Sub
    ' Move a slide to a different position in the slide order in the presentation.
    Private Shared Sub MoveSlide(presentationDocument As PresentationDocument, from As Integer, toIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Call the CountSlides method to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If from < 0 OrElse from >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(from))
        End If

        If toIndex < 0 OrElse from >= slidesCount OrElse toIndex = from Then
            Throw New ArgumentOutOfRangeException(NameOf(toIndex))
        End If
        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' The slide count is not zero, so the presentation must contain slides.            
        Dim presentation As Presentation = presentationPart?.Presentation

        If presentation Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentation))
        End If

        Dim slideIdList As SlideIdList = presentation.SlideIdList

        If slideIdList Is Nothing Then
            Throw New ArgumentNullException(NameOf(slideIdList))
        End If

        ' Get the slide ID of the source slide.
        Dim sourceSlide As SlideId = TryCast(slideIdList.ChildElements(from), SlideId)

        If sourceSlide Is Nothing Then
            Throw New ArgumentNullException(NameOf(sourceSlide))
        End If

        Dim targetSlide As SlideId = Nothing

        ' Identify the position of the target slide after which toIndex move the source slide.
        If toIndex = 0 Then
            targetSlide = Nothing
        ElseIf from < toIndex Then
            targetSlide = TryCast(slideIdList.ChildElements(toIndex), SlideId)
        Else
            targetSlide = TryCast(slideIdList.ChildElements(toIndex - 1), SlideId)
        End If
        ' Remove the source slide from its current position.
        sourceSlide.Remove()

        ' Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide)
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
