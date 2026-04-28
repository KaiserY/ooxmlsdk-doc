# Get all the text in all slides in a presentation

This topic shows how to use the classes in the Open XML SDK to get
all of the text in all of the slides in a presentation programmatically.

--------------------------------------------------------------------------------
## Getting a PresentationDocument object 

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open`
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read/write access, assign the value `true` to this parameter; for read-only access
assign it the value `false` as shown in the
following `using` statement. In this code,
the `presentationFile` parameter is a string
that represents the path for the file from which you want to open the
document.

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

## Sample Code 
The following code gets all the text in all the slides in a specific
presentation file. For example, you can pass the name of the file as an argument, 
and then use a `foreach` loop in your program to get the array of
strings returned by the method `GetSlideIdAndText` as shown in the following
example.

### [C#](#tab/cs-2)
```csharp
if (args is [{ } path])
{
    int numberOfSlides = CountSlides(path);
    Console.WriteLine($"Number of slides = {numberOfSlides}");

    for (int i = 0;  i < numberOfSlides; i++)
    {
        GetSlideIdAndText(out string text, path, i);
        Console.WriteLine($"Side #{i + 1} contains: {text}");
    }
}
```

### [Visual Basic](#tab/vb-2)
```vb
        If args.Length = 1 Then
            Dim path As String = args(0)
            Dim numberOfSlides As Integer = CountSlides(path)
            Console.WriteLine($"Number of slides = {numberOfSlides}")

            For i As Integer = 0 To numberOfSlides - 1
                Dim text As String = String.Empty
                GetSlideIdAndText(text, path, i)
                Console.WriteLine($"Slide #{i + 1} contains: {text}")
            Next
        End If
```
***

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static int CountSlides(string presentationFile)
{
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
```

### [Visual Basic](#tab/vb)
```vb
    Function CountSlides(presentationFile As String) As Integer
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
```
***

--------------------------------------------------------------------------------
## See also 

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
