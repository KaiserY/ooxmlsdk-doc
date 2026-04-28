# Apply a theme to a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to apply the theme from one presentation to another presentation
programmatically.

-----------------------------------------------------------------------------
## Getting a PresentationDocument Object
In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)` method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read-only access,
specify the value **false** for this parameter.
To open a document for read/write access, specify the value **true** for this parameter. In the following **using** statement, two presentation files are
opened, the target presentation, to which to apply a theme, and the
source presentation, which already has that theme applied. The source
presentation file is opened for read-only access, and the target
presentation file is opened for read/write access. In this code, the
**themePresentation** parameter is a string
that represents the path for the source presentation document, and the
**presentationFile** parameter is a string that
represents the path for the target presentation document.

### [C#](#tab/cs-0)
```csharp
    using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, false))
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
```

### [Visual Basic](#tab/vb-0)
```vb
        Dim themeDocument As PresentationDocument = PresentationDocument.Open(themePresentation, False)
        Dim presentationDoc As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        Using (themeDocument)
            Using (presentationDoc)
```
***

 With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
This ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `themeDocument` and `presentationDocument`.

-----------------------------------------------------------------------------

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

-----------------------------------------------------------------------------
## Structure of the Theme Element
The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification can
be useful when working with this element.

> This element defines the root level complex type associated with a
> shared style sheet (or theme). This element holds all the different
> formatting options available to a document through a theme, and
> defines the overall look and feel of the document when themed objects
> are used within the document. [*Example*: Consider the following image
> as an example of different themes in use applied to a presentation. In
> this example, you can see how a theme can affect font, colors,
> backgrounds, fills, and effects for different objects in a
> presentation. end example]

![Theme sample](../media/a-theme01.gif)
> In this example, we see how a theme can affect font, colors,
> backgrounds, fills, and effects for different objects in a
> presentation. *end example*]
> 
> &copy; ISO/IEC 29500: 2016

The following table lists the possible child types of the Theme class.

| PresentationML Element | Open XML SDK Class | Description |
|---|---|---|
| custClrLst | CustomColorList | Custom Color List |
| extLst | ExtensionList | Extension List |
| extraClrSchemeLst | ExtraColorSchemeList | Extra Color Scheme List |
| objectDefaults | ObjectDefaults | Object Defaults |
| themeElements | ThemeElements | Theme Elements |

The following XML Schema fragment defines the four parts of the theme
element. The **themeElements** element is the
piece that holds the main formatting defined within the theme. The other
parts provide overrides, defaults, and additions to the information
contained in **themeElements**. The complex
type defining a theme, **CT\_OfficeStyleSheet**, is defined in the following
manner.

```xml
    <complexType name="CT_OfficeStyleSheet">
       <sequence>
           <element name="themeElements" type="CT_BaseStyles" minOccurs="1" maxOccurs="1"/>
           <element name="objectDefaults" type="CT_ObjectStyleDefaults" minOccurs="0" maxOccurs="1"/>
           <element name="extraClrSchemeLst" type="CT_ColorSchemeList" minOccurs="0" maxOccurs="1"/>
           <element name="custClrLst" type="CT_CustomColorList" minOccurs="0" maxOccurs="1"/>
           <element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
       </sequence>
       <attribute name="name" type="xsd:string" use="optional" default=""/>
    </complexType>
```

This complex type also holds a **CT\_OfficeArtExtensionList**, which is used for
future extensibility of this complex type.

-----------------------------------------------------------------------------
## How the Sample Code Works
The sample code consists of two overloads of the method **ApplyThemeToPresentation**, and the **GetSlideLayoutType** method. The following code
segment shows the first overloaded method, in which the two presentation
files, **themePresentation** and **presentationFile**, are opened and passed to the
second overloaded method as parameters.

### [C#](#tab/cs-1)
```csharp
// Apply a new theme to the presentation. 
static void ApplyThemeToPresentation(string presentationFile, string themePresentation)
{
    using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, false))
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        ApplyThemeToPresentationDocument(presentationDocument, themeDocument);
    }
}
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationFile As String, ByVal themePresentation As String)
        Dim themeDocument As PresentationDocument = PresentationDocument.Open(themePresentation, False)
        Dim presentationDoc As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        Using (themeDocument)
            Using (presentationDoc)
                ApplyThemeToPresentation(presentationDoc, themeDocument)
            End Using
        End Using

    End Sub
```
***

In the second overloaded method, the code starts by checking whether any
of the presentation files is empty, in which case it throws an
exception. The code then gets the presentation part of the presentation
document by declaring a **PresentationPart**
object and setting it equal to the presentation part of the target **PresentationDocument** object passed in. It then
gets the slide master parts from the presentation parts of both objects
passed in, and gets the relationship ID of the slide master part of the
target presentation.

### [C#](#tab/cs-2)
```csharp
// Apply a new theme to the presentation. 
static void ApplyThemeToPresentationDocument(PresentationDocument presentationDocument, PresentationDocument themeDocument)
{
    // Get the presentation part of the presentation document.
    PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();

    // Get the existing slide master part.
    SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.ElementAt(0);

    string relationshipId = presentationPart.GetIdOfPart(slideMasterPart);

    // Get the new slide master part.
    PresentationPart themeDocPresentationPart = themeDocument.PresentationPart ?? themeDocument.AddPresentationPart();
    SlideMasterPart newSlideMasterPart = themeDocPresentationPart.SlideMasterParts.ElementAt(0);
```

### [Visual Basic](#tab/vb-2)
```vb
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationDocument As PresentationDocument, ByVal themeDocument As PresentationDocument)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        If (themeDocument Is Nothing) Then
            Throw New ArgumentNullException("themeDocument")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the existing slide master part.
        Dim slideMasterPart As SlideMasterPart = presentationPart.SlideMasterParts.ElementAt(0)

        Dim relationshipId As String = presentationPart.GetIdOfPart(slideMasterPart)

        ' Get the new slide master part.
        Dim newSlideMasterPart As SlideMasterPart = themeDocument.PresentationPart.SlideMasterParts.ElementAt(0)
```
***

The code then removes the existing theme part and the slide master part
from the target presentation. By reusing the old relationship ID, it
adds the new slide master part from the source presentation to the
target presentation. It also adds the theme part to the target
presentation.

### [C#](#tab/cs-3)
```csharp
    if (presentationPart.ThemePart is not null)
    {
        // Remove the existing theme part.
        presentationPart.DeletePart(presentationPart.ThemePart);
    }

    // Remove the old slide master part.
    presentationPart.DeletePart(slideMasterPart);

    // Import the new slide master part, and reuse the old relationship ID.
    newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId);

    if (newSlideMasterPart.ThemePart is not null)
    {
        // Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart);
    }
```

### [Visual Basic](#tab/vb-3)
```vb
        ' Remove the theme part.
        presentationPart.DeletePart(presentationPart.ThemePart)

        ' Remove the old slide master part.
        presentationPart.DeletePart(slideMasterPart)

        ' Import the new slide master part, and reuse the old relationship ID.
        newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId)

        ' Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart)
```
***

The code iterates through all the slide layout parts in the slide master
part and adds them to the list of new slide layouts. It specifies the
default layout type. For this example, the code for the default layout
type is "Title and Content".

### [C#](#tab/cs-4)
```csharp
    Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();

    foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
    {
        string? newLayoutType = GetSlideLayoutType(slideLayoutPart);

        if (newLayoutType is not null)
        {
            newSlideLayouts.Add(newLayoutType, slideLayoutPart);
        }
    }

    string? layoutType = null;
    SlideLayoutPart? newLayoutPart = null;

    // Insert the code for the layout for this example.
    string defaultLayoutType = "Title and Content";
```

### [Visual Basic](#tab/vb-4)
```vb
        Dim newSlideLayouts As Dictionary(Of String, SlideLayoutPart) = New Dictionary(Of String, SlideLayoutPart)()
        For Each slideLayoutPart As Object In newSlideMasterPart.SlideLayoutParts
            newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart)
        Next
        Dim layoutType As String = Nothing
        Dim newLayoutPart As SlideLayoutPart = Nothing

        ' Insert the code for the layout for this example.
        Dim defaultLayoutType As String = "Title and Content"
```
***

The code iterates through all the slide parts in the target presentation
and removes the slide layout relationship on all slides. It uses the
**GetSlideLayoutType** method to find the
layout type of the slide layout part. For any slide with an existing
slide layout part, it adds a new slide layout part of the same type it
had previously. For any slide without an existing slide layout part, it
adds a new slide layout part of the default type.

### [C#](#tab/cs-5)
```csharp
    // Remove the slide layout relationship on all slides. 
    foreach (var slidePart in presentationPart.SlideParts)
    {
        layoutType = null;

        if (slidePart.SlideLayoutPart is not null)
        {
            // Determine the slide layout type for each slide.
            layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);

            // Delete the old layout part.
            slidePart.DeletePart(slidePart.SlideLayoutPart);
        }

        if (layoutType is not null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
        {
            // Apply the new layout part.
            slidePart.AddPart(newLayoutPart);
        }
        else
        {
            newLayoutPart = newSlideLayouts[defaultLayoutType];

            // Apply the new default layout part.
            slidePart.AddPart(newLayoutPart);
        }
    }
```

### [Visual Basic](#tab/vb-5)
```vb
        Dim newSlideLayouts As Dictionary(Of String, SlideLayoutPart) = New Dictionary(Of String, SlideLayoutPart)()
        For Each slideLayoutPart As Object In newSlideMasterPart.SlideLayoutParts
            newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart)
        Next
        Dim layoutType As String = Nothing
        Dim newLayoutPart As SlideLayoutPart = Nothing

        ' Insert the code for the layout for this example.
        Dim defaultLayoutType As String = "Title and Content"
        ' Remove the slide layout relationship on all slides. 
        For Each slidePart As Object In presentationPart.SlideParts
            layoutType = Nothing
            If ((slidePart.SlideLayoutPart) IsNot Nothing) Then

                ' Determine the slide layout type for each slide.
                layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart)

                ' Delete the old layout part.
                slidePart.DeletePart(slidePart.SlideLayoutPart)
            End If

            If (((layoutType) IsNot Nothing) AndAlso newSlideLayouts.TryGetValue(layoutType, newLayoutPart)) Then

                ' Apply the new layout part.
                slidePart.AddPart(newLayoutPart)
            Else
                newLayoutPart = newSlideLayouts(defaultLayoutType)

                ' Apply the new default layout part.
                slidePart.AddPart(newLayoutPart)
            End If
```
***

To get the type of the slide layout, the code uses the **GetSlideLayoutType** method that takes the slide
layout part as a parameter, and returns to the second overloaded **ApplyThemeToPresentation** method a string that
represents the name of the slide layout type

### [C#](#tab/cs-6)
```csharp
// Get the slide layout type.
static string? GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
{
    CommonSlideData? slideData = slideLayoutPart.SlideLayout?.CommonSlideData;

    return slideData?.Name;
}
```

### [Visual Basic](#tab/vb-6)
```vb
    ' Get the type of the slide layout.
    Public Function GetSlideLayoutType(ByVal slideLayoutPart As SlideLayoutPart) As String
        Dim slideData As CommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData

        ' Remarks: If this is used in production code, check for a null reference.
        Return slideData.Name
    End Function
```
***
-----------------------------------------------------------------------------
## Sample Code
The following is the complete sample code to copy a theme from one
presentation to another. To use the program, you must create two
presentations, a source presentation with the theme you would like to
copy, for example, Myppt9-theme.pptx, and the other one is the target
presentation, for example, Myppt9.pptx. You can use the following call
in your program to perform the copying.

### [C#](#tab/cs-7)
```csharp
ApplyThemeToPresentation(args[0], args[1]);
```

### [Visual Basic](#tab/vb-7)
```vb
    Sub Main(args As String())
        ApplyThemeToPresentation(args(0), args(1))
    End Sub
```
***

After performing that call you can inspect the file Myppt2.pptx, and you
would see the same theme of the file Myppt9-theme.pptx.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.Linq;
ApplyThemeToPresentation(args[0], args[1]);
// Apply a new theme to the presentation. 
static void ApplyThemeToPresentation(string presentationFile, string themePresentation)
{
    using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, false))
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        ApplyThemeToPresentationDocument(presentationDocument, themeDocument);
    }
}
// Apply a new theme to the presentation. 
static void ApplyThemeToPresentationDocument(PresentationDocument presentationDocument, PresentationDocument themeDocument)
{
    // Get the presentation part of the presentation document.
    PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();

    // Get the existing slide master part.
    SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.ElementAt(0);

    string relationshipId = presentationPart.GetIdOfPart(slideMasterPart);

    // Get the new slide master part.
    PresentationPart themeDocPresentationPart = themeDocument.PresentationPart ?? themeDocument.AddPresentationPart();
    SlideMasterPart newSlideMasterPart = themeDocPresentationPart.SlideMasterParts.ElementAt(0);
    if (presentationPart.ThemePart is not null)
    {
        // Remove the existing theme part.
        presentationPart.DeletePart(presentationPart.ThemePart);
    }

    // Remove the old slide master part.
    presentationPart.DeletePart(slideMasterPart);

    // Import the new slide master part, and reuse the old relationship ID.
    newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId);

    if (newSlideMasterPart.ThemePart is not null)
    {
        // Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart);
    }
    Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();

    foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
    {
        string? newLayoutType = GetSlideLayoutType(slideLayoutPart);

        if (newLayoutType is not null)
        {
            newSlideLayouts.Add(newLayoutType, slideLayoutPart);
        }
    }

    string? layoutType = null;
    SlideLayoutPart? newLayoutPart = null;

    // Insert the code for the layout for this example.
    string defaultLayoutType = "Title and Content";
    // Remove the slide layout relationship on all slides. 
    foreach (var slidePart in presentationPart.SlideParts)
    {
        layoutType = null;

        if (slidePart.SlideLayoutPart is not null)
        {
            // Determine the slide layout type for each slide.
            layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);

            // Delete the old layout part.
            slidePart.DeletePart(slidePart.SlideLayoutPart);
        }

        if (layoutType is not null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
        {
            // Apply the new layout part.
            slidePart.AddPart(newLayoutPart);
        }
        else
        {
            newLayoutPart = newSlideLayouts[defaultLayoutType];

            // Apply the new default layout part.
            slidePart.AddPart(newLayoutPart);
        }
    }
}
// Get the slide layout type.
static string? GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
{
    CommonSlideData? slideData = slideLayoutPart.SlideLayout?.CommonSlideData;

    return slideData?.Name;
}
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module MyModule
    Sub Main(args As String())
        ApplyThemeToPresentation(args(0), args(1))
    End Sub
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationFile As String, ByVal themePresentation As String)
        Dim themeDocument As PresentationDocument = PresentationDocument.Open(themePresentation, False)
        Dim presentationDoc As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        Using (themeDocument)
            Using (presentationDoc)
                ApplyThemeToPresentation(presentationDoc, themeDocument)
            End Using
        End Using

    End Sub
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationDocument As PresentationDocument, ByVal themeDocument As PresentationDocument)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        If (themeDocument Is Nothing) Then
            Throw New ArgumentNullException("themeDocument")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the existing slide master part.
        Dim slideMasterPart As SlideMasterPart = presentationPart.SlideMasterParts.ElementAt(0)

        Dim relationshipId As String = presentationPart.GetIdOfPart(slideMasterPart)

        ' Get the new slide master part.
        Dim newSlideMasterPart As SlideMasterPart = themeDocument.PresentationPart.SlideMasterParts.ElementAt(0)
        ' Remove the theme part.
        presentationPart.DeletePart(presentationPart.ThemePart)

        ' Remove the old slide master part.
        presentationPart.DeletePart(slideMasterPart)

        ' Import the new slide master part, and reuse the old relationship ID.
        newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId)

        ' Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart)
        Dim newSlideLayouts As Dictionary(Of String, SlideLayoutPart) = New Dictionary(Of String, SlideLayoutPart)()
        For Each slideLayoutPart As Object In newSlideMasterPart.SlideLayoutParts
            newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart)
        Next
        Dim layoutType As String = Nothing
        Dim newLayoutPart As SlideLayoutPart = Nothing

        ' Insert the code for the layout for this example.
        Dim defaultLayoutType As String = "Title and Content"
        ' Remove the slide layout relationship on all slides. 
        For Each slidePart As Object In presentationPart.SlideParts
            layoutType = Nothing
            If ((slidePart.SlideLayoutPart) IsNot Nothing) Then

                ' Determine the slide layout type for each slide.
                layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart)

                ' Delete the old layout part.
                slidePart.DeletePart(slidePart.SlideLayoutPart)
            End If

            If (((layoutType) IsNot Nothing) AndAlso newSlideLayouts.TryGetValue(layoutType, newLayoutPart)) Then

                ' Apply the new layout part.
                slidePart.AddPart(newLayoutPart)
            Else
                newLayoutPart = newSlideLayouts(defaultLayoutType)

                ' Apply the new default layout part.
                slidePart.AddPart(newLayoutPart)
            End If
        Next
    End Sub
    ' Get the type of the slide layout.
    Public Function GetSlideLayoutType(ByVal slideLayoutPart As SlideLayoutPart) As String
        Dim slideData As CommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData

        ' Remarks: If this is used in production code, check for a null reference.
        Return slideData.Name
    End Function
End Module
```

-----------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
