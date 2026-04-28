# Working with presentations

This topic discusses the Open XML SDK for Office `DocumentFormat.OpenXml.Presentation.Presentation` class and how it relates to
the Open XML File Format PresentationML schema. For more information
about the overall structure of the parts and elements that make up a
PresentationML document, see [Structure of a PresentationML document](structure-of-a-presentationml-document.md).

---------------------------------------------------------------------------------
## Presentations in PresentationML
The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<presentation/>`
element used to represent a presentation in a PresentationML document as
follows:

This element specifies within it fundamental presentation-wide
properties.

**Example:** Consider the following presentation with a single slide master
and two slides. In addition to these commonly used elements there can
also be the specification of other properties such as slide size, notes
size and default text styles.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p="">  
    <p:sldMasterIdLst>  
        <p:sldMasterId id="2147483648" r:id="rId1">  
    </p:sldMasterIdLst>  
    <p:sldIdLst>  
        <p:sldId id="256" r:id="rId3"/>  
        <p:sldId id="257" r:id="rId4"/>  
    </p:sldIdLst>  
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>  
    <p:notesSz cx="6858000" cy="9144000"/>  
    <p:defaultTextStyle>  
        …  
    </p:defaultTextStyle>  
</p:presentation>
```

&copy; ISO/IEC 29500: 2016

The `<presentation/>` element typically contains child elements that list
slide masters, slides, and custom slide shows contained within the
presentation. In addition, it also commonly contains elements that
specify other properties of the presentation, such as slide size, notes
size, and default text styles.

The `<presentation/>` element is the root element of the PresentationML
Presentation part. For more information about the overall structure of
the parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists some of the most common child elements of the
`<presentation/>` element used when working with presentations and the
Open XML SDK classes that correspond to them.

| **PresentationML Element** |                                                     **Open XML SDK Class**                                                      |
|----------------------------|-------------------------------------------------------------------------------------------------------------------------------------|
|     `<sldMasterIdLst/>`     |   `DocumentFormat.OpenXml.Presentation.SlideMasterIdList`   |
|      `<sldMasterId/>`       |       `DocumentFormat.OpenXml.Presentation.SlideMasterId`       |
|        `<sldIdLst/>`        |         `DocumentFormat.OpenXml.Presentation.SlideIdList`         |
|         `<sldId/>`          |             `DocumentFormat.OpenXml.Presentation.SlideId`            |
|    `<notesMasterIdLst/>`    |   `DocumentFormat.OpenXml.Presentation.NotesMasterIdList`   |
|   `<handoutMasterIdLst/>`   | `DocumentFormat.OpenXml.Presentation.SlideMasterIdList` |
|      `<custShowLst/>`       |      `DocumentFormat.OpenXml.Presentation.CustomShowList`     |
|         `<sldSz/>`          |           `DocumentFormat.OpenXml.Presentation.SlideSize`       |
|        `<notesSz/>`         |           <xrefDocumentFormat.OpenXml.Presentation.NotesSize>          |
|    `<defaultTextStyle/>`    |    `DocumentFormat.OpenXml.Presentation.DefaultTextStyle`    |

--------------------------------------------------------------------------------
## Open XML SDK Presentation Class
The Open XML SDK `Presentation` class
represents the `<presentation/>` element defined in the Open XML File
Format schema for PresentationML documents. Use the `Presentation` class to manipulate an individual
`<presentation/>` element in a PresentationML document.

Classes commonly associated with the `Presentation` class are shown in the following
sections.

### SlideMasterIdList Class

All slides that share the same master inherit the same layout from that
master. The `SlideMasterIdList` class
corresponds to the `<sldMasterIdList/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<sldMasterIdList/>`
element used to represent a slide master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
slide master slides that are available within the corresponding
presentation. A slide master is a slide that is specifically designed to
be a template for all related child layout slides.

&copy; ISO/IEC 29500: 2016

### SlideMasterId Class

The `SlideMasterId` class corresponds to the
`<sldMasterId/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<sldMasterId/>`
element used to represent a slide master ID in a PresentationML document
as follows:

This element specifies a slide master that is available within the
corresponding presentation. A slide master is a slide that is
specifically designed to be a template for all related child layout
slides.

**Example**: Consider the following specification of a slide master within
a presentation

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldMasterIdLst>  
        <p:sldMasterId id="2147483648" r:id="rId1"/>  
    </p:sldMasterIdLst>  
    …  
</p:presentation>
```

&copy; ISO/IEC 29500: 2016

### SlideIdList Class

The `SlideIdList` class corresponds to the
`<sldIdLst/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<sldIdLst/>` element
used to represent a slide ID list in a PresentationML document as
follows:

This element specifies a list of identification information for the
slides that are available within the corresponding presentation. A slide
contains the information that is specific to a single slide such as
slide-specific shape and text information.

&copy; ISO/IEC 29500: 2016

### SlideId Class

The `SlideId` class corresponds to the
`<sldId/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<sldId/>` element
used to represent a slide ID in a PresentationML document as follows:

This element specifies a presentation slide that is available within the
corresponding presentation. A slide contains the information that is
specific to a single slide such as slide-specific shape and text
information.

**Example**: Consider the following specification of a slide master within
a presentation

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p="" embedTrueTypeFonts="1">  
    …  
    <p:sldIdLst>  
        <p:sldId id="256" r:id="rId3"/>  
        <p:sldId id="257" r:id="rId4"/>  
        <p:sldId id="258" r:id="rId5"/>  
        <p:sldId id="259" r:id="rId6"/>  
        <p:sldId id="260" r:id="rId7"/>  
    </p:sldIdLst>  
    ...  
</p:presentation>
```

&copy; ISO/IEC 29500: 2016

### NotesMasterIdList Class

The `NotesMasterIdList` class corresponds to
the `<notesMasterIdLst/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<notesMasterIdLst/>`
element used to represent a notes master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
notes master slides that are available within the corresponding
presentation. A notes master is a slide that is specifically designed
for the printing of the slide along with any attached notes.

&copy; ISO/IEC 29500: 2016

### HandoutMasterIdList Class

The `HandoutMasterIdList` class corresponds
to the `<handoutMasterIdLst/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML
`<handoutMasterIdLst/>` element used to represent a handout master ID
list in a PresentationML document as follows:

This element specifies a list of identification information for the
handout master slides that are available within the corresponding
presentation. A handout master is a slide that is specifically designed
for printing as a handout.

&copy; ISO/IEC 29500: 2016

### CustomShowList Class

The `CustomShowList` class corresponds to the
`<custShowLst/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<custShowLst/>`
element used to represent a custom show list in a PresentationML
document as follows:

This element specifies a list of all custom shows that are available
within the corresponding presentation. A custom show is a defined slide
sequence that allows for the displaying of the slides with the
presentation in any arbitrary order.

&copy; ISO/IEC 29500: 2016

### SlideSize Class

The `SlideSize` class corresponds to the
`<sldSz/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<sldSz/>` element
used to represent presentation slide size in a PresentationML document
as follows:

This element specifies the size of the presentation slide surface.
Objects within a presentation slide can be specified outside these
extents, but this is the size of background surface that is shown when
the slide is presented or printed.

**Example**: Consider the following specifying of the size of a
presentation slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>  
    …  
</p:presentation>  
```

&copy; ISO/IEC 29500: 2016

### NotesSize Class

The `NotesSize` class corresponds to the
`<notesSz/>` element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<notesSz/>` element
used to represent notes slide size in a PresentationML document as
follows:

This element specifies the size of slide surface used for notes slides
and handout slides. Objects within a notes slide can be specified
outside these extents, but the notes slide has a background surface of
the specified size when presented or printed. This element is intended
to specify the region to which content is fitted in any special format
of printout the application might choose to generate, such as an outline
handout.

**Example**: Consider the following specifying of the size of a notes
slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:notesSz cx="9144000" cy="6858000"/>  
    …  
</p:presentation>
```

&copy; ISO/IEC 29500: 2016

### DefaultTextStyle Class

The `DefaultTextStyle` class corresponds to the `<defaultTextStyle/>`
element. The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<defaultTextStyle/>`
element used to represent default text style in a PresentationML
document as follows:

This element specifies the default text styles that are to be used
within the presentation. The text style defined here can be referenced
when inserting a new slide if that slide is not associated with a master
slide or if no styling information has been otherwise specified for the
text within the presentation slide.

&copy; ISO/IEC 29500: 2016

--------------------------------------------------------------------------------
## Working with the Presentation Class

As shown in the Open XML SDK code example that follows, every instance
of the `Presentation` class is associated
with an instance of the `DocumentFormat.OpenXml.Packaging.PresentationPart` class, which represents a
presentation part, one of the required parts of a PresentationML
presentation file package.

The `Presentation` class, which represents
the `<presentation/>` element, is therefore also associated with a series
of other classes that represent the child elements of the
`<presentation/>` element. Among these classes, as shown in the following
code example, are the `SlideMasterIdList`,
`SlideIdList`, `SlideSize`, `NotesSize`, and `DefaultTextStyle` classes.

--------------------------------------------------------------------------------
## Open XML SDK Code Example

The following code example from the article [How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md) uses the `DocumentFormat.OpenXml.Packaging.PresentationDocument.Create` 
method of the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class of the Open XML
SDK to create an instance of that same class that has the specified
name and file path. Then it uses the `DocumentFormat.OpenXml.Packaging.PresentationDocument.AddPresentationPart` method to add an
instance of the `DocumentFormat.OpenXml.Packaging.PresentationPart` class to the document
file. Next, it creates an instance of the `DocumentFormat.OpenXml.Presentation.Presentation` class that represents the
presentation. It passes a reference to the `PresentationPart` class instance to the
`CreatePresentationParts` procedure, which creates the other required
parts of the presentation file. The `CreatePresentation` procedure
cleans up by closing the `PresentationDocument` class instance that it
opened previously.

The `CreatePresentationParts` procedure creates instances of the `SlideMasterIdList`, `SlideIdList`, `SlideSize`, `NotesSize`, and `DefaultTextStyle` classes and appends them to the
presentation.

### [C#](#tab/cs)
```csharp
static void CreatePresentation(string filepath)
{
    // Create a presentation at a specified file path. The presentation document type is pptx, by default.
    using (PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation))
    {
        PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        CreatePresentationParts(presentationPart);
    }
}
static void CreatePresentationParts(PresentationPart presentationPart)
{
    SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
    SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
    SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
    NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
    DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

    presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

    SlidePart slidePart1;
    SlideLayoutPart slideLayoutPart1;
    SlideMasterPart slideMasterPart1;
    ThemePart themePart1;

    slidePart1 = CreateSlidePart(presentationPart);
    slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
    string slideLayoutPart1RelId = slidePart1.GetIdOfPart(slideLayoutPart1);
    slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
    themePart1 = CreateTheme(slideMasterPart1);

    slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
    presentationPart.AddPart(slideMasterPart1, "rId1");
    presentationPart.AddPart(themePart1, "rId5");
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub CreatePresentation(filepath As String)
        ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
        Using presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
            Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
            presentationPart.Presentation = New Presentation()

            CreatePresentationParts(presentationPart)
        End Using
    End Sub
    Sub CreatePresentationParts(presentationPart As PresentationPart)
        Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With {.Id = CType(2147483648UI, UInt32Value), .RelationshipId = "rId1"})
        Dim slideIdList1 As New SlideIdList(New SlideId() With {.Id = CType(256UI, UInt32Value), .RelationshipId = "rId2"})
        Dim slideSize1 As New SlideSize() With {.Cx = 9144000, .Cy = 6858000, .Type = SlideSizeValues.Screen4x3}
        Dim notesSize1 As New NotesSize() With {.Cx = 6858000, .Cy = 9144000}
        Dim defaultTextStyle1 As New DefaultTextStyle()

        presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

        Dim slidePart1 As SlidePart
        Dim slideLayoutPart1 As SlideLayoutPart
        Dim slideMasterPart1 As SlideMasterPart
        Dim themePart1 As ThemePart

        slidePart1 = CreateSlidePart(presentationPart)
        slideLayoutPart1 = CreateSlideLayoutPart(slidePart1)
        slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1)
        themePart1 = CreateTheme(slideMasterPart1)

        slideMasterPart1.AddPart(slideLayoutPart1, "rId1")
        presentationPart.AddPart(slideMasterPart1, "rId1")
        presentationPart.AddPart(themePart1, "rId5")
    End Sub
```

---------------------------------------------------------------------------------
## Resulting PresentationML
When the Open XML SDK code is run, the following XML is written to the
PresentationML document referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8" ?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
     <p:sldMasterIdLst>
      <p:sldMasterId id="2147483648" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
     </p:sldMasterIdLst>
     <p:sldIdLst>
      <p:sldId id="256" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
     </p:sldIdLst>
     <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
     <p:notesSz cx="6858000" cy="9144000"/>
     <p:defaultTextStyle/>
    </p:presentation>
```

--------------------------------------------------------------------------------
## See also

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)
