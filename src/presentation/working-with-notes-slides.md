# Working with notes slides

This topic discusses the Open XML SDK for Office `DocumentFormat.OpenXml.Presentation.NotesSlide` class and how it relates to the
Open XML File Format PresentationML schema.

--------------------------------------------------------------------------------

## Notes Slides in PresentationML

The [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the Open XML PresentationML `<notes/>` element
used to represent notes slides in a PresentationML document as follows:

This element specifies the existence of a notes slide along with its
corresponding data. Contained within a notes slide are all the common
slide elements along with additional properties that are specific to the
notes element.

**Example**: Consider the following PresentationML notes slide:

```xml
<p:notes>
    <p:cSld>
        …
    </p:cSld>
    …
</p:notes>
```

In the above example a notes element specifies the existence of a notes
slide with all of its parts. Notice the cSld element, which specifies
the common elements that can appear on any slide type and then any
elements specify additional non-common properties for this notes slide.

&copy; ISO/IEC 29500: 2016

The `<notes/>` element is the root element of the PresentationML Notes
Slide part. For more information about the overall structure of the
parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the `<notes/>` element
used when working with notes slides and the Open XML SDK classes
that correspond to them.

| **PresentationML Element** |    **Open XML SDK Class**    |
|----------------------------|-------------------------------------|
|       `<clrMapOvr/>`        |              `DocumentFormat.OpenXml.Presentation.ColorMapOverride`              |
|          `<cSld/>`          |               `DocumentFormat.OpenXml.Presentation.CommonSlideData`               |
|         `<extLst/>`         | `DocumentFormat.OpenXml.Presentation.ExtensionListWithModification` |

The following table from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification describes the attributes of the `<notes/>` element.

|                    **Attributes**                     | **Description**    |
|-------------------------------------------------------|---------------------|
| `showMasterPhAnim` (Show Master Placeholder Animations) | Specifies whether or not to display animations on placeholders from the master slide.<br/><br/>The possible values for this attribute are defined by the W3C XML Schema `boolean` datatype. |
|           `showMasterSp` (Show Master Shapes)           |       Specifies if shapes on the master slide should be shown on slides or not.<br/><br/>The possible values for this attribute are defined by the W3C XML Schema `boolean` datatype.       |

&copy; ISO/IEC 29500: 2016

---------------------------------------------------------------------------------

## Open XML SDK NotesSlide Class

The OXML SDK `NotesSlide` class represents
the `<notes/>` element defined in the Open XML File Format schema for
PresentationML documents. Use the `NotesSlide` class to manipulate individual
`<notes/>` elements in a PresentationML document.

Classes that represent child elements of the `<notes/>` element and that
are therefore commonly associated with the `NotesSlide` class are shown in the following list.

- ### ColorMapOverride Class

  The `ColorMapOverride` class corresponds to
  the `<clrMapOvr/>` element. The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
  specification introduces the `<clrMapOvr/>` element:

  This element provides a mechanism with which to override the color
  schemes listed within the `<ClrMap/>` element. If the
  `<masterClrMapping/>` child element is present, the color scheme defined
  by the master is used. If the `<overrideClrMapping/>` child element is
  present, it defines a new color scheme specific to the parent notes
  slide, presentation slide, or slide layout.

&copy; ISO/IEC 29500: 2016

- ### CommonSlideData Class

  The `CommonSlideData` class corresponds to
  the `<cSld/>` element. The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
  specification introduces the `<cSld/>` element:

  This element specifies a container for the type of slide information
  that is relevant to all of the slide types. All slides share a common
  set of properties that is independent of the slide type; the description
  of these properties for any particular slide is stored within the
  slide's `<cSld/>` container. Slide data specific to the slide type
  indicated by the parent element is stored elsewhere.

  The actual data in `<cSld/>` describe only the particular parent slide;
  it is only the type of information stored that is common across all
  slides.

  &copy; ISO/IEC 29500: 2016

- ### ExtensionListWithModification Class

  The `ExtensionListWithModification` class
  corresponds to the `<extLst/>`element. The following information from the
  [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification introduces the `<extLst/>` element:

  This element specifies the extension list with modification ability
  within which all future extensions of element type `<ext/>` are defined.
  The extension list along with corresponding future extensions is used to
  extend the storage capabilities of the PresentationML framework. This
  allows for various new kinds of data to be stored natively within the
  framework.

  [Note: Using this extLst element allows the generating application to
  store whether this extension property has been modified. end note]

  &copy; ISO/IEC 29500: 2016

---------------------------------------------------------------------------------
## Working with the NotesSlide Class

As shown in the Open XML SDK code sample that follows, every instance of
the `NotesSlide` class is associated with an
instance of the `DocumentFormat.OpenXml.Packaging.NotesSlidePart` class, which represents a
notes slide part, one of the parts of a PresentationML presentation file
package, and a part that is required for each notes slide in a
presentation file. Each `NotesSlide` class
instance may also be associated with an instance of the `DocumentFormat.OpenXml.Presentation.NotesMaster` class, which in turn is
associated with a similarly named presentation part, represented by the
`DocumentFormat.OpenXml.Packaging.NotesMasterPart` class.

The `NotesSlide` class, which represents the
`<notes/>` element, is therefore also associated with a series of other
classes that represent the child elements of the `<notes/>` element.
Among these classes, as shown in the following code sample, are the
`CommonSlideData` class and the `ColorMapOverride` class. The `DocumentFormat.OpenXml.Presentation.ShapeTree` class and the `DocumentFormat.OpenXml.Presentation.Shape` classes are in turn associated with
the `CommonSlideData` class.

--------------------------------------------------------------------------------
## Open XML SDK Code Example

> In the following code snippets `P` represents the `DocumentFormat.OpenXml.Presentation` namespace and `D` represents the <ref:DocumentFormat.OpenXml.Drawing> namespace.

In the snippet below, a presentation is opened with `Presentation.Open` and the first `DocumentFormat.OpenXml.Packaging.SlidePart`
is retrieved or added if the presentation does not already have a `SlidePart`.

### [C#](#tab/cs-0)
```csharp
using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true) ?? throw new Exception("Presentation Document does not exist"))
{
    // Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideIntoPresentation helper method to insert a new slide.
    SlidePart slidePart = presentationDocument.PresentationPart?.SlideParts.FirstOrDefault() ?? InsertNewSlideNS.InsertNewSlide(presentationDocument, 1, "my new slide");

    // Add a new NoteSlidePart if one does not already exist
    NotesSlidePart notesSlidePart = slidePart.NotesSlidePart ?? slidePart.AddNewPart<NotesSlidePart>();
```
### [Visual Basic](#tab/vb-0)
```vb
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(pptxPath, True)
            If presentationDocument Is Nothing Then Throw New Exception("Presentation Document does not exist")

            ' Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideIntoPresentation helper method to insert a new slide.
            Dim slidePart As SlidePart = If(presentationDocument.PresentationPart?.SlideParts.FirstOrDefault(), InsertNewSlide.InsertNewSlide(presentationDocument, 1, "my new slide"))

            ' Add a new NoteSlidePart if one does not already exist
            Dim notesSlidePart As NotesSlidePart = If(slidePart.NotesSlidePart, slidePart.AddNewPart(Of NotesSlidePart)())
```
***

In this snippet the a `NoteSlide` is added to the `NoteSlidePart` if one does not already exist.
The `NotesSlide` class constructor creates instances of the `CommonSlideData` class.
The `CommonSlideData` class constructor creates an instance of the `DocumentFormat.OpenXml.Presentation.ShapeTree` class, whose constructor in turn
creates additional class instances: an instance of the `DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties` class, an
instance of the `DocumentFormat.OpenXml.Presentation.GroupShapeProperties` class, and an instance
of the `DocumentFormat.OpenXml.Presentation.Shape` class.

### [C#](#tab/cs-1)
```csharp
    // Add a NoteSlide to the NoteSlidePart if one does not already exist.
    notesSlidePart.NotesSlide ??= new P.NotesSlide(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new D.Transform2D(
                        new D.Offset() { X = 0, Y = 0 },
                        new D.Extents() { Cx = 0, Cy = 0 },
                        new D.ChildOffset() { X = 0, Y = 0 },
                        new D.ChildExtents() { Cx = 0, Cy = 0 })),
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Add a NoteSlide to the NoteSlidePart if one does not already exist.
            If notesSlidePart.NotesSlide Is Nothing Then
                notesSlidePart.NotesSlide = New P.NotesSlide(
                New P.CommonSlideData(
                    New P.ShapeTree(
                        New P.NonVisualGroupShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = 1, .Name = ""},
                            New P.NonVisualGroupShapeDrawingProperties(),
                            New P.ApplicationNonVisualDrawingProperties()),
                        New P.GroupShapeProperties(
                            New D.Transform2D(
                                New D.Offset() With {.X = 0, .Y = 0},
                                New D.Extents() With {.Cx = 0, .Cy = 0},
                                New D.ChildOffset() With {.X = 0, .Y = 0},
                                New D.ChildExtents() With {.Cx = 0, .Cy = 0})),
                        shape)))
```
***

The `Shape` constructor creates an instance of `DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties`, `DocumentFormat.OpenXml.Presentation.ShapeProperties`, and `DocumentFormat.OpenXml.Presentation.TextBody` classes along with their required child elements. The `TextBody`
contains the `DocumentFormat.OpenXml.Drawing.Paragraph`, that has a `DocumentFormat.OpenXml.Drawing.Run`, which contains the text of the note. The slide part is then added to the notes slide part.

### [C#](#tab/cs-2)
```csharp
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 3, Name = "test Placeholder 3" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties(
                            new P.PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 })),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph(
                            new D.Run(
                                new D.Text("This is a test note!"))))))));

    notesSlidePart.AddPart(slidePart);
```
### [Visual Basic](#tab/vb-2)
```vb
            Dim shape As New P.Shape(
                            New P.NonVisualShapeProperties(
                                New P.NonVisualDrawingProperties() With {.Id = 3, .Name = "test Placeholder 3"},
                                New P.NonVisualShapeDrawingProperties(),
                                New P.ApplicationNonVisualDrawingProperties(
                                    New P.PlaceholderShape() With {.Type = PlaceholderValues.Body, .Index = 1})),
                            New P.ShapeProperties(),
                            New P.TextBody(
                                New D.BodyProperties(),
                                New D.Paragraph(
                                    New D.Run(
                                        New D.Text("This is a test note!")))))
```
***

The notes slide part created with the code above contains the following XML

```xml
<?xml version="1.0" encoding="utf-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="test Placeholder 3"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:r>
              <a:t>This is a test note!</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
```

The following snippets add the required `DocumentFormat.OpenXml.Packaging.NotesMasterPart` and `DocumentFormat.OpenXml.Packaging.ThemePart` if they are missing.

### [C#](#tab/cs-3)
```csharp
    // Add the required NotesMasterPart if it is missing
    NotesMasterPart notesMasterPart = notesSlidePart.NotesMasterPart ?? notesSlidePart.AddNewPart<NotesMasterPart>();

    // Add a NotesMaster to the NotesMasterPart if not present
    notesMasterPart.NotesMaster ??= new NotesMaster(
    new P.CommonSlideData(
        new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1, Name = "New Placeholder" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties())),
    new P.ColorMap()
    {
        Background1 = D.ColorSchemeIndexValues.Light1,
        Background2 = D.ColorSchemeIndexValues.Light2,
        Text1 = D.ColorSchemeIndexValues.Dark1,
        Text2 = D.ColorSchemeIndexValues.Dark2,
        Accent1 = D.ColorSchemeIndexValues.Accent1,
        Accent2 = D.ColorSchemeIndexValues.Accent2,
        Accent3 = D.ColorSchemeIndexValues.Accent3,
        Accent4 = D.ColorSchemeIndexValues.Accent4,
        Accent5 = D.ColorSchemeIndexValues.Accent5,
        Accent6 = D.ColorSchemeIndexValues.Accent6,
        Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink,
    });

    // Add a new ThemePart for the NotesMasterPart
    ThemePart themePart = notesMasterPart.ThemePart ?? notesMasterPart.AddNewPart<ThemePart>();

    // Add the Theme if it is missing
    themePart.Theme ??= new Theme(
        new ThemeElements(
            new ColorScheme(
                new Dark1Color(
                    new SystemColor() { Val = SystemColorValues.WindowText }),
                new Light1Color(
                    new SystemColor() { Val = SystemColorValues.Window }),
                new Dark2Color(
                    new RgbColorModelHex() { Val = "f1d7be" }),
                new Light2Color(
                    new RgbColorModelHex() { Val = "171717" }),
                new Accent1Color(
                    new RgbColorModelHex() { Val = "ea9f7d" }),
                new Accent2Color(
                    new RgbColorModelHex() { Val = "168ecd" }),
                new Accent3Color(
                    new RgbColorModelHex() { Val = "e694db" }),
                new Accent4Color(
                    new RgbColorModelHex() { Val = "f0612a" }),
                new Accent5Color(
                    new RgbColorModelHex() { Val = "5fd46c" }),
                new Accent6Color(
                    new RgbColorModelHex() { Val = "b158d1" }),
                new D.Hyperlink(
                    new RgbColorModelHex() { Val = "699f82" }),
                new FollowedHyperlinkColor(
                    new RgbColorModelHex() { Val = "699f82" }))
            { Name = "Office2" },
            new D.FontScheme(
                new MajorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()),
                new MinorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()))
            { Name = "Office2" },
            new FormatScheme(
                new FillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()),
                new LineStyleList(
                    new D.Outline(),
                    new D.Outline(),
                    new D.Outline()),
                new EffectStyleList(
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList())),
                new BackgroundFillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()))
            { Name = "Office2" }),
        new ObjectDefaults(),
        new ExtraColorSchemeList());
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Add the required NotesMasterPart if it is missing
            Dim notesMasterPart As NotesMasterPart = If(notesSlidePart.NotesMasterPart, notesSlidePart.AddNewPart(Of NotesMasterPart)())

            ' Add a NotesMaster to the NotesMasterPart if not present
            If notesMasterPart.NotesMaster Is Nothing Then
                notesMasterPart.NotesMaster = New NotesMaster(
                New P.CommonSlideData(
                    New P.ShapeTree(
                        New P.NonVisualGroupShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = 1, .Name = "New Placeholder"},
                            New P.NonVisualGroupShapeDrawingProperties(),
                            New P.ApplicationNonVisualDrawingProperties()),
                        New P.GroupShapeProperties())),
                New P.ColorMap() With {
                    .Background1 = D.ColorSchemeIndexValues.Light1,
                    .Background2 = D.ColorSchemeIndexValues.Light2,
                    .Text1 = D.ColorSchemeIndexValues.Dark1,
                    .Text2 = D.ColorSchemeIndexValues.Dark2,
                    .Accent1 = D.ColorSchemeIndexValues.Accent1,
                    .Accent2 = D.ColorSchemeIndexValues.Accent2,
                    .Accent3 = D.ColorSchemeIndexValues.Accent3,
                    .Accent4 = D.ColorSchemeIndexValues.Accent4,
                    .Accent5 = D.ColorSchemeIndexValues.Accent5,
                    .Accent6 = D.ColorSchemeIndexValues.Accent6,
                    .Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                })
            End If

            ' Add a new ThemePart for the NotesMasterPart
            Dim themePart As ThemePart = If(notesMasterPart.ThemePart, notesMasterPart.AddNewPart(Of ThemePart)())

            ' Add the Theme if it is missing
            If themePart.Theme Is Nothing Then
                themePart.Theme = New Theme(
                New ThemeElements(
                    New ColorScheme(
                        New Dark1Color(
                            New SystemColor() With {.Val = SystemColorValues.WindowText}),
                        New Light1Color(
                            New SystemColor() With {.Val = SystemColorValues.Window}),
                        New Dark2Color(
                            New RgbColorModelHex() With {.Val = "f1d7be"}),
                        New Light2Color(
                            New RgbColorModelHex() With {.Val = "171717"}),
                        New Accent1Color(
                            New RgbColorModelHex() With {.Val = "ea9f7d"}),
                        New Accent2Color(
                            New RgbColorModelHex() With {.Val = "168ecd"}),
                        New Accent3Color(
                            New RgbColorModelHex() With {.Val = "e694db"}),
                        New Accent4Color(
                            New RgbColorModelHex() With {.Val = "f0612a"}),
                        New Accent5Color(
                            New RgbColorModelHex() With {.Val = "5fd46c"}),
                        New Accent6Color(
                            New RgbColorModelHex() With {.Val = "b158d1"}),
                        New D.Hyperlink(
                            New RgbColorModelHex() With {.Val = "699f82"}),
                        New FollowedHyperlinkColor(
                            New RgbColorModelHex() With {.Val = "699f82"})) With {.Name = "Office2"},
                    New D.FontScheme(
                        New MajorFont(
                            New LatinFont(),
                            New EastAsianFont(),
                            New ComplexScriptFont()),
                        New MinorFont(
                            New LatinFont(),
                            New EastAsianFont(),
                            New ComplexScriptFont())) With {.Name = "Office2"},
                    New FormatScheme(
                        New FillStyleList(
                            New NoFill(),
                            New SolidFill(),
                            New D.GradientFill(),
                            New D.BlipFill(),
                            New D.PatternFill(),
                            New GroupFill()),
                        New LineStyleList(
                            New D.Outline(),
                            New D.Outline(),
                            New D.Outline()),
                        New EffectStyleList(
                            New EffectStyle(
                                New EffectList()),
                            New EffectStyle(
                                New EffectList()),
                            New EffectStyle(
                                New EffectList())),
                        New BackgroundFillStyleList(
                            New NoFill(),
                            New SolidFill(),
                            New D.GradientFill(),
                            New D.BlipFill(),
                            New D.PatternFill(),
                            New GroupFill())) With {.Name = "Office2"}),
                New ObjectDefaults(),
                New ExtraColorSchemeList())
            End If
```
***

## Sample code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs-4)
```csharp
using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true) ?? throw new Exception("Presentation Document does not exist"))
{
    // Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideIntoPresentation helper method to insert a new slide.
    SlidePart slidePart = presentationDocument.PresentationPart?.SlideParts.FirstOrDefault() ?? InsertNewSlideNS.InsertNewSlide(presentationDocument, 1, "my new slide");

    // Add a new NoteSlidePart if one does not already exist
    NotesSlidePart notesSlidePart = slidePart.NotesSlidePart ?? slidePart.AddNewPart<NotesSlidePart>();
    // Add a NoteSlide to the NoteSlidePart if one does not already exist.
    notesSlidePart.NotesSlide ??= new P.NotesSlide(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new D.Transform2D(
                        new D.Offset() { X = 0, Y = 0 },
                        new D.Extents() { Cx = 0, Cy = 0 },
                        new D.ChildOffset() { X = 0, Y = 0 },
                        new D.ChildExtents() { Cx = 0, Cy = 0 })),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 3, Name = "test Placeholder 3" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties(
                            new P.PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 })),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph(
                            new D.Run(
                                new D.Text("This is a test note!"))))))));

    notesSlidePart.AddPart(slidePart);
    // Add the required NotesMasterPart if it is missing
    NotesMasterPart notesMasterPart = notesSlidePart.NotesMasterPart ?? notesSlidePart.AddNewPart<NotesMasterPart>();

    // Add a NotesMaster to the NotesMasterPart if not present
    notesMasterPart.NotesMaster ??= new NotesMaster(
    new P.CommonSlideData(
        new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1, Name = "New Placeholder" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties())),
    new P.ColorMap()
    {
        Background1 = D.ColorSchemeIndexValues.Light1,
        Background2 = D.ColorSchemeIndexValues.Light2,
        Text1 = D.ColorSchemeIndexValues.Dark1,
        Text2 = D.ColorSchemeIndexValues.Dark2,
        Accent1 = D.ColorSchemeIndexValues.Accent1,
        Accent2 = D.ColorSchemeIndexValues.Accent2,
        Accent3 = D.ColorSchemeIndexValues.Accent3,
        Accent4 = D.ColorSchemeIndexValues.Accent4,
        Accent5 = D.ColorSchemeIndexValues.Accent5,
        Accent6 = D.ColorSchemeIndexValues.Accent6,
        Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink,
    });

    // Add a new ThemePart for the NotesMasterPart
    ThemePart themePart = notesMasterPart.ThemePart ?? notesMasterPart.AddNewPart<ThemePart>();

    // Add the Theme if it is missing
    themePart.Theme ??= new Theme(
        new ThemeElements(
            new ColorScheme(
                new Dark1Color(
                    new SystemColor() { Val = SystemColorValues.WindowText }),
                new Light1Color(
                    new SystemColor() { Val = SystemColorValues.Window }),
                new Dark2Color(
                    new RgbColorModelHex() { Val = "f1d7be" }),
                new Light2Color(
                    new RgbColorModelHex() { Val = "171717" }),
                new Accent1Color(
                    new RgbColorModelHex() { Val = "ea9f7d" }),
                new Accent2Color(
                    new RgbColorModelHex() { Val = "168ecd" }),
                new Accent3Color(
                    new RgbColorModelHex() { Val = "e694db" }),
                new Accent4Color(
                    new RgbColorModelHex() { Val = "f0612a" }),
                new Accent5Color(
                    new RgbColorModelHex() { Val = "5fd46c" }),
                new Accent6Color(
                    new RgbColorModelHex() { Val = "b158d1" }),
                new D.Hyperlink(
                    new RgbColorModelHex() { Val = "699f82" }),
                new FollowedHyperlinkColor(
                    new RgbColorModelHex() { Val = "699f82" }))
            { Name = "Office2" },
            new D.FontScheme(
                new MajorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()),
                new MinorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()))
            { Name = "Office2" },
            new FormatScheme(
                new FillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()),
                new LineStyleList(
                    new D.Outline(),
                    new D.Outline(),
                    new D.Outline()),
                new EffectStyleList(
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList())),
                new BackgroundFillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()))
            { Name = "Office2" }),
        new ObjectDefaults(),
        new ExtraColorSchemeList());
}
```
### [Visual Basic](#tab/vb-4)
```vb
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(pptxPath, True)
            If presentationDocument Is Nothing Then Throw New Exception("Presentation Document does not exist")

            ' Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideIntoPresentation helper method to insert a new slide.
            Dim slidePart As SlidePart = If(presentationDocument.PresentationPart?.SlideParts.FirstOrDefault(), InsertNewSlide.InsertNewSlide(presentationDocument, 1, "my new slide"))

            ' Add a new NoteSlidePart if one does not already exist
            Dim notesSlidePart As NotesSlidePart = If(slidePart.NotesSlidePart, slidePart.AddNewPart(Of NotesSlidePart)())
            Dim shape As New P.Shape(
                            New P.NonVisualShapeProperties(
                                New P.NonVisualDrawingProperties() With {.Id = 3, .Name = "test Placeholder 3"},
                                New P.NonVisualShapeDrawingProperties(),
                                New P.ApplicationNonVisualDrawingProperties(
                                    New P.PlaceholderShape() With {.Type = PlaceholderValues.Body, .Index = 1})),
                            New P.ShapeProperties(),
                            New P.TextBody(
                                New D.BodyProperties(),
                                New D.Paragraph(
                                    New D.Run(
                                        New D.Text("This is a test note!")))))
            ' Add a NoteSlide to the NoteSlidePart if one does not already exist.
            If notesSlidePart.NotesSlide Is Nothing Then
                notesSlidePart.NotesSlide = New P.NotesSlide(
                New P.CommonSlideData(
                    New P.ShapeTree(
                        New P.NonVisualGroupShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = 1, .Name = ""},
                            New P.NonVisualGroupShapeDrawingProperties(),
                            New P.ApplicationNonVisualDrawingProperties()),
                        New P.GroupShapeProperties(
                            New D.Transform2D(
                                New D.Offset() With {.X = 0, .Y = 0},
                                New D.Extents() With {.Cx = 0, .Cy = 0},
                                New D.ChildOffset() With {.X = 0, .Y = 0},
                                New D.ChildExtents() With {.Cx = 0, .Cy = 0})),
                        shape)))
            End If

            notesSlidePart.AddPart(slidePart)
            ' Add the required NotesMasterPart if it is missing
            Dim notesMasterPart As NotesMasterPart = If(notesSlidePart.NotesMasterPart, notesSlidePart.AddNewPart(Of NotesMasterPart)())

            ' Add a NotesMaster to the NotesMasterPart if not present
            If notesMasterPart.NotesMaster Is Nothing Then
                notesMasterPart.NotesMaster = New NotesMaster(
                New P.CommonSlideData(
                    New P.ShapeTree(
                        New P.NonVisualGroupShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = 1, .Name = "New Placeholder"},
                            New P.NonVisualGroupShapeDrawingProperties(),
                            New P.ApplicationNonVisualDrawingProperties()),
                        New P.GroupShapeProperties())),
                New P.ColorMap() With {
                    .Background1 = D.ColorSchemeIndexValues.Light1,
                    .Background2 = D.ColorSchemeIndexValues.Light2,
                    .Text1 = D.ColorSchemeIndexValues.Dark1,
                    .Text2 = D.ColorSchemeIndexValues.Dark2,
                    .Accent1 = D.ColorSchemeIndexValues.Accent1,
                    .Accent2 = D.ColorSchemeIndexValues.Accent2,
                    .Accent3 = D.ColorSchemeIndexValues.Accent3,
                    .Accent4 = D.ColorSchemeIndexValues.Accent4,
                    .Accent5 = D.ColorSchemeIndexValues.Accent5,
                    .Accent6 = D.ColorSchemeIndexValues.Accent6,
                    .Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                })
            End If

            ' Add a new ThemePart for the NotesMasterPart
            Dim themePart As ThemePart = If(notesMasterPart.ThemePart, notesMasterPart.AddNewPart(Of ThemePart)())

            ' Add the Theme if it is missing
            If themePart.Theme Is Nothing Then
                themePart.Theme = New Theme(
                New ThemeElements(
                    New ColorScheme(
                        New Dark1Color(
                            New SystemColor() With {.Val = SystemColorValues.WindowText}),
                        New Light1Color(
                            New SystemColor() With {.Val = SystemColorValues.Window}),
                        New Dark2Color(
                            New RgbColorModelHex() With {.Val = "f1d7be"}),
                        New Light2Color(
                            New RgbColorModelHex() With {.Val = "171717"}),
                        New Accent1Color(
                            New RgbColorModelHex() With {.Val = "ea9f7d"}),
                        New Accent2Color(
                            New RgbColorModelHex() With {.Val = "168ecd"}),
                        New Accent3Color(
                            New RgbColorModelHex() With {.Val = "e694db"}),
                        New Accent4Color(
                            New RgbColorModelHex() With {.Val = "f0612a"}),
                        New Accent5Color(
                            New RgbColorModelHex() With {.Val = "5fd46c"}),
                        New Accent6Color(
                            New RgbColorModelHex() With {.Val = "b158d1"}),
                        New D.Hyperlink(
                            New RgbColorModelHex() With {.Val = "699f82"}),
                        New FollowedHyperlinkColor(
                            New RgbColorModelHex() With {.Val = "699f82"})) With {.Name = "Office2"},
                    New D.FontScheme(
                        New MajorFont(
                            New LatinFont(),
                            New EastAsianFont(),
                            New ComplexScriptFont()),
                        New MinorFont(
                            New LatinFont(),
                            New EastAsianFont(),
                            New ComplexScriptFont())) With {.Name = "Office2"},
                    New FormatScheme(
                        New FillStyleList(
                            New NoFill(),
                            New SolidFill(),
                            New D.GradientFill(),
                            New D.BlipFill(),
                            New D.PatternFill(),
                            New GroupFill()),
                        New LineStyleList(
                            New D.Outline(),
                            New D.Outline(),
                            New D.Outline()),
                        New EffectStyleList(
                            New EffectStyle(
                                New EffectList()),
                            New EffectStyle(
                                New EffectList()),
                            New EffectStyle(
                                New EffectList())),
                        New BackgroundFillStyleList(
                            New NoFill(),
                            New SolidFill(),
                            New D.GradientFill(),
                            New D.BlipFill(),
                            New D.PatternFill(),
                            New GroupFill())) With {.Name = "Office2"}),
                New ObjectDefaults(),
                New ExtraColorSchemeList())
            End If
        End Using
```
***

--------------------------------------------------------------------------------
## See also

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)

[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)
