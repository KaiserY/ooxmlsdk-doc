# Create a presentation document by providing a file name

This topic shows how to use the classes in the Open XML SDK to
create a presentation document programmatically.

--------------------------------------------------------------------------------

## Create a Presentation

A presentation file, like all files defined by the Open XML standard,
consists of a package file container. This is the file that users see in
their file explorer; it usually has a .pptx extension. The package file
is represented in the Open XML SDK by the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class. The
presentation document contains, among other parts, a presentation part.
The presentation part, represented in the Open XML SDK by the `DocumentFormat.OpenXml.Packaging.PresentationPart` class, contains the basic
*PresentationML* definition for the slide presentation. PresentationML
is the markup language used for creating presentations. Each package can
contain only one presentation part, and its root element must be `<presentation/>`.

The API calls used to create a new presentation document package are
relatively simple. The first step is to call the static `DocumentFormat.OpenXml.Packaging.PresentationDocument.Create`
method of the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class, as shown here
in the `CreatePresentation` procedure, which is the first part of the
complete code sample presented later in the article. The
`CreatePresentation` code calls the override of the `Create` method that takes as arguments the path to
the new document and the type of presentation document to be created.
The types of presentation documents available in that argument are
defined by a `DocumentFormat.OpenXml.PresentationDocumentType` enumerated value.

Next, the code calls `DocumentFormat.OpenXml.Packaging.PresentationDocument.AddPresentationPart`, which creates and
returns a `PresentationPart`. After the `PresentationPart` class instance is created, a new
root element for the presentation is added by setting the `DocumentFormat.OpenXml.Packaging.PresentationPart.Presentation` property equal to the instance of the `DocumentFormat.OpenXml.Presentation.Presentation` class returned from a call to
the `Presentation` class constructor.

In order to create a complete, useable, and valid presentation, the code
must also add a number of other parts to the presentation package. In
the example code, this is taken care of by a call to a utility function
named `CreatePresentationsParts`. That function then calls a number of
other utility functions that, taken together, create all the
presentation parts needed for a basic presentation, including slide,
slide layout, slide master, and theme parts.

### [C#](#tab/cs-1)
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
```

### [Visual Basic](#tab/vb-1)
```vb
    Sub CreatePresentation(filepath As String)
        ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
        Using presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
            Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
            presentationPart.Presentation = New Presentation()

            CreatePresentationParts(presentationPart)
        End Using
    End Sub
```
***

Using the Open XML SDK, you can create presentation structure and
content by using strongly-typed classes that correspond to
PresentationML elements. You can find these classes in the `DocumentFormat.OpenXml.Presentation`
namespace. The following table lists the names of the classes that
correspond to the presentation, slide, slide master, slide layout, and
theme elements. The class that corresponds to the theme element is
actually part of the `DocumentFormat.OpenXml.Drawing` namespace.
Themes are common to all Open XML markup languages.

| PresentationML Element | Open XML SDK Class |
|---|---|
| `<presentation/>` | `DocumentFormat.OpenXml.Presentation.Presentation` |
| `<sld/>` | `DocumentFormat.OpenXml.Presentation.Slide` |
| `<sldMaster/>` | `DocumentFormat.OpenXml.Presentation.SlideMaster` |
| `<sldLayout/>` | `DocumentFormat.OpenXml.Presentation.SlideLayout` |
| `<theme/>` | `DocumentFormat.OpenXml.Drawing.Theme` |

The PresentationML code that follows is the XML in the presentation part
(in the file presentation.xml) for a simple presentation that contains
two slides.

```xml
    <p:presentation xmlns:p="…" … >
      <p:sldMasterIdLst>
        <p:sldMasterId xmlns:rel="https://…/relationships" rel:id="rId1"/>
      </p:sldMasterIdLst>
      <p:notesMasterIdLst>
        <p:notesMasterId xmlns:rel="https://…/relationships" rel:id="rId4"/>
      </p:notesMasterIdLst>
      <p:handoutMasterIdLst>
        <p:handoutMasterId xmlns:rel="https://…/relationships" rel:id="rId5"/>
      </p:handoutMasterIdLst>
      <p:sldIdLst>
        <p:sldId id="267" xmlns:rel="https://…/relationships" rel:id="rId2"/>
        <p:sldId id="256" xmlns:rel="https://…/relationships" rel:id="rId3"/>
      </p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
      <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

--------------------------------------------------------------------------------

## Sample Code

Following is the complete sample C\# and VB code to create a
presentation, given a file path.

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
static SlidePart CreateSlidePart(PresentationPart presentationPart)
{
    SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
    slidePart1.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new BodyProperties(),
                            new ListStyle(),
                            new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
            new ColorMapOverride(new MasterColorMapping()));
    return slidePart1;
}
static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
{
    SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>();
    SlideLayout slideLayout = new SlideLayout(
    new CommonSlideData(new ShapeTree(
      new P.NonVisualGroupShapeProperties(
      new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
      new P.NonVisualGroupShapeDrawingProperties(),
      new ApplicationNonVisualDrawingProperties()),
      new GroupShapeProperties(new TransformGroup()),
      new P.Shape(
      new P.NonVisualShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
      new P.ShapeProperties(),
      new P.TextBody(
        new BodyProperties(),
        new ListStyle(),
        new Paragraph(new EndParagraphRunProperties()))))),
    new ColorMapOverride(new MasterColorMapping()));
    slideLayoutPart1.SlideLayout = slideLayout;
    return slideLayoutPart1;
}
static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
{
    SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
    SlideMaster slideMaster = new SlideMaster(
    new CommonSlideData(new ShapeTree(
      new P.NonVisualGroupShapeProperties(
      new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
      new P.NonVisualGroupShapeDrawingProperties(),
      new ApplicationNonVisualDrawingProperties()),
      new GroupShapeProperties(new TransformGroup()),
      new P.Shape(
      new P.NonVisualShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
      new P.ShapeProperties(),
      new P.TextBody(
        new BodyProperties(),
        new ListStyle(),
        new Paragraph())))),
    new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
    new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
    new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
    slideMasterPart1.SlideMaster = slideMaster;

    return slideMasterPart1;
}
static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
{
    ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
    D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

    D.ThemeElements themeElements1 = new D.ThemeElements(
    new D.ColorScheme(
      new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
      new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
      new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
      new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
      new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
      new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
      new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
      new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
      new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
      new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
      new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
      new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
    { Name = "Office" },
      new D.FontScheme(
      new D.MajorFont(
      new D.LatinFont() { Typeface = "Calibri" },
      new D.EastAsianFont() { Typeface = "" },
      new D.ComplexScriptFont() { Typeface = "" }),
      new D.MinorFont(
      new D.LatinFont() { Typeface = "Calibri" },
      new D.EastAsianFont() { Typeface = "" },
      new D.ComplexScriptFont() { Typeface = "" }))
      { Name = "Office" },
      new D.FormatScheme(
      new D.FillStyleList(
      new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
          new D.SaturationModulation() { Val = 300000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
         new D.SaturationModulation() { Val = 300000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 35000 },
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
         new D.SaturationModulation() { Val = 350000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 100000 }
        ),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
      new D.NoFill(),
      new D.PatternFill(),
      new D.GroupFill()),
      new D.LineStyleList(
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      },
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      },
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      }),
      new D.EffectStyleList(
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
      new D.BackgroundFillStyleList(
      new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 }),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 }),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
      { Name = "Office" });

    theme1.Append(themeElements1);
    theme1.Append(new D.ObjectDefaults());
    theme1.Append(new D.ExtraColorSchemeList());

    themePart1.Theme = theme1;
    return themePart1;

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
    Function CreateSlidePart(presentationPart As PresentationPart) As SlidePart
        Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
        Dim shape As New P.Shape(
                        New P.NonVisualShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = CType(2UI, UInt32Value), .Name = "Title 1"},
                            New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {.NoGrouping = True}),
                            New ApplicationNonVisualDrawingProperties(New PlaceholderShape())),
                        New P.ShapeProperties(),
                        New P.TextBody(
                            New BodyProperties(),
                            New ListStyle(),
                            New Paragraph(New EndParagraphRunProperties() With {.Language = "en-US"})))
        slidePart1.Slide = New Slide(
            New CommonSlideData(
                New ShapeTree(
                    New P.NonVisualGroupShapeProperties(
                        New P.NonVisualDrawingProperties() With {.Id = CType(1UI, UInt32Value), .Name = ""},
                        New P.NonVisualGroupShapeDrawingProperties(),
                        New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New TransformGroup(shape))),
            New ColorMapOverride(New MasterColorMapping())))
        Return slidePart1
    End Function
    Function CreateSlideLayoutPart(slidePart1 As SlidePart) As SlideLayoutPart
        Dim slideLayoutPart1 As SlideLayoutPart = slidePart1.AddNewPart(Of SlideLayoutPart)("rId1")
        Dim slideLayout As New SlideLayout(
            New CommonSlideData(New ShapeTree(
                New P.NonVisualGroupShapeProperties(
                    New P.NonVisualDrawingProperties() With {.Id = CType(1UI, UInt32Value), .Name = ""},
                    New P.NonVisualGroupShapeDrawingProperties(),
                    New ApplicationNonVisualDrawingProperties()),
                New GroupShapeProperties(New TransformGroup()),
                New P.Shape(
                    New P.NonVisualShapeProperties(
                        New P.NonVisualDrawingProperties() With {.Id = CType(2UI, UInt32Value), .Name = ""},
                        New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {.NoGrouping = True}),
                        New ApplicationNonVisualDrawingProperties(New PlaceholderShape())),
                    New P.ShapeProperties(),
                    New P.TextBody(
                        New BodyProperties(),
                        New ListStyle(),
                        New Paragraph(New EndParagraphRunProperties()))))),
            New ColorMapOverride(New MasterColorMapping()))
        slideLayoutPart1.SlideLayout = slideLayout
        Return slideLayoutPart1
    End Function
    Function CreateSlideMasterPart(slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
        Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
        Dim slideMaster As New SlideMaster(
            New CommonSlideData(New ShapeTree(
                New P.NonVisualGroupShapeProperties(
                    New P.NonVisualDrawingProperties() With {.Id = CType(1UI, UInt32Value), .Name = ""},
                    New P.NonVisualGroupShapeDrawingProperties(),
                    New ApplicationNonVisualDrawingProperties()),
                New GroupShapeProperties(New TransformGroup()),
                New P.Shape(
                    New P.NonVisualShapeProperties(
                        New P.NonVisualDrawingProperties() With {.Id = CType(2UI, UInt32Value), .Name = "Title Placeholder 1"},
                        New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {.NoGrouping = True}),
                        New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title})),
                    New P.ShapeProperties(),
                    New P.TextBody(
                        New BodyProperties(),
                        New ListStyle(),
                        New Paragraph())))),
            New P.ColorMap() With {.Background1 = D.ColorSchemeIndexValues.Light1, .Text1 = D.ColorSchemeIndexValues.Dark1, .Background2 = D.ColorSchemeIndexValues.Light2, .Text2 = D.ColorSchemeIndexValues.Dark2, .Accent1 = D.ColorSchemeIndexValues.Accent1, .Accent2 = D.ColorSchemeIndexValues.Accent2, .Accent3 = D.ColorSchemeIndexValues.Accent3, .Accent4 = D.ColorSchemeIndexValues.Accent4, .Accent5 = D.ColorSchemeIndexValues.Accent5, .Accent6 = D.ColorSchemeIndexValues.Accent6, .Hyperlink = D.ColorSchemeIndexValues.Hyperlink, .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink},
            New SlideLayoutIdList(New SlideLayoutId() With {.Id = CType(2147483649UI, UInt32Value), .RelationshipId = "rId1"}),
            New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
        slideMasterPart1.SlideMaster = slideMaster

        Return slideMasterPart1
    End Function
    Function CreateTheme(slideMasterPart1 As SlideMasterPart) As ThemePart
        Dim themePart1 As ThemePart = slideMasterPart1.AddNewPart(Of ThemePart)("rId5")
        Dim theme1 As New D.Theme() With {.Name = "Office Theme"}

        Dim themeElements1 As New D.ThemeElements(
            New D.ColorScheme(
                New D.Dark1Color(New D.SystemColor() With {.Val = D.SystemColorValues.WindowText, .LastColor = "000000"}),
                New D.Light1Color(New D.SystemColor() With {.Val = D.SystemColorValues.Window, .LastColor = "FFFFFF"}),
                New D.Dark2Color(New D.RgbColorModelHex() With {.Val = "1F497D"}),
                New D.Light2Color(New D.RgbColorModelHex() With {.Val = "EEECE1"}),
                New D.Accent1Color(New D.RgbColorModelHex() With {.Val = "4F81BD"}),
                New D.Accent2Color(New D.RgbColorModelHex() With {.Val = "C0504D"}),
                New D.Accent3Color(New D.RgbColorModelHex() With {.Val = "9BBB59"}),
                New D.Accent4Color(New D.RgbColorModelHex() With {.Val = "8064A2"}),
                New D.Accent5Color(New D.RgbColorModelHex() With {.Val = "4BACC6"}),
                New D.Accent6Color(New D.RgbColorModelHex() With {.Val = "F79646"}),
                New D.Hyperlink(New D.RgbColorModelHex() With {.Val = "0000FF"}),
                New D.FollowedHyperlinkColor(New D.RgbColorModelHex() With {.Val = "800080"})) With {.Name = "Office"},
            New D.FontScheme(
                New D.MajorFont(
                    New D.LatinFont() With {.Typeface = "Calibri"},
                    New D.EastAsianFont() With {.Typeface = ""},
                    New D.ComplexScriptFont() With {.Typeface = ""}),
                New D.MinorFont(
                    New D.LatinFont() With {.Typeface = "Calibri"},
                    New D.EastAsianFont() With {.Typeface = ""},
                    New D.ComplexScriptFont() With {.Typeface = ""})) With {.Name = "Office"},
            New D.FormatScheme(
                New D.FillStyleList(
                    New D.SolidFill(New D.SchemeColor() With {.Val = D.SchemeColorValues.PhColor}),
                    New D.GradientFill(
                        New D.GradientStopList(
                            New D.GradientStop(New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0},
                            New D.GradientStop(New D.SchemeColor(New D.Tint() With {.Val = 37000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 35000},
                            New D.GradientStop(New D.SchemeColor(New D.Tint() With {.Val = 15000}, New D.SaturationModulation() With {.Val = 350000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 100000}),
                        New D.LinearGradientFill() With {.Angle = 16200000, .Scaled = True}),
                    New D.NoFill(),
                    New D.PatternFill(),
                    New D.GroupFill()),
                New D.LineStyleList(
                    New D.Outline(
                        New D.SolidFill(
                            New D.SchemeColor(
                                New D.Shade() With {.Val = 95000},
                                New D.SaturationModulation() With {.Val = 105000}) With {.Val = D.SchemeColorValues.PhColor}),
                        New D.PresetDash() With {.Val = D.PresetLineDashValues.Solid}) With {
                            .Width = 9525,
                            .CapType = D.LineCapValues.Flat,
                            .CompoundLineType = D.CompoundLineValues.Single,
                            .Alignment = D.PenAlignmentValues.Center},
                    New D.Outline(
                        New D.SolidFill(
                            New D.SchemeColor(
                                New D.Shade() With {.Val = 95000},
                                New D.SaturationModulation() With {.Val = 105000}) With {.Val = D.SchemeColorValues.PhColor}),
                        New D.PresetDash() With {.Val = D.PresetLineDashValues.Solid}) With {
                            .Width = 9525,
                            .CapType = D.LineCapValues.Flat,
                            .CompoundLineType = D.CompoundLineValues.Single,
                            .Alignment = D.PenAlignmentValues.Center},
                    New D.Outline(
                        New D.SolidFill(
                            New D.SchemeColor(
                                New D.Shade() With {.Val = 95000},
                                New D.SaturationModulation() With {.Val = 105000}) With {.Val = D.SchemeColorValues.PhColor}),
                        New D.PresetDash() With {.Val = D.PresetLineDashValues.Solid}) With {
                            .Width = 9525,
                            .CapType = D.LineCapValues.Flat,
                            .CompoundLineType = D.CompoundLineValues.Single,
                            .Alignment = D.PenAlignmentValues.Center}),
                New D.EffectStyleList(
                    New D.EffectStyle(
                        New D.EffectList(
                            New D.OuterShadow(
                                New D.RgbColorModelHex(
                                    New D.Alpha() With {.Val = 38000}) With {.Val = "000000"}) With {.BlurRadius = 40000L, .Distance = 20000L, .Direction = 5400000, .RotateWithShape = False})),
                    New D.EffectStyle(
                        New D.EffectList(
                            New D.OuterShadow(
                                New D.RgbColorModelHex(
                                    New D.Alpha() With {.Val = 38000}) With {.Val = "000000"}) With {.BlurRadius = 40000L, .Distance = 20000L, .Direction = 5400000, .RotateWithShape = False})),
                    New D.EffectStyle(
                        New D.EffectList(
                            New D.OuterShadow(
                                New D.RgbColorModelHex(
                                    New D.Alpha() With {.Val = 38000}) With {.Val = "000000"}) With {.BlurRadius = 40000L, .Distance = 20000L, .Direction = 5400000, .RotateWithShape = False}))),
                New D.BackgroundFillStyleList(
                    New D.SolidFill(New D.SchemeColor() With {.Val = D.SchemeColorValues.PhColor}),
                    New D.GradientFill(
                        New D.GradientStopList(
                            New D.GradientStop(
                                New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0},
                            New D.GradientStop(
                                New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0},
                            New D.GradientStop(
                                New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0}),
                        New D.LinearGradientFill() With {.Angle = 16200000, .Scaled = True}),
                    New D.GradientFill(
                        New D.GradientStopList(
                            New D.GradientStop(
                                New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0},
                            New D.GradientStop(
                                New D.SchemeColor(New D.Tint() With {.Val = 50000}, New D.SaturationModulation() With {.Val = 300000}) With {.Val = D.SchemeColorValues.PhColor}) With {.Position = 0}),
                        New D.LinearGradientFill() With {.Angle = 16200000, .Scaled = True}))) With {.Name = "Office"})

        theme1.Append(themeElements1)
        theme1.Append(New D.ObjectDefaults())
        theme1.Append(New D.ExtraColorSchemeList())

        themePart1.Theme = theme1
        Return themePart1
    End Function
```
***

--------------------------------------------------------------------------------

## See also 

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[Structure of a PresentationML Document](structure-of-a-presentationml-document.md)  

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)  

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
