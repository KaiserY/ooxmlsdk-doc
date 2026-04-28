# Insert a picture into a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add a picture to a word processing document.

--------------------------------------------------------------------------------

## Opening an Existing Document for Editing

To open an existing document, instantiate the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument`
class as shown in the following `using` statement. In the same
statement, open the word processing file at the specified `filepath`
by using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)`
method, with the Boolean parameter set to `true` in order to
enable editing the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
```
### [Visual Basic](#tab/vb-0)
```vb
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(document, True)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
It ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement. Because the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class in the Open XML SDK
automatically saves and closes the object as part of its `System.IDisposable` implementation, and because
`System.IDisposable.Dispose` is automatically called when you
exit the block, you do not have to explicitly call `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save` or
`DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Dispose` as long as you use a `using` statement.

--------------------------------------------------------------------------------
## The XML Representation of the Graphic Object
The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the Graphic Object Data element.

> This element specifies the reference to a graphic object within the
> document. This graphic object is provided entirely by the document
> authors who choose to persist this data within the document.
> 
> [*Note*: Depending on the type of graphical object used not every
> generating application that supports the OOXML framework will have the
> ability to render the graphical object. *end note*]
> 
> © ISO/IEC 29500: 2016

The following XML Schema fragment defines the contents of this element

```xml
    <complexType name="CT_GraphicalObjectData">
       <sequence>
           <any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
       </sequence>
       <attribute name="uri" type="xsd:token"/>
    </complexType>
```

--------------------------------------------------------------------------------

## How the Sample Code Works

After you have opened the document, add the `DocumentFormat.OpenXml.Packaging.ImagePart`
object to the `DocumentFormat.OpenXml.Packaging.MainDocumentPart` object by using a file
stream as shown in the following code segment.

### [C#](#tab/cs-1)
```csharp
        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
```
### [Visual Basic](#tab/vb-1)
```vb
            Dim mainPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            Dim imagePart As ImagePart = mainPart.AddImagePart(ImagePartType.Jpeg)

            Using stream As New FileStream(fileName, FileMode.Open)
                imagePart.FeedData(stream)
            End Using

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart))
```
***

To add the image to the body, first define the reference of the image.
Then, append the reference to the body. The element should be in a `DocumentFormat.OpenXml.Wordprocessing.Run`.

### [C#](#tab/cs-2)
```csharp
    // Define the reference of the image.
    var element =
         new Drawing(
             new DW.Inline(
                 new DW.Extent() { Cx = 990000L, Cy = 792000L },
                 new DW.EffectExtent()
                 {
                     LeftEdge = 0L,
                     TopEdge = 0L,
                     RightEdge = 0L,
                     BottomEdge = 0L
                 },
                 new DW.DocProperties()
                 {
                     Id = (UInt32Value)1U,
                     Name = "Picture 1"
                 },
                 new DW.NonVisualGraphicFrameDrawingProperties(
                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
                 new A.Graphic(
                     new A.GraphicData(
                         new PIC.Picture(
                             new PIC.NonVisualPictureProperties(
                                 new PIC.NonVisualDrawingProperties()
                                 {
                                     Id = (UInt32Value)0U,
                                     Name = "New Bitmap Image.jpg"
                                 },
                                 new PIC.NonVisualPictureDrawingProperties()),
                             new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                                "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         })
                                 )
                                 {
                                     Embed = relationshipId,
                                     CompressionState =
                                     A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(
                                     new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                 new A.PresetGeometry(
                                     new A.AdjustValueList()
                                 )
                                 { Preset = A.ShapeTypeValues.Rectangle }))
                     )
                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             )
             {
                 DistanceFromTop = (UInt32Value)0U,
                 DistanceFromBottom = (UInt32Value)0U,
                 DistanceFromLeft = (UInt32Value)0U,
                 DistanceFromRight = (UInt32Value)0U,
                 EditId = "50D07946"
             });

    if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.Document.Body is null)
    {
        throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
    }

    // Append the reference to body, the element should be in a Run.
    wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
```
### [Visual Basic](#tab/vb-2)
```vb
        ' Define the reference of the image.
        Dim element = New Drawing(
                              New DW.Inline(
                          New DW.Extent() With {.Cx = 990000L, .Cy = 792000L},
                          New DW.EffectExtent() With {.LeftEdge = 0L, .TopEdge = 0L, .RightEdge = 0L, .BottomEdge = 0L},
                          New DW.DocProperties() With {.Id = CType(1UI, UInt32Value), .Name = "Picture1"},
                          New DW.NonVisualGraphicFrameDrawingProperties(
                              New A.GraphicFrameLocks() With {.NoChangeAspect = True}
                              ),
                          New A.Graphic(New A.GraphicData(
                                        New PIC.Picture(
                                            New PIC.NonVisualPictureProperties(
                                                New PIC.NonVisualDrawingProperties() With {.Id = 0UI, .Name = "Koala.jpg"},
                                                New PIC.NonVisualPictureDrawingProperties()
                                                ),
                                            New PIC.BlipFill(
                                                New A.Blip(
                                                    New A.BlipExtensionList(
                                                        New A.BlipExtension() With {.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"})
                                                    ) With {.Embed = relationshipId, .CompressionState = A.BlipCompressionValues.Print},
                                                New A.Stretch(
                                                    New A.FillRectangle()
                                                    )
                                                ),
                                            New PIC.ShapeProperties(
                                                New A.Transform2D(
                                                    New A.Offset() With {.X = 0L, .Y = 0L},
                                                    New A.Extents() With {.Cx = 990000L, .Cy = 792000L}),
                                                New A.PresetGeometry(
                                                    New A.AdjustValueList()
                                                    ) With {.Preset = A.ShapeTypeValues.Rectangle}
                                                )
                                            )
                                        ) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"}
                                    )
                                ) With {.DistanceFromTop = 0UI,
                                        .DistanceFromBottom = 0UI,
                                        .DistanceFromLeft = 0UI,
                                        .DistanceFromRight = 0UI}
                            )

        ' Append the reference to body, the element should be in a Run.
        wordDoc.MainDocumentPart.Document.Body.AppendChild(New Paragraph(New Run(element)))
```
***

--------------------------------------------------------------------------------

## Sample Code
The following code example adds a picture to an existing word document.
In your code, you can call the `InsertAPicture` method by passing in the path of
the word document, and the path of the file that contains the picture.
For example, the following call inserts the picture.

### [C#](#tab/cs-3)
```csharp
string documentPath = args[0];
string picturePath = args[1];

InsertAPicture(documentPath, picturePath);
```
### [Visual Basic](#tab/vb-3)
```vb
        Dim documentPath As String = args(0)
        Dim picturePath As String = args(1)

        InsertAPicture(documentPath, picturePath)
```
***

After you run the code, look at the file to see the inserted picture.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

static void InsertAPicture(string document, string fileName)
{
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
```

### [Visual Basic](#tab/vb)
```vb
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports A = DocumentFormat.OpenXml.Drawing
Imports DW = DocumentFormat.OpenXml.Drawing.Wordprocessing
Imports PIC = DocumentFormat.OpenXml.Drawing.Pictures

Module Program
    Sub Main(args As String())
        Dim documentPath As String = args(0)
        Dim picturePath As String = args(1)

        InsertAPicture(documentPath, picturePath)
```

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
