# Add a video to a slide in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to add a video to the first slide in a presentation
programmatically.

## Getting a Presentation Object 

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open` method that uses a file path, and a
Boolean value as the second parameter to specify whether a document is
editable. To open a document for read/write, specify the value `true` for this parameter as shown in the following
`using` statement. In this code, the file
parameter is a string that represents the path for the file from which
you want to open the document.

### [C#](#tab/cs-1)
```csharp
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(filePath, True)
```
***

With v3.0.0+ the `DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close` method
has been removed in favor of relying on the [using statement](https://learn.microsoft.com/dotnet/csharp/language-reference/statements/using).
This ensures that the `System.IDisposable.Dispose` method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `ppt`.

## The Structure of the Video From File

The PresentationML document consists of a number of parts, among which is the Picture (`<pic/>`) element.

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification introduces the overall form of a `PresentationML` package.

Video File (`<videoFile/>`) specifies the presence of a video file. It is defined within the non-visual properties of an object. The video shall be attached to an object as this is how it is represented within the document. The actual playing of the video however is done within the timing node list that is specified under the timing element.

Consider the following `Picture` object that has a video attached to it.

```xml
<p:pic>  
  <p:nvPicPr>  
    <p:cNvPr id="7" name="Rectangle 6">  
      <a:hlinkClick r:id="" action="ppaction://media"/>  
    </p:cNvPr>  
    <p:cNvPicPr>  
      <a:picLocks noRot="1"/>  
    </p:cNvPicPr>  
    <p:nvPr>  
      <a:videoFile r:link="rId1"/>  
    </p:nvPr>  
  </p:nvPicPr>  
</p:pic>
```

In the above example, we see that there is a single videoFile element attached to this picture. This picture is placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this case, is used to refer to this videoFile element from within the timing node list. The Linked relationship id is used to retrieve the actual video file for playback purposes. 

&copy; ISO/IEC 29500: 2016

The following XML Schema fragment defines the contents of videoFile.

```xml
<xsd:complexType name="CT_TLMediaNodeVideo">
	<xsd:sequence>
		<xsd:element name="cMediaNode" type="CT_TLCommonMediaNodeData" minOccurs="1" maxOccurs="1"/>
	</xsd:sequence>
	<xsd:attribute name="fullScrn" type="xsd:boolean" use="optional" default="false"/>
</xsd:complexType>
```

## How the Sample Code Works

After opening the presentation file for read/write access in the `using` statement, the code gets the presentation
part from the presentation document. Then it gets the relationship ID of
the last slide, and gets the slide part from the relationship ID.

### [C#](#tab/cs-2)
```csharp
        //Get presentation part
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        //Get slides ids.
        OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;

        //Get relationsipId of the last slide
        string? videoSldRelationshipId = ((SlideId) slidesIds[slidesIds.ToArray().Length - 1]).RelationshipId;

        if (videoSldRelationshipId == null)
        {
            throw new NullReferenceException("Slide id not found");
        }

        //Get slide part by relationshipID
        SlidePart? slidePart = (SlidePart) presentationPart.GetPartById(videoSldRelationshipId);
```

### [Visual Basic](#tab/vb-2)
```vb
            ' Get presentation part
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Get slides ids
            Dim slidesIds As OpenXmlElementList = presentationPart.Presentation.SlideIdList.ChildElements

            ' Get relationshipId of the last slide
            Dim videoSldRelationshipId As String = CType(slidesIds(slidesIds.ToArray().Length - 1), SlideId).RelationshipId

            If videoSldRelationshipId Is Nothing Then
                Throw New NullReferenceException("Slide id not found")
            End If

            ' Get slide part by relationshipID
            Dim slidePart As SlidePart = CType(presentationPart.GetPartById(videoSldRelationshipId), SlidePart)
```
***

The code first creates a media data part for the video file to be added. With the video file stream open, it feeds the media data part object. Next, video and media relationship references are added to the slide using the provided embedId for future reference to the video file and mediaEmbedId for media reference.

An image part is then added with a sample picture to be used as a placeholder for the video. A picture object is created with various elements, such as Non-Visual Drawing Properties (`<cNvPr/>`), which specify non-visual canvas properties. This allows for additional information that does not affect the appearance of the picture to be stored. The `<videoFile/>` element, explained above, is also included. The HyperLinkOnClick (`<hlinkClick/>`) element specifies the on-click hyperlink information to be applied to a run of text or image. When the hyperlink text or image is clicked, the link is fetched. Non-Visual Picture Drawing Properties (`<cNvPicPr/>`) specify the non-visual properties for the picture canvas. For a detailed explanation of the elements used, please refer to [ISO/IEC 29500](https://www.iso.org/standard/71691.html)

### [C#](#tab/cs-3)
```csharp
        // Create video Media Data Part (content type, extension)
        MediaDataPart mediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

        //Get the video file and feed the stream
        using (Stream mediaDataPartStream = File.OpenRead(videoFilePath))
        {
            mediaDataPart.FeedData(mediaDataPartStream);
        }
        //Adds a VideoReferenceRelationship to the MainDocumentPart
        slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId);

        //Adds a MediaReferenceRelationship to the SlideLayoutPart
        slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId);

        NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = shapeId, Name = "video" };
        A.VideoFromFile videoFromFile = new A.VideoFromFile() { Link = embedId };

        ApplicationNonVisualDrawingProperties appNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
        appNonVisualDrawingProperties.Append(videoFromFile);
       
        //adds sample image to the slide with id to be used as reference in blip
        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId);
        using (Stream data = File.OpenRead(coverPicPath))
        {
            imagePart.FeedData(data);
        }
       
        if (slidePart!.Slide!.CommonSlideData!.ShapeTree == null)
        {
            throw new NullReferenceException("Presentation shape tree is empty");
        }

        //Getting existing shape tree object from PowerPoint
        ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

        // specifies the existence of a picture within a presentation.
        // It can have non-visual properties, a picture fill as well as shape properties attached to it.
        Picture picture = new Picture();
        NonVisualPictureProperties nonVisualPictureProperties = new NonVisualPictureProperties();

        A.HyperlinkOnClick hyperlinkOnClick = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };
        nonVisualDrawingProperties.Append(hyperlinkOnClick);

        NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
        A.PictureLocks pictureLocks = new A.PictureLocks() { NoChangeAspect = true };
        nonVisualPictureDrawingProperties.Append(pictureLocks);

        ApplicationNonVisualDrawingPropertiesExtensionList appNonVisualDrawingPropertiesExtensionList = new ApplicationNonVisualDrawingPropertiesExtensionList();
        ApplicationNonVisualDrawingPropertiesExtension appNonVisualDrawingPropertiesExtension = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
```

### [Visual Basic](#tab/vb-3)
```vb
            ' Create video Media Data Part (content type, extension)
            Dim mediaDataPart As MediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4")

            ' Get the video file and feed the stream
            Using mediaDataPartStream As Stream = File.OpenRead(videoFilePath)
                mediaDataPart.FeedData(mediaDataPartStream)
            End Using

            ' Adds a VideoReferenceRelationship to the MainDocumentPart
            slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId)

            ' Adds a MediaReferenceRelationship to the SlideLayoutPart
            slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId)

            Dim nonVisualDrawingProperties As New NonVisualDrawingProperties() With {
                .Id = shapeId,
                .Name = "video"
            }
            Dim videoFromFile As New A.VideoFromFile() With {
                .Link = embedId
            }

            Dim appNonVisualDrawingProperties As New ApplicationNonVisualDrawingProperties()
            appNonVisualDrawingProperties.Append(videoFromFile)

            ' Adds sample image to the slide with id to be used as reference in blip
            Dim imagePart As ImagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId)
            Using data As Stream = File.OpenRead(coverPicPath)
                imagePart.FeedData(data)
            End Using

            If slidePart.Slide.CommonSlideData.ShapeTree Is Nothing Then
                Throw New NullReferenceException("Presentation shape tree is empty")
            End If

            ' Getting existing shape tree object from PowerPoint
            Dim shapeTree As ShapeTree = slidePart.Slide.CommonSlideData.ShapeTree

            ' Specifies the existence of a picture within a presentation
            Dim picture As New Picture()
            Dim nonVisualPictureProperties As New NonVisualPictureProperties()

            Dim hyperlinkOnClick As New A.HyperlinkOnClick() With {
                .Id = "",
                .Action = "ppaction://media"
            }
            nonVisualDrawingProperties.Append(hyperlinkOnClick)

            Dim nonVisualPictureDrawingProperties As New NonVisualPictureDrawingProperties()
            Dim pictureLocks As New A.PictureLocks() With {
                .NoChangeAspect = True
            }
            nonVisualPictureDrawingProperties.Append(pictureLocks)

            Dim appNonVisualDrawingPropertiesExtensionList As New ApplicationNonVisualDrawingPropertiesExtensionList()
            Dim appNonVisualDrawingPropertiesExtension As New ApplicationNonVisualDrawingPropertiesExtension() With {
                .Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
            }
```
***

Next Media(CT_Media) element is created with use of previously referenced mediaEmbedId(Embedded Picture Reference). The Blip element is also added; this element specifies the existence of an image (binary large image or picture) and contains a reference to the image data. Blip's Embed attribute is used to specify a placeholder image in the Image Part created previously.

### [C#](#tab/cs-4)
```csharp
        P14.Media media = new() { Embed = mediaEmbedId };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        appNonVisualDrawingPropertiesExtension.Append(media);
        appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension);
        appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList);

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        //Prepare shape properties to display picture
        BlipFill blipFill = new BlipFill();
        A.Blip blip = new A.Blip() { Embed = imgEmbedId };
```

### [Visual Basic](#tab/vb-4)
```vb
            Dim media As New P14.Media() With {
                .Embed = mediaEmbedId
            }
            media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main")

            appNonVisualDrawingPropertiesExtension.Append(media)
            appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension)
            appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList)

            nonVisualPictureProperties.Append(nonVisualDrawingProperties)
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties)
            nonVisualPictureProperties.Append(appNonVisualDrawingProperties)

            ' Prepare shape properties to display picture
            Dim blipFill As New BlipFill()
            Dim blip As New A.Blip() With {
                .Embed = imgEmbedId
            }
```
***

All other elements such Offset(`<off/>`), Stretch(`<stretch/>`), FillRectangle(`<fillRect/>`), are appended to the ShapeProperties(`<spPr/>`) and ShapeProperties are appended to the Picture element(`<pic/>`). Finally the picture element that incudes video is added to the ShapeTree(`<sp/>`) of the slide.

Following is the complete sample code that you can use to add video to the slide.

## Sample Code

### [C#](#tab/cs)
```csharp
AddVideo(args[0], args[1], args[2]);

static void AddVideo(string filePath, string videoFilePath, string coverPicPath)
{

    string imgEmbedId = "rId4", embedId = "rId3", mediaEmbedId = "rId2";
    UInt32Value shapeId = 5;
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
    {

        if (presentationDocument.PresentationPart == null || presentationDocument.PresentationPart.Presentation.SlideIdList == null)
        {
            throw new NullReferenceException("Presentation Part is empty or there are no slides in it");
        }
        //Get presentation part
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        //Get slides ids.
        OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;

        //Get relationsipId of the last slide
        string? videoSldRelationshipId = ((SlideId) slidesIds[slidesIds.ToArray().Length - 1]).RelationshipId;

        if (videoSldRelationshipId == null)
        {
            throw new NullReferenceException("Slide id not found");
        }

        //Get slide part by relationshipID
        SlidePart? slidePart = (SlidePart) presentationPart.GetPartById(videoSldRelationshipId);
        // Create video Media Data Part (content type, extension)
        MediaDataPart mediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

        //Get the video file and feed the stream
        using (Stream mediaDataPartStream = File.OpenRead(videoFilePath))
        {
            mediaDataPart.FeedData(mediaDataPartStream);
        }
        //Adds a VideoReferenceRelationship to the MainDocumentPart
        slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId);

        //Adds a MediaReferenceRelationship to the SlideLayoutPart
        slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId);

        NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = shapeId, Name = "video" };
        A.VideoFromFile videoFromFile = new A.VideoFromFile() { Link = embedId };

        ApplicationNonVisualDrawingProperties appNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
        appNonVisualDrawingProperties.Append(videoFromFile);
       
        //adds sample image to the slide with id to be used as reference in blip
        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId);
        using (Stream data = File.OpenRead(coverPicPath))
        {
            imagePart.FeedData(data);
        }
       
        if (slidePart!.Slide!.CommonSlideData!.ShapeTree == null)
        {
            throw new NullReferenceException("Presentation shape tree is empty");
        }

        //Getting existing shape tree object from PowerPoint
        ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

        // specifies the existence of a picture within a presentation.
        // It can have non-visual properties, a picture fill as well as shape properties attached to it.
        Picture picture = new Picture();
        NonVisualPictureProperties nonVisualPictureProperties = new NonVisualPictureProperties();

        A.HyperlinkOnClick hyperlinkOnClick = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };
        nonVisualDrawingProperties.Append(hyperlinkOnClick);

        NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
        A.PictureLocks pictureLocks = new A.PictureLocks() { NoChangeAspect = true };
        nonVisualPictureDrawingProperties.Append(pictureLocks);

        ApplicationNonVisualDrawingPropertiesExtensionList appNonVisualDrawingPropertiesExtensionList = new ApplicationNonVisualDrawingPropertiesExtensionList();
        ApplicationNonVisualDrawingPropertiesExtension appNonVisualDrawingPropertiesExtension = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
        P14.Media media = new() { Embed = mediaEmbedId };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        appNonVisualDrawingPropertiesExtension.Append(media);
        appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension);
        appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList);

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        //Prepare shape properties to display picture
        BlipFill blipFill = new BlipFill();
        A.Blip blip = new A.Blip() { Embed = imgEmbedId };
        A.Stretch stretch = new A.Stretch();
        A.FillRectangle fillRectangle = new A.FillRectangle();
        A.Transform2D transform2D = new A.Transform2D();
        A.Offset offset = new A.Offset() { X = 1524000L, Y = 857250L };
        A.Extents extents = new A.Extents() { Cx = 9144000L, Cy = 5143500L };
        A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjValueList = new A.AdjustValueList();

        stretch.Append(fillRectangle);
        blipFill.Append(blip);
        blipFill.Append(stretch);
        transform2D.Append(offset);
        transform2D.Append(extents);
        presetGeometry.Append(adjValueList);

        ShapeProperties shapeProperties = new ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        //adds all elements to the slide's shape tree
        picture.Append(nonVisualPictureProperties);
        picture.Append(blipFill);
        picture.Append(shapeProperties);

        shapeTree.Append(picture);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
Module Program
    Sub Main(args As String())
        AddVideo(args(0), args(1), args(2))
    End Sub

    Sub AddVideo(filePath As String, videoFilePath As String, coverPicPath As String)
        Dim imgEmbedId As String = "rId4"
        Dim embedId As String = "rId3"
        Dim mediaEmbedId As String = "rId2"
        Dim shapeId As UInt32Value = 5
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(filePath, True)
            If presentationDocument.PresentationPart Is Nothing OrElse presentationDocument.PresentationPart.Presentation.SlideIdList Is Nothing Then
                Throw New NullReferenceException("Presentation Part is empty or there are no slides in it")
            End If
            ' Get presentation part
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Get slides ids
            Dim slidesIds As OpenXmlElementList = presentationPart.Presentation.SlideIdList.ChildElements

            ' Get relationshipId of the last slide
            Dim videoSldRelationshipId As String = CType(slidesIds(slidesIds.ToArray().Length - 1), SlideId).RelationshipId

            If videoSldRelationshipId Is Nothing Then
                Throw New NullReferenceException("Slide id not found")
            End If

            ' Get slide part by relationshipID
            Dim slidePart As SlidePart = CType(presentationPart.GetPartById(videoSldRelationshipId), SlidePart)
            ' Create video Media Data Part (content type, extension)
            Dim mediaDataPart As MediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4")

            ' Get the video file and feed the stream
            Using mediaDataPartStream As Stream = File.OpenRead(videoFilePath)
                mediaDataPart.FeedData(mediaDataPartStream)
            End Using

            ' Adds a VideoReferenceRelationship to the MainDocumentPart
            slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId)

            ' Adds a MediaReferenceRelationship to the SlideLayoutPart
            slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId)

            Dim nonVisualDrawingProperties As New NonVisualDrawingProperties() With {
                .Id = shapeId,
                .Name = "video"
            }
            Dim videoFromFile As New A.VideoFromFile() With {
                .Link = embedId
            }

            Dim appNonVisualDrawingProperties As New ApplicationNonVisualDrawingProperties()
            appNonVisualDrawingProperties.Append(videoFromFile)

            ' Adds sample image to the slide with id to be used as reference in blip
            Dim imagePart As ImagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId)
            Using data As Stream = File.OpenRead(coverPicPath)
                imagePart.FeedData(data)
            End Using

            If slidePart.Slide.CommonSlideData.ShapeTree Is Nothing Then
                Throw New NullReferenceException("Presentation shape tree is empty")
            End If

            ' Getting existing shape tree object from PowerPoint
            Dim shapeTree As ShapeTree = slidePart.Slide.CommonSlideData.ShapeTree

            ' Specifies the existence of a picture within a presentation
            Dim picture As New Picture()
            Dim nonVisualPictureProperties As New NonVisualPictureProperties()

            Dim hyperlinkOnClick As New A.HyperlinkOnClick() With {
                .Id = "",
                .Action = "ppaction://media"
            }
            nonVisualDrawingProperties.Append(hyperlinkOnClick)

            Dim nonVisualPictureDrawingProperties As New NonVisualPictureDrawingProperties()
            Dim pictureLocks As New A.PictureLocks() With {
                .NoChangeAspect = True
            }
            nonVisualPictureDrawingProperties.Append(pictureLocks)

            Dim appNonVisualDrawingPropertiesExtensionList As New ApplicationNonVisualDrawingPropertiesExtensionList()
            Dim appNonVisualDrawingPropertiesExtension As New ApplicationNonVisualDrawingPropertiesExtension() With {
                .Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
            }
            Dim media As New P14.Media() With {
                .Embed = mediaEmbedId
            }
            media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main")

            appNonVisualDrawingPropertiesExtension.Append(media)
            appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension)
            appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList)

            nonVisualPictureProperties.Append(nonVisualDrawingProperties)
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties)
            nonVisualPictureProperties.Append(appNonVisualDrawingProperties)

            ' Prepare shape properties to display picture
            Dim blipFill As New BlipFill()
            Dim blip As New A.Blip() With {
                .Embed = imgEmbedId
            }
            Dim stretch As New A.Stretch()
            Dim fillRectangle As New A.FillRectangle()
            Dim transform2D As New A.Transform2D()
            Dim offset As New A.Offset() With {
                .X = 1524000L,
                .Y = 857250L
            }
            Dim extents As New A.Extents() With {
                .Cx = 9144000L,
                .Cy = 5143500L
            }
            Dim presetGeometry As New A.PresetGeometry() With {
                .Preset = A.ShapeTypeValues.Rectangle
            }
            Dim adjValueList As New A.AdjustValueList()

            stretch.Append(fillRectangle)
            blipFill.Append(blip)
            blipFill.Append(stretch)
            transform2D.Append(offset)
            transform2D.Append(extents)
            presetGeometry.Append(adjValueList)

            Dim shapeProperties As New ShapeProperties()
            shapeProperties.Append(transform2D)
            shapeProperties.Append(presetGeometry)

            ' Adds all elements to the slide's shape tree
            picture.Append(nonVisualPictureProperties)
            picture.Append(blipFill)
            picture.Append(shapeProperties)

            shapeTree.Append(picture)
        End Using
    End Sub
End Module
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
