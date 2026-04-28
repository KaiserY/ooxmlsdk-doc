# Open a presentation document for read-only access

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically open a presentation document for read-only
access.

## How to Open a File for Read-Only Access

You may want to open a presentation document to read the slides. You
might want to extract information from a slide, copy a slide to a slide
library, or list the titles of the slides. In such cases you want to do
so in a way that ensures the document remains unchanged. You can do that
by opening the document for read-only access. This How-To topic
discusses several ways to programmatically open a read-only presentation
document.

## Create an Instance of the PresentationDocument Class 

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call one
of the `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open` methods. Several Open methods are
provided, each with a different signature. The following table contains
a subset of the overloads for the `Open`
method that you can use to open the package.

| Name | Description |
|---|---|
| `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)` | Create a new instance of the `PresentationDocument` class from the specified file. |
| `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-io-stream-system-boolean)` | Create a new instance of the `PresentationDocument` class from the I/O stream. |
| `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-io-packaging-package)` | Create a new instance of the `PresentationDocument` class from the specified package. |

The previous table includes two `Open`
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read-only access,
specify the value `false` for this parameter.

For example, you can open the presentation file as read-only and assign
it to a `DocumentFormat.OpenXml.Packaging.PresentationDocument` object as shown in the
following `using` statement. In this code,
the `presentationFile` parameter is a string
that represents the path of the file from which you want to open the
document.

### [C#](#tab/cs-0)
```csharp
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFilePath, false))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFilePath, False)
        ' Insert other code here.
    End Using
```
***

You can also use the second overload of the `Open` method, in the table above, to create an
instance of the `PresentationDocument` class
based on an I/O stream. You might use this approach if you have a
Microsoft SharePoint Foundation 2010 application that uses stream I/O
and you want to use the Open XML SDK to work with a document. The
following code segment opens a document based on a stream.

### [C#](#tab/cs-1)
```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, false)) 
    {
        // Place other code here.
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(stream, False)
        ' Other code goes here.
    End Using
```
***

Suppose you have an application that employs the Open XML support in the
`System.IO.Packaging` namespace of the .NET
Framework Class Library, and you want to use the Open XML SDK to
work with a package read-only. The Open XML SDK includes a method
overload that accepts a `Package` as the only
parameter. There is no Boolean parameter to indicate whether the
document should be opened for editing. The recommended approach is to
open the package as read-only prior to creating the instance of the
`PresentationDocument` class. The following
code segment performs this operation.

### [C#](#tab/cs-2)
```csharp
    Package presentationPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationPackage))
    {
        // Other code goes here.
    }
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim presentationPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationPackage)
        ' Other code goes here.
    End Using
```
***

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

In the sample code, after you open the presentation document in the
`using` statement for read-only access,
instantiate the `PresentationPart`, and open
the slide list. Then you get the relationship ID of the first slide.

### [C#](#tab/cs-3)
```csharp
        // Get the relationship ID of the first slide.
        PresentationPart? part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part?.Presentation?.SlideIdList?.ChildElements ?? default;

        // If there are no slide IDs then there are no slides.
        if (slideIds.Count == 0)
        {
            sldText = "";
            return;
        }

        string? relId = (slideIds[index] as SlideId)?.RelationshipId;

        if (relId is null)
        {
            sldText = "";
            return;
        }
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = If(part?.Presentation?.SlideIdList?.ChildElements, New OpenXmlElementList())

            ' If there are no slide IDs then there are no slides.
            If slideIds.Count = 0 Then
                sldText = ""
                Return
            End If

            Dim relId As String = TryCast(slideIds(index), SlideId)?.RelationshipId

            If relId Is Nothing Then
                sldText = ""
                Return
            End If
```
***

From the relationship ID, `relId`, you get the
slide part, and then the inner text of the slide by building a text
string using `StringBuilder`.

### [C#](#tab/cs-4)
```csharp
        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part!.GetPartById(relId);

        // Build a StringBuilder object.
        StringBuilder paragraphText = new StringBuilder();

        // Get the inner text of the slide:
        IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
        foreach (A.Text text in texts)
        {
            paragraphText.Append(text.Text);
        }
        sldText = paragraphText.ToString();
```
### [Visual Basic](#tab/vb-4)
```vb
            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

            ' Build a StringBuilder object.
            Dim paragraphText As New StringBuilder()

            ' Get the inner text of the slide:
            Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
            For Each text As A.Text In texts
                paragraphText.Append(text.Text)
            Next
            sldText = paragraphText.ToString()
```
***

The inner text of the slide, which is an `out` parameter of the `GetSlideIdAndText` method, is passed back to the
main method to be displayed.

> **Important**
> This example displays only the text in the presentation file. Non-text parts, such as shapes or graphics, are not displayed.

## Sample Code

The following example opens a presentation file for read-only access and
gets the inner text of a slide at a specified index. To call the method `GetSlideIdAndText` pass in the full path of the
presentation document. Also pass in the `out`
parameter `sldText`, which will be assigned a
value in the method itself, and then you can display its value in the
main program. For example, the following call to the `GetSlideIdAndText` method gets the inner text in a presentation file 
from the index and file path passed to the application as arguments.

> **Tip**
> The most expected exception in this program is the `ArgumentOutOfRangeException` exception. It could be thrown if, for example, you have a file with two slides, and you wanted to display the text in slide number 4. Therefore, it is best to use a `try` block when you call the `GetSlideIdAndText` method as shown in the following example.

### [C#](#tab/cs-5)
```csharp
try
{
    string file = args[0];
    bool isInt = int.TryParse(args[1], out int i);

    if (isInt)
    {
        GetSlideIdAndText(out string sldText, file, i);
        Console.WriteLine($"The text in slide #{i + 1} is {sldText}");
    }
}
catch(ArgumentOutOfRangeException exp) {
    Console.Error.WriteLine(exp.Message);
}
```
### [Visual Basic](#tab/vb-5)
```vb
        Try
            Dim file As String = args(0)
            Dim i As Integer
            Dim isInt As Boolean = Integer.TryParse(args(1), i)

            If isInt Then
                Dim sldText As String = String.Empty
                GetSlideIdAndText(sldText, file, i)
                Console.WriteLine($"The text in slide #{i + 1} is {sldText}")
            End If
        Catch exp As ArgumentOutOfRangeException
            Console.Error.WriteLine(exp.Message)
        End Try
```
***

The following is the complete code listing in C\# and Visual Basic.

### [C#](#tab/cs-6)
```csharp
static void GetSlideIdAndText(out string sldText, string docName, int index)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
    {
        // Get the relationship ID of the first slide.
        PresentationPart? part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part?.Presentation?.SlideIdList?.ChildElements ?? default;

        // If there are no slide IDs then there are no slides.
        if (slideIds.Count == 0)
        {
            sldText = "";
            return;
        }

        string? relId = (slideIds[index] as SlideId)?.RelationshipId;

        if (relId is null)
        {
            sldText = "";
            return;
        }
        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part!.GetPartById(relId);

        // Build a StringBuilder object.
        StringBuilder paragraphText = new StringBuilder();

        // Get the inner text of the slide:
        IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
        foreach (A.Text text in texts)
        {
            paragraphText.Append(text.Text);
        }
        sldText = paragraphText.ToString();
    }
}
```

### [Visual Basic](#tab/vb-6)
```vb
    Sub GetSlideIdAndText(ByRef sldText As String, docName As String, index As Integer)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = If(part?.Presentation?.SlideIdList?.ChildElements, New OpenXmlElementList())

            ' If there are no slide IDs then there are no slides.
            If slideIds.Count = 0 Then
                sldText = ""
                Return
            End If

            Dim relId As String = TryCast(slideIds(index), SlideId)?.RelationshipId

            If relId Is Nothing Then
                sldText = ""
                Return
            End If
            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

            ' Build a StringBuilder object.
            Dim paragraphText As New StringBuilder()

            ' Get the inner text of the slide:
            Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
            For Each text As A.Text In texts
                paragraphText.Append(text.Text)
            Next
            sldText = paragraphText.ToString()
        End Using
    End Sub
End Module
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
