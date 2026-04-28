# Add Transitions between slides in a presentation

This topic shows how to use the classes in the Open XML SDK to
add transition between all slides in a presentation programmatically.

## Getting a Presentation Object 

In the Open XML SDK, the `DocumentFormat.OpenXml.Packaging.PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document, call the
`DocumentFormat.OpenXml.Packaging.PresentationDocument.Open` method, that uses a file path, and a
Boolean value as the second parameter to specify whether a document is
editable. To open a document for read/write, specify the value `true` for this parameter as shown in the following
`using` statement. In this code, the file parameter, is a string that represents the path for the file from which you want to open the document.

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

## The Structure of the Transition

Transition element `<transition>` specifies the kind of slide transition that should be used to transition to the current slide from the
previous slide. That is, the transition information is stored on the slide that appears after the transition is
complete.

The following table lists the attributes of the Transition along
with the description of each.

| Attribute | Description |
|---|---|
| advClick (Advance on Click) | Specifies whether a mouse click advances the slide or not. If this attribute is not specified then a value of true is assumed. |
| advTm (Advance after time) | Specifies the time, in milliseconds, after which the transition should start. This setting can be used in conjunction with the advClick attribute. If this attribute is not specified then it is assumed that no auto-advance occurs. |
| spd (Transition Speed) |Specifies the transition speed that is to be used when transitioning from the current slide to the next. |

[*Example*: Consider the following example

```xml
      <p:transition spd="slow" advClick="1" advTm="3000">
        <p:randomBar dir="horz"/>
      </p:transition>
```
In the above example, the transition speed `<speed>` is set to slow (available options: slow, med, fast). Advance on Click `<advClick>` is set to true, and Advance after time `<advTm>` is set to 3000 milliseconds. The Random Bar child element `<randomBar>` describes the randomBar slide transition effect, which uses a set of randomly placed horizontal `<dir="horz">` or vertical `<dir="vert">` bars on the slide that continue to be added until the new slide is fully shown. *end example*]

A full list of Transition's child elements can be viewed here: `DocumentFormat.OpenXml.Presentation.Transition`

## The Structure of the Alternate Content

Office Open XML defines a mechanism for the storage of content that is not defined by the ISO/IEC 29500 Office Open XML specification, such as extensions developed by future software applications that leverage the Office Open XML formats. This mechanism allows for the storage of a series of alternative representations of content, from which the consuming application can use the first alternative whose requirements are met.

Consider an application that creates a new transition object intended to specify the duration of the transition. This functionality is not defined in the Office Open XML specification. Using an AlternateContent block as follows allows specifying the duration `<p14:dur>` in milliseconds.

[*Example*: 
```xml
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
   xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
    <mc:Choice Requires="p14">
      <p:transition spd="slow" p14:dur="2000" advClick="1" advTm="3000">
        <p:randomBar/>
      </p:transition>
    </mc:Choice>
    <mc:Fallback>
      <p:transition spd="slow" advClick="1" advTm="3000">
        <p:randomBar/>
      </p:transition>
    </mc:Fallback>
  </mc:AlternateContent>
```

The Choice element in the above example requires the `DocumentFormat.OpenXml.Linq.P14.dur` attribute to specify the duration of the transition, and the Fallback element allows clients that do not support this namespace to see an appropriate alternative representation. *end example*]

More details on the P14 class can be found here: `DocumentFormat.OpenXml.Linq.P14`

## How the Sample Code Works ##
After opening the presentation file for read/write access in the using statement, the code gets the presentation part from the presentation document. Then, it retrieves the relationship IDs of all slides in the presentation and gets the slides part from the relationship ID. The code then checks if there are no existing transitions set on the slides and replaces them with a new RandomBarTransition.

### [C#](#tab/cs-2)
```csharp
        // Define the transition start time and duration in milliseconds
        string startTransitionAfterMs = "3000", durationMs = "2000";

        // Set to true if you want to advance to the next slide on mouse click
        bool advanceOnClick = true;
     
        // Iterate through each slide ID to get slides parts
        foreach (SlideId slideId in slidesIds)
        {
            // Get the relationship ID of the slide
            string? relId = slideId!.RelationshipId!.ToString();

            if (relId == null)
            {
                throw new NullReferenceException("RelationshipId not found");
            }

            // Get the slide part using the relationship ID
            SlidePart? slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(relId);

            // Remove existing transitions if any
            if (slidePart.Slide.Transition != null)
            {
                slidePart.Slide.Transition.Remove();
            }

            // Check if there are any AlternateContent elements
            if (slidePart!.Slide.Descendants<AlternateContent>().ToList().Count > 0)
            {
                // Get all AlternateContent elements
                List<AlternateContent> alternateContents = [.. slidePart.Slide.Descendants<AlternateContent>()];
                foreach (AlternateContent alternateContent in alternateContents)
                {
                    // Remove transitions in AlternateContentChoice within AlternateContent
                    List<OpenXmlElement> childElements = alternateContent.ChildElements.ToList();

                    foreach (OpenXmlElement element in childElements)
                    {
                        List<Transition> transitions = element.Descendants<Transition>().ToList();
                        foreach (Transition transition in transitions)
                        {
                            transition.Remove();
                        }
                    }
                    // Add new transitions to AlternateContentChoice and AlternateContentFallback
                    alternateContent!.GetFirstChild<AlternateContentChoice>();
                    Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    Transition fallbackTransition = new Transition(new RandomBarTransition()) {AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    alternateContent!.GetFirstChild<AlternateContentChoice>()!.Append(choiceTransition);
                    alternateContent!.GetFirstChild<AlternateContentFallback>()!.Append(fallbackTransition);
                }
            }
```

### [Visual Basic](#tab/vb-2)
```vb
            ' Define the transition start time and duration in milliseconds
            Dim startTransitionAfterMs As String = "3000"
            Dim durationMs As String = "2000"

            ' Set to true if you want to advance to the next slide on mouse click
            Dim advanceOnClick As Boolean = True

            ' Iterate through each slide ID to get slides parts
            For Each slideId As SlideId In slidesIds
                ' Get the relationship ID of the slide
                Dim relId As String = slideId.RelationshipId.ToString()

                If relId Is Nothing Then
                    Throw New NullReferenceException("RelationshipId not found")
                End If

                ' Get the slide part using the relationship ID
                Dim slidePart As SlidePart = CType(presentationDocument.PresentationPart.GetPartById(relId), SlidePart)

                ' Remove existing transitions if any
                If slidePart.Slide.Transition IsNot Nothing Then
                    slidePart.Slide.Transition.Remove()
                End If

                ' Check if there are any AlternateContent elements
                If slidePart.Slide.Descendants(Of AlternateContent)().ToList().Count > 0 Then
                    ' Get all AlternateContent elements
                    Dim alternateContents As List(Of AlternateContent) = slidePart.Slide.Descendants(Of AlternateContent)().ToList()
                    For Each alternateContent In alternateContents
                        ' Remove transitions in AlternateContentChoice within AlternateContent
                        Dim childElements As List(Of OpenXmlElement) = alternateContent.ChildElements.ToList()

                        For Each element In childElements
                            Dim transitions As List(Of Transition) = element.Descendants(Of Transition)().ToList()
                            For Each transition In transitions
                                transition.Remove()
                            Next
                        Next
                        ' Add new transitions to AlternateContentChoice and AlternateContentFallback
                        alternateContent.GetFirstChild(Of AlternateContentChoice)()
                        Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        alternateContent.GetFirstChild(Of AlternateContentChoice)().Append(choiceTransition)
                        alternateContent.GetFirstChild(Of AlternateContentFallback)().Append(fallbackTransition)
                    Next
```
***

If there are currently no transitions on the slide, code creates new transition. In both cases as a fallback transition,
RandomBarTransition is used but without `P14:dur`(duration) to allow grater support for clients that aren't supporting this namespace

### [C#](#tab/cs-3)
```csharp
            // Add transition if there is none
            else
            {
                // Check if there is a transition appended to the slide and set it to null
                if (slidePart.Slide.Transition != null)
                {
                    slidePart.Slide.Transition = null;
                }
                // Create a new AlternateContent element
                AlternateContent alternateContent = new AlternateContent();
                alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                // Create a new AlternateContentChoice element and add the transition
                AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "p14" };
                Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                Transition fallbackTransition = new Transition(new RandomBarTransition()) { AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                alternateContentChoice.Append(choiceTransition);

                // Create a new AlternateContentFallback element and add the transition
                AlternateContentFallback alternateContentFallback = new AlternateContentFallback(fallbackTransition);
                alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main");
                alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative");
                alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main");

                // Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                alternateContent.Append(alternateContentChoice);
                alternateContent.Append(alternateContentFallback);
                slidePart.Slide.Append(alternateContent);
            } 
```

### [Visual Basic](#tab/vb-3)
```vb
                    ' Add transition if there is none
                Else
                    ' Check if there is a transition appended to the slide and set it to null
                    If slidePart.Slide.Transition IsNot Nothing Then
                        slidePart.Slide.Transition = Nothing
                    End If

                    ' Create a new AlternateContent element
                    Dim alternateContent As New AlternateContent()
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")

                    ' Create a new AlternateContentChoice element and add the transition
                    Dim alternateContentChoice As New AlternateContentChoice() With {
                    .Requires = "p14"
                }
                    Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    alternateContentChoice.Append(choiceTransition)

                    ' Create a new AlternateContentFallback element and add the transition
                    Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    Dim alternateContentFallback As New AlternateContentFallback(fallbackTransition)

                    alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main")
                    alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main")
                    alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative")
                    alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main")

                    ' Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                    alternateContent.Append(alternateContentChoice)
                    alternateContent.Append(alternateContentFallback)
                    slidePart.Slide.Append(alternateContent)
                End If
```
***

## Sample Code

Following is the complete sample code that you can use to add RandomBarTransition to all slides.

### [C#](#tab/cs)
```csharp
AddTransmitionToSlides(args[0]);
static void AddTransmitionToSlides(string filePath)
{
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
    {
    
     // Check if the presentation part and slide list are available
        if (presentationDocument.PresentationPart == null || presentationDocument.PresentationPart.Presentation.SlideIdList == null)
        {
            throw new NullReferenceException("Presentation part is empty or there are no slides");
        }

        // Get the presentation part
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the list of slide IDs
        OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;
        // Define the transition start time and duration in milliseconds
        string startTransitionAfterMs = "3000", durationMs = "2000";

        // Set to true if you want to advance to the next slide on mouse click
        bool advanceOnClick = true;
     
        // Iterate through each slide ID to get slides parts
        foreach (SlideId slideId in slidesIds)
        {
            // Get the relationship ID of the slide
            string? relId = slideId!.RelationshipId!.ToString();

            if (relId == null)
            {
                throw new NullReferenceException("RelationshipId not found");
            }

            // Get the slide part using the relationship ID
            SlidePart? slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(relId);

            // Remove existing transitions if any
            if (slidePart.Slide.Transition != null)
            {
                slidePart.Slide.Transition.Remove();
            }

            // Check if there are any AlternateContent elements
            if (slidePart!.Slide.Descendants<AlternateContent>().ToList().Count > 0)
            {
                // Get all AlternateContent elements
                List<AlternateContent> alternateContents = [.. slidePart.Slide.Descendants<AlternateContent>()];
                foreach (AlternateContent alternateContent in alternateContents)
                {
                    // Remove transitions in AlternateContentChoice within AlternateContent
                    List<OpenXmlElement> childElements = alternateContent.ChildElements.ToList();

                    foreach (OpenXmlElement element in childElements)
                    {
                        List<Transition> transitions = element.Descendants<Transition>().ToList();
                        foreach (Transition transition in transitions)
                        {
                            transition.Remove();
                        }
                    }
                    // Add new transitions to AlternateContentChoice and AlternateContentFallback
                    alternateContent!.GetFirstChild<AlternateContentChoice>();
                    Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    Transition fallbackTransition = new Transition(new RandomBarTransition()) {AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    alternateContent!.GetFirstChild<AlternateContentChoice>()!.Append(choiceTransition);
                    alternateContent!.GetFirstChild<AlternateContentFallback>()!.Append(fallbackTransition);
                }
            }
            // Add transition if there is none
            else
            {
                // Check if there is a transition appended to the slide and set it to null
                if (slidePart.Slide.Transition != null)
                {
                    slidePart.Slide.Transition = null;
                }
                // Create a new AlternateContent element
                AlternateContent alternateContent = new AlternateContent();
                alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                // Create a new AlternateContentChoice element and add the transition
                AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "p14" };
                Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                Transition fallbackTransition = new Transition(new RandomBarTransition()) { AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                alternateContentChoice.Append(choiceTransition);

                // Create a new AlternateContentFallback element and add the transition
                AlternateContentFallback alternateContentFallback = new AlternateContentFallback(fallbackTransition);
                alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main");
                alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative");
                alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main");

                // Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                alternateContent.Append(alternateContentChoice);
                alternateContent.Append(alternateContentFallback);
                slidePart.Slide.Append(alternateContent);
            } 
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub AddTransitionToSlides(filePath As String)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(filePath, True)
            ' Check if the presentation part and slide list are available
            If presentationDocument.PresentationPart Is Nothing OrElse presentationDocument.PresentationPart.Presentation.SlideIdList Is Nothing Then
                Throw New NullReferenceException("Presentation part is empty or there are no slides")
            End If

            ' Get the presentation part
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Get the list of slide IDs
            Dim slidesIds As OpenXmlElementList = presentationPart.Presentation.SlideIdList.ChildElements
            ' Define the transition start time and duration in milliseconds
            Dim startTransitionAfterMs As String = "3000"
            Dim durationMs As String = "2000"

            ' Set to true if you want to advance to the next slide on mouse click
            Dim advanceOnClick As Boolean = True

            ' Iterate through each slide ID to get slides parts
            For Each slideId As SlideId In slidesIds
                ' Get the relationship ID of the slide
                Dim relId As String = slideId.RelationshipId.ToString()

                If relId Is Nothing Then
                    Throw New NullReferenceException("RelationshipId not found")
                End If

                ' Get the slide part using the relationship ID
                Dim slidePart As SlidePart = CType(presentationDocument.PresentationPart.GetPartById(relId), SlidePart)

                ' Remove existing transitions if any
                If slidePart.Slide.Transition IsNot Nothing Then
                    slidePart.Slide.Transition.Remove()
                End If

                ' Check if there are any AlternateContent elements
                If slidePart.Slide.Descendants(Of AlternateContent)().ToList().Count > 0 Then
                    ' Get all AlternateContent elements
                    Dim alternateContents As List(Of AlternateContent) = slidePart.Slide.Descendants(Of AlternateContent)().ToList()
                    For Each alternateContent In alternateContents
                        ' Remove transitions in AlternateContentChoice within AlternateContent
                        Dim childElements As List(Of OpenXmlElement) = alternateContent.ChildElements.ToList()

                        For Each element In childElements
                            Dim transitions As List(Of Transition) = element.Descendants(Of Transition)().ToList()
                            For Each transition In transitions
                                transition.Remove()
                            Next
                        Next
                        ' Add new transitions to AlternateContentChoice and AlternateContentFallback
                        alternateContent.GetFirstChild(Of AlternateContentChoice)()
                        Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        alternateContent.GetFirstChild(Of AlternateContentChoice)().Append(choiceTransition)
                        alternateContent.GetFirstChild(Of AlternateContentFallback)().Append(fallbackTransition)
                    Next
                    ' Add transition if there is none
                Else
                    ' Check if there is a transition appended to the slide and set it to null
                    If slidePart.Slide.Transition IsNot Nothing Then
                        slidePart.Slide.Transition = Nothing
                    End If

                    ' Create a new AlternateContent element
                    Dim alternateContent As New AlternateContent()
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")

                    ' Create a new AlternateContentChoice element and add the transition
                    Dim alternateContentChoice As New AlternateContentChoice() With {
                    .Requires = "p14"
                }
                    Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    alternateContentChoice.Append(choiceTransition)

                    ' Create a new AlternateContentFallback element and add the transition
                    Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    Dim alternateContentFallback As New AlternateContentFallback(fallbackTransition)

                    alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main")
                    alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main")
                    alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative")
                    alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main")

                    ' Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                    alternateContent.Append(alternateContentChoice)
                    alternateContent.Append(alternateContentFallback)
                    slidePart.Slide.Append(alternateContent)
                End If
            Next
        End Using
    End Sub

End Module
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
