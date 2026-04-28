# Retrieve the number of slides in a presentation document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve the number of slides in a
presentation document, either including hidden slides or not, without
loading the document into Microsoft PowerPoint. It contains an example
`RetrieveNumberOfSlides` method to illustrate
this task.

---------------------------------------------------------------------------------

## RetrieveNumberOfSlides Method

You can use the `RetrieveNumberOfSlides`
method to get the number of slides in a presentation document,
optionally including the hidden slides. The `RetrieveNumberOfSlides` method accepts two
parameters: a string that indicates the path of the file that you want
to examine, and an optional Boolean value that indicates whether to
include hidden slides in the count.

### [C#](#tab/cs-0)
```csharp
static int RetrieveNumberOfSlides(string fileName, string includeHidden = "true")
```

### [Visual Basic](#tab/vb-0)
```vb
    Function RetrieveNumberOfSlides(fileName As String, Optional includeHidden As String = "true") As Integer
```
***

---------------------------------------------------------------------------------
## Calling the RetrieveNumberOfSlides Method

The method returns an integer that indicates the number of slides,
counting either all the slides or only visible slides, depending on the
second parameter value. To call the method, pass all the parameter
values, as shown in the following code.

### [C#](#tab/cs-1)
```csharp
if (args is [{ } fileName, { } includeHidden])
{
    RetrieveNumberOfSlides(fileName, includeHidden);
}
else if (args is [{ } fileName2])
{
    RetrieveNumberOfSlides(fileName2);
}
```

### [Visual Basic](#tab/vb-1)
```vb
        If args.Length = 2 Then
            RetrieveNumberOfSlides(args(0), args(1))
        ElseIf args.Length = 1 Then
            RetrieveNumberOfSlides(args(0))
        End If
```
***

---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating an integer variable, `slidesCount`, to hold the number of slides. The code then opens the specified presentation by using the `DocumentFormat.OpenXml.Packaging.PresentationDocument.Open` method and indicating that the document should be open for read-only access (the
final `false` parameter value). Given the open presentation, the code uses the `DocumentFormat.OpenXml.Packaging.PresentationDocument.PresentationPart` property to navigate to the main presentation part, storing the reference in a variable named `presentationPart`.

### [C#](#tab/cs)
```csharp
    using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
    {
        if (doc.PresentationPart is not null)
        {
            // Get the presentation part of the document.
            PresentationPart presentationPart = doc.PresentationPart;
```

### [Visual Basic](#tab/vb)
```vb
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, False)
            If doc.PresentationPart IsNot Nothing Then
                ' Get the presentation part of the document.
                Dim presentationPart As PresentationPart = doc.PresentationPart
```
***

---------------------------------------------------------------------------------

## Retrieving the Count of All Slides

If the presentation part reference is not null (and it will not be, for any valid presentation that loads correctly into PowerPoint), the code next calls the `Count` method on the value of the `DocumentFormat.OpenXml.Packaging.PresentationPart.SlideParts` property of the presentation part. If you requested all slides, including hidden slides, that is all there is to do. There is slightly more work to be done if you want to exclude hidden slides, as shown in the following code.

### [C#](#tab/cs)
```csharp
                if (includeHidden.ToUpper() == "TRUE")
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                else
                {
```

### [Visual Basic](#tab/vb)
```vb
                    If includeHidden.ToUpper() = "TRUE" Then
                        slidesCount = presentationPart.SlideParts.Count()
                    Else
```
***

---------------------------------------------------------------------------------

## Retrieving the Count of Visible Slides

If you requested that the code should limit the return value to include
only visible slides, the code must filter its collection of slides to
include only those slides that have a `DocumentFormat.OpenXml.Presentation.Slide.Show` property that contains a value, and
the value is `true`. If the `Show` property is null, that also indicates that
the slide is visible. This is the most likely scenario. PowerPoint does
not set the value of this property, in general, unless the slide is to
be hidden. The only way the `Show` property
would exist and have a value of `true` would
be if you had hidden and then unhidden the slide. The following code
uses the `System.Linq.Enumerable.Where`
function with a lambda expression to do the work.

### [C#](#tab/cs)
```csharp
                    // Each slide can include a Show property, which if hidden 
                    // will contain the value "0". The Show property may not 
                    // exist, and most likely will not, for non-hidden slides.
                    var slides = presentationPart.SlideParts.Where(
                        (s) => (s.Slide is not null) &&
                          ((s.Slide.Show is null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));

                    slidesCount = slides.Count();
```

### [Visual Basic](#tab/vb)
```vb
                        ' Each slide can include a Show property, which if hidden 
                        ' will contain the value "0". The Show property may not 
                        ' exist, and most likely will not, for non-hidden slides.
                        Dim slides = presentationPart.SlideParts.Where(
                            Function(s) (s.Slide IsNot Nothing) AndAlso
                                        ((s.Slide.Show Is Nothing) OrElse (s.Slide.Show.HasValue AndAlso s.Slide.Show.Value)))

                        slidesCount = slides.Count()
```
***

---------------------------------------------------------------------------------

## Sample Code

The following is the complete `RetrieveNumberOfSlides` code sample in C\# and
Visual Basic.

### [C#](#tab/cs-2)
```csharp
if (args is [{ } fileName, { } includeHidden])
{
    RetrieveNumberOfSlides(fileName, includeHidden);
}
else if (args is [{ } fileName2])
{
    RetrieveNumberOfSlides(fileName2);
}
static int RetrieveNumberOfSlides(string fileName, string includeHidden = "true")
{
    int slidesCount = 0;
    using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
    {
        if (doc.PresentationPart is not null)
        {
            // Get the presentation part of the document.
            PresentationPart presentationPart = doc.PresentationPart;
            if (presentationPart is not null)
            {
                if (includeHidden.ToUpper() == "TRUE")
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                else
                {
                    // Each slide can include a Show property, which if hidden 
                    // will contain the value "0". The Show property may not 
                    // exist, and most likely will not, for non-hidden slides.
                    var slides = presentationPart.SlideParts.Where(
                        (s) => (s.Slide is not null) &&
                          ((s.Slide.Show is null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));

                    slidesCount = slides.Count();
                }
            }
        }
    }

    Console.WriteLine($"Slide Count: {slidesCount}");

    return slidesCount;
}
```

### [Visual Basic](#tab/vb-2)
```vb
        If args.Length = 2 Then
            RetrieveNumberOfSlides(args(0), args(1))
        ElseIf args.Length = 1 Then
            RetrieveNumberOfSlides(args(0))
        End If
    End Sub
    Function RetrieveNumberOfSlides(fileName As String, Optional includeHidden As String = "true") As Integer
        Dim slidesCount As Integer = 0
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, False)
            If doc.PresentationPart IsNot Nothing Then
                ' Get the presentation part of the document.
                Dim presentationPart As PresentationPart = doc.PresentationPart
                If presentationPart IsNot Nothing Then
                    If includeHidden.ToUpper() = "TRUE" Then
                        slidesCount = presentationPart.SlideParts.Count()
                    Else
                        ' Each slide can include a Show property, which if hidden 
                        ' will contain the value "0". The Show property may not 
                        ' exist, and most likely will not, for non-hidden slides.
                        Dim slides = presentationPart.SlideParts.Where(
                            Function(s) (s.Slide IsNot Nothing) AndAlso
                                        ((s.Slide.Show Is Nothing) OrElse (s.Slide.Show.HasValue AndAlso s.Slide.Show.Value)))

                        slidesCount = slides.Count()
                    End If
                End If
            End If
        End Using

        Console.WriteLine($"Slide Count: {slidesCount}")

        Return slidesCount
    End Function
```

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
