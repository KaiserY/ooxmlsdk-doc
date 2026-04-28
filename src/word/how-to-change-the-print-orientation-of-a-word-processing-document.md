# Change the print orientation of a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically set the print orientation of a Microsoft Word document. It contains an example
`SetPrintOrientation` method to illustrate this task.

-----------------------------------------------------------------------------

## SetPrintOrientation Method

You can use the `SetPrintOrientation` method
to change the print orientation of a word processing document. The
method accepts two parameters that indicate the name of the document to
modify (string) and the new print orientation (`DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues`).

The following code shows the `SetPrintOrientation` method.

### [C#](#tab/cs-0)
```csharp
static void SetPrintOrientation(string fileName, string orientation)
```
### [Visual Basic](#tab/vb-0)
```vb
    Sub SetPrintOrientation(fileName As String, orientation As String)
```
***

For each section in the document, if the new orientation differs from
the section's current print orientation, the code modifies the print
orientation for the section. In addition, the code must manually update
the width, height, and margins for each section.

-----------------------------------------------------------------------------

## Calling the Sample SetPrintOrientation Method

To call the sample `SetPrintOrientation`
method, pass a string that contains the name of the file to convert and the string "landscape" or "portrait"
depending on which orientation you want. The following code shows an example method call.

### [C#](#tab/cs-1)
```csharp
SetPrintOrientation(args[0], args[1]);
```
### [Visual Basic](#tab/vb-1)
```vb
        SetPrintOrientation(args(0), args(1))
```
***

-----------------------------------------------------------------------------

## How the Code Works

The following code first determines which orientation to apply and
then opens the document by using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A`
method and sets the `isEditable` parameter to
`true` to indicate that the document should
be read/write. The code retrieves a reference to the main
document part, and then uses that reference to retrieve a collection of
all of the descendants of type `DocumentFormat.OpenXml.Wordprocessing.SectionProperties` within the content of the
document. Later code will use this collection to set the orientation for
each section in turn.

### [C#](#tab/cs-2)
```csharp
    PageOrientationValues newOrientation = orientation.ToLower() switch
    {
        "landscape" => PageOrientationValues.Landscape,
        "portrait" => PageOrientationValues.Portrait,
        _ => throw new System.ArgumentException("Invalid argument: " + orientation)
    };

    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        if (document?.MainDocumentPart?.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        Body docBody = document.MainDocumentPart.Document.Body;

        IEnumerable<SectionProperties> sections = docBody.ChildElements.OfType<SectionProperties>();

        if (sections.Count() == 0)
        {
            docBody.AddChild(new SectionProperties());

            sections = docBody.ChildElements.OfType<SectionProperties>();
        }
```
### [Visual Basic](#tab/vb-2)
```vb
        Dim newOrientation As PageOrientationValues

        Select Case orientation.ToLower()
            Case "landscape"
                newOrientation = PageOrientationValues.Landscape
            Case "portrait"
                newOrientation = PageOrientationValues.Portrait
            Case Else
                Throw New ArgumentException("Invalid argument: " & orientation)
        End Select

        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If document?.MainDocumentPart?.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Dim docBody As Body = document.MainDocumentPart.Document.Body

            Dim sections As IEnumerable(Of SectionProperties) = docBody.ChildElements.OfType(Of SectionProperties)()

            If sections.Count() = 0 Then
                docBody.AddChild(New SectionProperties())

                sections = docBody.ChildElements.OfType(Of SectionProperties)()
            End If
```
***

-----------------------------------------------------------------------------

## Iterating Through All the Sections

The next block of code iterates through all the sections in the collection of `SectionProperties` elements. For each section, the code initializes a variable that tracks whether the page orientation for the section was changed so the code can update the page size and margins. (If the new orientation matches the original orientation, the code will not update the page.) The code continues by retrieving a reference to the first `DocumentFormat.OpenXml.Wordprocessing.PageSize` descendant of the `SectionProperties` element. If the reference is not null, the code updates the orientation as required.

### [C#](#tab/cs-3)
```csharp
        foreach (SectionProperties sectPr in sections)
        {
            bool pageOrientationChanged = false;

            PageSize pgSz = sectPr.ChildElements.OfType<PageSize>().FirstOrDefault() ?? sectPr.AppendChild(new PageSize() { Width = 12240, Height = 15840 });
```
### [Visual Basic](#tab/vb-3)
```vb
            For Each sectPr As SectionProperties In sections
                Dim pageOrientationChanged As Boolean = False

                Dim pgSz As PageSize = If(sectPr.ChildElements.OfType(Of PageSize)().FirstOrDefault(), sectPr.AppendChild(New PageSize() With {.Width = 12240, .Height = 15840}))
```
***

-----------------------------------------------------------------------------

## Setting the Orientation for the Section

The next block of code first checks whether the `DocumentFormat.OpenXml.Wordprocessing.PageSize.Orient`
property of the `PageSize` element exists. As with many properties
of Open XML elements, the property or attribute might not exist yet. In
that case, retrieving the property returns a null reference. By default,
if the property does not exist, and the new orientation is Portrait, the
code will not update the page. If the `Orient` property already exists, and its value
differs from the new orientation value supplied as a parameter to the
method, the code sets the `Value` property of
the `Orient` property, and sets the
`pageOrientationChanged` flag. (The code uses the `pageOrientationChanged` flag to determine whether it
must update the page size and margins.)

> **Note**
> If the code must create the `Orient` property, it must also create the value to store in the property, as a new `DocumentFormat.OpenXml.EnumValue%601` instance, supplying the new orientation in the `EnumValue` constructor.

### [C#](#tab/cs-4)
```csharp
            if (pgSz.Orient is null)
            {
                // Need to create the attribute. You do not need to
                // create the Orient property if the property does not
                // already exist, and you are setting it to Portrait.
                // That is the default value.
                if (newOrientation != PageOrientationValues.Portrait)
                {
                    pageOrientationChanged = true;
                    pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                }
            }
            else
            {
                // The Orient property exists, but its value
                // is different than the new value.
                if (pgSz.Orient.Value != newOrientation)
                {
                    pgSz.Orient.Value = newOrientation;
                    pageOrientationChanged = true;
                }
```
### [Visual Basic](#tab/vb-4)
```vb
                If pgSz.Orient Is Nothing Then
                    ' Need to create the attribute. You do not need to 
                    ' create the Orient property if the property does not 
                    ' already exist, and you are setting it to Portrait. 
                    ' That is the default value.
                    If newOrientation <> PageOrientationValues.Portrait Then
                        pageOrientationChanged = True
                        pgSz.Orient = New EnumValue(Of PageOrientationValues)(newOrientation)
                    End If
                Else
                    ' The Orient property exists, but its value
                    ' is different than the new value.
                    If pgSz.Orient.Value <> newOrientation Then
                        pgSz.Orient.Value = newOrientation
                        pageOrientationChanged = True
                    End If
                End If
```
***

-----------------------------------------------------------------------------

## Updating the Page Size

At this point in the code, the page orientation may have changed. If so,
the code must complete two more tasks. It must update the page size, and
update the page margins for the section. The first task is easy—the
following code just swaps the page height and width, storing the values
in the `PageSize` element.

### [C#](#tab/cs-5)
```csharp
                if (pageOrientationChanged)
                {
                    // Changing the orientation is not enough. You must also
                    // change the page size.
                    var width = pgSz.Width;
                    var height = pgSz.Height;
                    pgSz.Width = height;
                    pgSz.Height = width;
```
### [Visual Basic](#tab/vb-5)
```vb
                If pageOrientationChanged Then
                    ' Changing the orientation is not enough. You must also 
                    ' change the page size.
                    Dim width = pgSz.Width
                    Dim height = pgSz.Height
                    pgSz.Width = height
                    pgSz.Height = width
```
***

-----------------------------------------------------------------------------

## Updating the Margins

The next step in the sample procedure handles margins for the section.
If the page orientation has changed, the code must rotate the margins to
match. To do so, the code retrieves a reference to the `DocumentFormat.OpenXml.Wordprocessing.PageMargin` element for the section. If the element exists, the code rotates the margins. Note that the code rotates
the margins by 90 degrees—some printers rotate the margins by 270
degrees instead and you could modify the code to take that into account.
Also be aware that the `DocumentFormat.OpenXml.Wordprocessing.PageMargin.Top` and `DocumentFormat.OpenXml.Wordprocessing.PageMargin.Bottom` properties of the `PageMargin` object are signed values, and the
`DocumentFormat.OpenXml.Wordprocessing.PageMargin.Left` and `DocumentFormat.OpenXml.Wordprocessing.PageMargin.Right` properties are unsigned values. The code must convert between the two types of values as it rotates the
margin settings, as shown in the following code.

### [C#](#tab/cs-6)
```csharp
                    PageMargin? pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();

                    if (pgMar is not null)
                    {
                        // Rotate margins. Printer settings control how far you
                        // rotate when switching to landscape mode. Not having those
                        // settings, this code rotates 90 degrees. You could easily
                        // modify this behavior, or make it a parameter for the
                        // procedure.
                        if (pgMar.Top is null || pgMar.Bottom is null || pgMar.Left is null || pgMar.Right is null)
                        {
                            throw new ArgumentNullException("One or more of the PageMargin elements is null.");
                        }

                        var top = pgMar.Top.Value;
                        var bottom = pgMar.Bottom.Value;
                        var left = pgMar.Left.Value;
                        var right = pgMar.Right.Value;

                        pgMar.Top = new Int32Value((int)left);
                        pgMar.Bottom = new Int32Value((int)right);
                        pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                        pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                    }
```
### [Visual Basic](#tab/vb-6)
```vb
                    Dim pgMar As PageMargin = sectPr.Descendants(Of PageMargin)().FirstOrDefault()

                    If pgMar IsNot Nothing Then
                        ' Rotate margins. Printer settings control how far you 
                        ' rotate when switching to landscape mode. Not having those
                        ' settings, this code rotates 90 degrees. You could easily
                        ' modify this behavior, or make it a parameter for the 
                        ' procedure.
                        If pgMar.Top Is Nothing OrElse pgMar.Bottom Is Nothing OrElse pgMar.Left Is Nothing OrElse pgMar.Right Is Nothing Then
                            Throw New ArgumentNullException("One or more of the PageMargin elements is null.")
                        End If

                        Dim top = pgMar.Top.Value
                        Dim bottom = pgMar.Bottom.Value
                        Dim left = pgMar.Left.Value
                        Dim right = pgMar.Right.Value

                        pgMar.Top = New Int32Value(CInt(left))
                        pgMar.Bottom = New Int32Value(CInt(right))
                        pgMar.Left = New UInt32Value(CUInt(System.Math.Max(0, bottom)))
                        pgMar.Right = New UInt32Value(CUInt(System.Math.Max(0, top)))
                    End If
```
***

-----------------------------------------------------------------------------

## Sample Code

The following is the complete `SetPrintOrientation` code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
```csharp
static void SetPrintOrientation(string fileName, string orientation)
{
    PageOrientationValues newOrientation = orientation.ToLower() switch
    {
        "landscape" => PageOrientationValues.Landscape,
        "portrait" => PageOrientationValues.Portrait,
        _ => throw new System.ArgumentException("Invalid argument: " + orientation)
    };

    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        if (document?.MainDocumentPart?.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        Body docBody = document.MainDocumentPart.Document.Body;

        IEnumerable<SectionProperties> sections = docBody.ChildElements.OfType<SectionProperties>();

        if (sections.Count() == 0)
        {
            docBody.AddChild(new SectionProperties());

            sections = docBody.ChildElements.OfType<SectionProperties>();
        }
        foreach (SectionProperties sectPr in sections)
        {
            bool pageOrientationChanged = false;

            PageSize pgSz = sectPr.ChildElements.OfType<PageSize>().FirstOrDefault() ?? sectPr.AppendChild(new PageSize() { Width = 12240, Height = 15840 });
            // No Orient property? Create it now. Otherwise, just
            // set its value. Assume that the default orientation  is Portrait.
            if (pgSz.Orient is null)
            {
                // Need to create the attribute. You do not need to
                // create the Orient property if the property does not
                // already exist, and you are setting it to Portrait.
                // That is the default value.
                if (newOrientation != PageOrientationValues.Portrait)
                {
                    pageOrientationChanged = true;
                    pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                }
            }
            else
            {
                // The Orient property exists, but its value
                // is different than the new value.
                if (pgSz.Orient.Value != newOrientation)
                {
                    pgSz.Orient.Value = newOrientation;
                    pageOrientationChanged = true;
                }
                if (pageOrientationChanged)
                {
                    // Changing the orientation is not enough. You must also
                    // change the page size.
                    var width = pgSz.Width;
                    var height = pgSz.Height;
                    pgSz.Width = height;
                    pgSz.Height = width;
                    PageMargin? pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();

                    if (pgMar is not null)
                    {
                        // Rotate margins. Printer settings control how far you
                        // rotate when switching to landscape mode. Not having those
                        // settings, this code rotates 90 degrees. You could easily
                        // modify this behavior, or make it a parameter for the
                        // procedure.
                        if (pgMar.Top is null || pgMar.Bottom is null || pgMar.Left is null || pgMar.Right is null)
                        {
                            throw new ArgumentNullException("One or more of the PageMargin elements is null.");
                        }

                        var top = pgMar.Top.Value;
                        var bottom = pgMar.Bottom.Value;
                        var left = pgMar.Left.Value;
                        var right = pgMar.Right.Value;

                        pgMar.Top = new Int32Value((int)left);
                        pgMar.Bottom = new Int32Value((int)right);
                        pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                        pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                    }
                }
            }
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub SetPrintOrientation(fileName As String, orientation As String)
        Dim newOrientation As PageOrientationValues

        Select Case orientation.ToLower()
            Case "landscape"
                newOrientation = PageOrientationValues.Landscape
            Case "portrait"
                newOrientation = PageOrientationValues.Portrait
            Case Else
                Throw New ArgumentException("Invalid argument: " & orientation)
        End Select

        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If document?.MainDocumentPart?.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Dim docBody As Body = document.MainDocumentPart.Document.Body

            Dim sections As IEnumerable(Of SectionProperties) = docBody.ChildElements.OfType(Of SectionProperties)()

            If sections.Count() = 0 Then
                docBody.AddChild(New SectionProperties())

                sections = docBody.ChildElements.OfType(Of SectionProperties)()
            End If
            For Each sectPr As SectionProperties In sections
                Dim pageOrientationChanged As Boolean = False

                Dim pgSz As PageSize = If(sectPr.ChildElements.OfType(Of PageSize)().FirstOrDefault(), sectPr.AppendChild(New PageSize() With {.Width = 12240, .Height = 15840}))
                ' No Orient property? Create it now. Otherwise, just 
                ' set its value. Assume that the default orientation  is Portrait.
                If pgSz.Orient Is Nothing Then
                    ' Need to create the attribute. You do not need to 
                    ' create the Orient property if the property does not 
                    ' already exist, and you are setting it to Portrait. 
                    ' That is the default value.
                    If newOrientation <> PageOrientationValues.Portrait Then
                        pageOrientationChanged = True
                        pgSz.Orient = New EnumValue(Of PageOrientationValues)(newOrientation)
                    End If
                Else
                    ' The Orient property exists, but its value
                    ' is different than the new value.
                    If pgSz.Orient.Value <> newOrientation Then
                        pgSz.Orient.Value = newOrientation
                        pageOrientationChanged = True
                    End If
                End If
                If pageOrientationChanged Then
                    ' Changing the orientation is not enough. You must also 
                    ' change the page size.
                    Dim width = pgSz.Width
                    Dim height = pgSz.Height
                    pgSz.Width = height
                    pgSz.Height = width
                    Dim pgMar As PageMargin = sectPr.Descendants(Of PageMargin)().FirstOrDefault()

                    If pgMar IsNot Nothing Then
                        ' Rotate margins. Printer settings control how far you 
                        ' rotate when switching to landscape mode. Not having those
                        ' settings, this code rotates 90 degrees. You could easily
                        ' modify this behavior, or make it a parameter for the 
                        ' procedure.
                        If pgMar.Top Is Nothing OrElse pgMar.Bottom Is Nothing OrElse pgMar.Left Is Nothing OrElse pgMar.Right Is Nothing Then
                            Throw New ArgumentNullException("One or more of the PageMargin elements is null.")
                        End If

                        Dim top = pgMar.Top.Value
                        Dim bottom = pgMar.Bottom.Value
                        Dim left = pgMar.Left.Value
                        Dim right = pgMar.Right.Value

                        pgMar.Top = New Int32Value(CInt(left))
                        pgMar.Bottom = New Int32Value(CInt(right))
                        pgMar.Left = New UInt32Value(CUInt(System.Math.Max(0, bottom)))
                        pgMar.Right = New UInt32Value(CUInt(System.Math.Max(0, top)))
                    End If
                End If
            Next
        End Using
    End Sub
```
***

-----------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
