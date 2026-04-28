# Extract styles from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically extract the styles or stylesWithEffects part
from a word processing document to an `System.Xml.Linq.XDocument`
instance. It contains an example `ExtractStylesPart` method to
illustrate this task.

---------------------------------------------------------------------------------

## ExtractStylesPart Method

You can use the `ExtractStylesPart` sample method to retrieve an `XDocument` instance that contains the styles or
stylesWithEffects part for a Microsoft Word document. Be aware that in a document created in Word 2010, there will
only be a single styles part; Word 2013+ adds a second stylesWithEffects
part. To provide for "round-tripping" a document from Word 2013+ to Word
2010 and back, Word 2013+ maintains both the original styles part and the
new styles part. (The Office Open XML File Formats specification
requires that Microsoft Word ignore any parts that it does not
recognize; Word 2010 does not notice the stylesWithEffects part that
Word 2013+ adds to the document.) You (and your application) must
interpret the results of retrieving the styles or stylesWithEffects
part.

The `ExtractStylesPart` procedure accepts a two parameters: the first
parameter contains a string indicating the path of the file from which
you want to extract styles, and the second indicates whether you want to
retrieve the styles part, or the newer stylesWithEffects part
(basically, you must call this procedure two times for Word 2013+
documents, retrieving each the part). The procedure returns an `XDocument` instance that contains the complete
styles or stylesWithEffects part that you requested, with all the style
information for the document (or a null reference, if the part you
requested does not exist).

### [C#](#tab/cs-0)
```csharp
static XDocument? ExtractStylesPart(string fileName, string getStylesWithEffectsPart = "true")
```
### [Visual Basic](#tab/vb-0)
```vb
    Public Function ExtractStylesPart(ByVal fileName As String, Optional ByVal getStylesWithEffectsPart As String = "true") As XDocument
```
***

The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

---------------------------------------------------------------------------------

## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document from which to extract the styles,
and a Boolean for the second parameter that specifies whether the type
of part to retrieve is the styleWithEffects part (`true`), or the styles part (`false`). The following sample code shows an example.
When you have the `XDocument` instance you
can do what you want with it; in the following sample code the content
of the `XDocument` instance is displayed to
the console.

### [C#](#tab/cs-1)
```csharp
if (args is [{ } fileName, { } getStyleWithEffectsPart])
{
    var styles = ExtractStylesPart(fileName, getStyleWithEffectsPart);

    if (styles is not null)
    {
        Console.WriteLine(styles.ToString());
    }
}
else if (args is [{ } fileName2])
{
    var styles = ExtractStylesPart(fileName2);

    if (styles is not null)
    {
        Console.WriteLine(styles.ToString());
    }
}
```
### [Visual Basic](#tab/vb-1)
```vb
        If args.Length >= 2 Then
            Dim fileName As String = args(0)
            Dim getStyleWithEffectsPart As String = args(1)

            Dim styles As XDocument = ExtractStylesPart(fileName, getStyleWithEffectsPart)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            Dim styles As XDocument = ExtractStylesPart(fileName)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        End If
```
***

---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating a variable named `styles` to contain the return value for the method.
The code continues by opening the document by using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A`
method and indicating that the document should be open for read-only access (the final false
parameter). Given the open document, the code uses the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart`
property to navigate to the main document part, and then prepares a variable named `stylesPart` to hold a reference to the styles part.

### [C#](#tab/cs-3)
```csharp
    // Declare a variable to hold the XDocument.
    XDocument? styles = null;

    // Open the document for read access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, false))
    {
        if (
            document.MainDocumentPart is null ||
            (document.MainDocumentPart.StyleDefinitionsPart is null && document.MainDocumentPart.StylesWithEffectsPart is null)
        )
        {
            throw new ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.");
        }

        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart? stylesPart = null;
```
### [Visual Basic](#tab/vb-3)
```vb
        ' Declare a variable to hold the XDocument.
        Dim styles As XDocument = Nothing

        ' Open the document for read access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, False)

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            ' Assign a reference to the appropriate part to the 
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing
```
***

---------------------------------------------------------------------------------

## Find the Correct Styles Part

The code next retrieves a reference to the requested styles part by
using the `getStylesWithEffectsPart` `System.Boolean` parameter.
Based on this value, the code retrieves a specific property
of the `docPart` variable, and stores it in the
`stylesPart` variable.

### [C#](#tab/cs-4)
```csharp
        if (getStylesWithEffectsPart.ToLower() == "true")
        {
            stylesPart = docPart.StylesWithEffectsPart;
        }
        else
        {
            stylesPart = docPart.StyleDefinitionsPart;
        }
```
### [Visual Basic](#tab/vb-4)
```vb
            If getStylesWithEffectsPart.ToLower() = "true" Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If
```
***

---------------------------------------------------------------------------------

## Retrieve the Part Contents

If the requested styles part exists, the code must return the contents
of the part in an `XDocument` instance. Each part provides a
`DocumentFormat.OpenXml.Packaging.OpenXmlPart.GetStream` method, which returns a Stream.
The code passes the Stream instance to the `System.Xml.XmlReader.Create%2A`
method, and then calls the `System.Xml.Linq.XDocument.Load%2A`
method, passing the `XmlNodeReader` as a parameter.

### [C#](#tab/cs-5)
```csharp
        if (stylesPart is not null)
        {
            using var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read));

            // Create the XDocument.
            styles = XDocument.Load(reader);
        }
```
### [Visual Basic](#tab/vb-5)
```vb
            ' If the part exists, read it into the XDocument.
            If stylesPart IsNot Nothing Then
                Using reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read))

                    ' Create the XDocument:  
                    styles = XDocument.Load(reader)
                End Using
            End If
        End Using
```
***

---------------------------------------------------------------------------------

## Sample Code

The following is the complete **ExtractStylesPart** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;

// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
static XDocument? ExtractStylesPart(string fileName, string getStylesWithEffectsPart = "true")
```

### [Visual Basic](#tab/vb)
```vb
Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        If args.Length >= 2 Then
            Dim fileName As String = args(0)
            Dim getStyleWithEffectsPart As String = args(1)

            Dim styles As XDocument = ExtractStylesPart(fileName, getStyleWithEffectsPart)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            Dim styles As XDocument = ExtractStylesPart(fileName)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        End If
```

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
