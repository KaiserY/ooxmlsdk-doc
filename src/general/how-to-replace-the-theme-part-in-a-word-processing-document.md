# Replace the theme part in a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically replace a document part in a word processing
document.

## Packages and Document Parts 

An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500](https://www.iso.org/standard/71691.html). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.

## Getting a WordprocessingDocument Object

In the sample code, you start by opening the word processing file by
instantiating the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in
the following `using` statement. In the same
statement, you open the word processing file *document* by using the
`DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` method, with the Boolean parameter set
to `true` to enable editing the document.

### [C#](#tab/cs-1)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
```

### [Visual Basic](#tab/vb-1)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
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

## How to Change Theme in a Word Package

If you would like to change the theme in a Word document, click the
ribbon **Design** and then click **Themes**. The **Themes** pull-down
menu opens. To choose one of the built-in themes and apply it to the
Word document, click the theme icon. You can also use the option **Browse for Themes...** to locate and apply a theme file
in your computer.

## The Structure of the Theme Element

The theme element is constituted of color, font, and format schemes. In
this how-to you learn how to change the theme programmatically.
Therefore, it is useful to familiarize yourself with the theme element.
The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification can
be useful when working with this element.

This element defines the root level complex type associated with a
shared style sheet (or theme). This element holds all the different
formatting options available to a document through a theme, and defines
the overall look and feel of the document when themed objects are used
within the document.

[*Example*: Consider the following image as an example of different
themes in use applied to a presentation. In this example, you can see
how a theme can affect font, colors, backgrounds, fills, and effects for
different objects in a presentation. end example]

![Theme sample](../media/a-theme01.gif)

In this example, we see how a theme can affect font, colors,
backgrounds, fills, and effects for different objects in a presentation.
*end example*]

&copy; ISO/IEC 29500: 2016

The following table lists the possible child types of the Theme class.

| PresentationML Element | Open XML SDK Class | Description |
|---|---|---|
| `<custClrLst/>` | `DocumentFormat.OpenXml.Drawing.CustomColorList` |Custom Color List |
| `<extLst/>` | `DocumentFormat.OpenXml.Presentation.ExtensionList` | Extension List |
| `<extraClrSchemeLst/>` | `DocumentFormat.OpenXml.Drawing.Theme.ExtraColorSchemeList` | Extra Color Scheme List |
| `<objectDefaults/>` | `DocumentFormat.OpenXml.Drawing.Theme.ObjectDefaults` | Object Defaults |
| `<themeElements/>` | `DocumentFormat.OpenXml.Drawing.Theme.ThemeElements` | Theme Elements |

The following XML Schema fragment defines the four parts of the theme
element. The `themeElements` element is the
piece that holds the main formatting defined within the theme. The other
parts provide overrides, defaults, and additions to the information
contained in `themeElements`. The complex
type defining a theme, `CT_OfficeStyleSheet`, is defined in the following
manner:

```xml
    <complexType name="CT_OfficeStyleSheet">
       <sequence>
           <element name="themeElements" type="CT_BaseStyles" minOccurs="1" maxOccurs="1"/>
           <element name="objectDefaults" type="CT_ObjectStyleDefaults" minOccurs="0" maxOccurs="1"/>
           <element name="extraClrSchemeLst" type="CT_ColorSchemeList" minOccurs="0" maxOccurs="1"/>
           <element name="custClrLst" type="CT_CustomColorList" minOccurs="0" maxOccurs="1"/>
           <element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
       </sequence>
       <attribute name="name" type="xsd:string" use="optional" default=""/>
    </complexType>
```

This complex type also holds a `CT_OfficeArtExtensionList`, which is used for
future extensibility of this complex type.

## How the Sample Code Works

After opening the file, you can instantiate the `MainDocumentPart` in the `wordDoc` object, and
delete the old theme part.

### [C#](#tab/cs-2)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        if (wordDoc?.MainDocumentPart?.ThemePart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.");
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart;

        // Delete the old document part.
        mainPart.DeletePart(mainPart.ThemePart);
```

### [Visual Basic](#tab/vb-2)
```vb
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            If wordDoc?.MainDocumentPart?.ThemePart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.")
            End If

            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            ' Delete the old document part.
            mainPart.DeletePart(mainPart.ThemePart)
```
***

You can then create add a new `DocumentFormat.OpenXml.Packaging.ThemePart`
object and add it to the `MainDocumentPart`
object. Then you add content by using a `StreamReader` and `System.IO.StreamWriter` objects to copy the theme from the
`themeFile` to the `DocumentFormat.OpenXml.Packaging.ThemePart`  object.

### [C#](#tab/cs-3)
```csharp
        // Add a new document part and then add content.
        ThemePart themePart = mainPart.AddNewPart<ThemePart>();

        using (StreamReader streamReader = new StreamReader(themeFile))
        using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
        {
            streamWriter.Write(streamReader.ReadToEnd());
        }
```

### [Visual Basic](#tab/vb-3)
```vb
            ' Add a new document part and then add content.
            Dim themePart As ThemePart = mainPart.AddNewPart(Of ThemePart)()

            Using streamReader As New StreamReader(themeFile)
                Using streamWriter As New StreamWriter(themePart.GetStream(FileMode.Create))
                    streamWriter.Write(streamReader.ReadToEnd())
                End Using
            End Using
```
***

## Sample Code

The following code example shows how to replace the theme document part
in a word processing document with the theme part from another package.
The theme file passed as the second argument must be a valid theme part
in XML format (for example, Theme1.xml). You can extract this part from
an existing document or theme file (.THMX) that has been renamed to be a
.Zip file. To call the method `ReplaceTheme`
you can use the following call example to copy the theme from the file
from `arg[1]` and to the file located at `arg[0]`

### [C#](#tab/cs-4)
```csharp
string document = args[0];
string themeFile = args[1];

ReplaceTheme(document, themeFile);
```

### [Visual Basic](#tab/vb-4)
```vb
        Dim document As String = args(0)
        Dim themeFile As String = args(1)

        ReplaceTheme(document, themeFile)
```
***

After you run the program open the Word file and notice the new theme changes.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
// This method can be used to replace the theme part in a package.
static void ReplaceTheme(string document, string themeFile)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        if (wordDoc?.MainDocumentPart?.ThemePart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.");
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart;

        // Delete the old document part.
        mainPart.DeletePart(mainPart.ThemePart);
        // Add a new document part and then add content.
        ThemePart themePart = mainPart.AddNewPart<ThemePart>();

        using (StreamReader streamReader = new StreamReader(themeFile))
        using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
        {
            streamWriter.Write(streamReader.ReadToEnd());
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' This method can be used to replace the theme part in a package.
    Sub ReplaceTheme(document As String, themeFile As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            If wordDoc?.MainDocumentPart?.ThemePart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.")
            End If

            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            ' Delete the old document part.
            mainPart.DeletePart(mainPart.ThemePart)
            ' Add a new document part and then add content.
            Dim themePart As ThemePart = mainPart.AddNewPart(Of ThemePart)()

            Using streamReader As New StreamReader(themeFile)
                Using streamWriter As New StreamWriter(themePart.GetStream(FileMode.Create))
                    streamWriter.Write(streamReader.ReadToEnd())
                End Using
            End Using
        End Using
    End Sub
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
