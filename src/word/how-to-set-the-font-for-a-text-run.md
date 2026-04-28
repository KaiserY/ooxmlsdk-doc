# Set the font for a text run

This topic shows how to use the classes in the Open XML SDK for
Office to set the font for a portion of text within a word processing
document programmatically.

--------------------------------------------------------------------------------
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

--------------------------------------------------------------------------------

## Structure of the Run Fonts Element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification can
be useful when working with `rFonts` element.

This element specifies the fonts which shall be used to display the text
contents of this run. Within a single run, there may be up to four types
of content present which shall each be allowed to use a unique font:

-   ASCII

-   High ANSI

-   Complex Script

-   East Asian

The use of each of these fonts shall be determined by the Unicode
character values of the run content, unless manually overridden via use
of the cs element.

If this element is not present, the default value is to leave the
formatting applied at previous level in the style hierarchy. If this
element is never applied in the style hierarchy, then the text shall be
displayed in any default font which supports each type of content.

Consider a single text run with both Arabic and English text, as
follows:

English العربية

This content may be expressed in a single WordprocessingML run:

```xml
    <w:r>
      <w:t>English العربية</w:t>
    </w:r>
```

Although it is in the same run, the contents are in different font faces
by specifying a different font for ASCII and CS characters in the run:

```xml
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:cs="Times New Roman" />
      </w:rPr>
      <w:t>English العربية</w:t>
    </w:r>
```

This text run shall therefore use the Courier New font for all
characters in the ASCII range, and shall use the Times New Roman font
for all characters in the Complex Script range.

&copy; ISO/IEC 29500: 2016

--------------------------------------------------------------------------------
## How the Sample Code Works

After opening the package file for read/write, the code creates a `RunProperties` object that contains a `RunFonts` object that has its `Ascii` property set to "Arial". `RunProperties` and `RunFonts` objects represent run properties
`rPr` elements and run fonts elements
`rFont`, respectively, in the Open XML
Wordprocessing schema. Use a `RunProperties`
object to specify the properties of a given text run. In this case, to
set the font of the run to Arial, the code creates a `RunFonts` object and then sets the `Ascii` value to "Arial".

### [C#](#tab/cs-1)
```csharp
        // Set the font to Arial to the first Run.
        // Use an object initializer for RunProperties and rPr.
        RunProperties rPr = new RunProperties(
            new RunFonts()
            {
                Ascii = "Arial"
            });
```
### [Visual Basic](#tab/vb-1)
```vb
            ' Set the font to Arial to the first Run.
            ' Use an object initializer for RunProperties and rPr.
            Dim rPr As New RunProperties(New RunFonts() With {
                .Ascii = "Arial"
            })
```
***

The code then creates a `DocumentFormat.OpenXml.Wordprocessing.Run` object that represents the first text
run of the document. The code instantiates a `Run` and sets it to the first text run of the
document. The code then adds the `RunProperties` object to the `Run` object using the `DocumentFormat.OpenXml.OpenXmlElement.PrependChild` method. The `PrependChild` method adds an element as the first
child element to the specified element in the in-memory XML structure.
In this case, running the code sample produces an in-memory XML
structure where the `RunProperties` element
is added as the first child element of the `Run` element. There is no need to call `Save` directly, because
we are inside of a using statement.

### [C#](#tab/cs-2)
```csharp
        if (package.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
        r.PrependChild<RunProperties>(rPr);
```
### [Visual Basic](#tab/vb-2)
```vb
            If package.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run)().First()
            r.PrependChild(Of RunProperties)(rPr)
```
***

--------------------------------------------------------------------------------

> **Note**
> This code example assumes that the test word processing document at fileName path contains at least one text run.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void SetRunFont(string fileName)
{
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
    {
        // Set the font to Arial to the first Run.
        // Use an object initializer for RunProperties and rPr.
        RunProperties rPr = new RunProperties(
            new RunFonts()
            {
                Ascii = "Arial"
            });
        if (package.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
        r.PrependChild<RunProperties>(rPr);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub SetRunFont(fileName As String)
        ' Open a Wordprocessing document for editing.
        Using package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Set the font to Arial to the first Run.
            ' Use an object initializer for RunProperties and rPr.
            Dim rPr As New RunProperties(New RunFonts() With {
                .Ascii = "Arial"
            })
            If package.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run)().First()
            r.PrependChild(Of RunProperties)(rPr)
        End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
