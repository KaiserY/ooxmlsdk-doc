# Remove hidden text from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically remove hidden text from a word processing
document.

--------------------------------------------------------------------------------

## Structure of a WordProcessingML Document

The basic document structure of a `WordProcessingML` document consists of the `document` and `body` elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph contains one or more `r` elements. The `r` stands for run, which is a region of text with a common set of properties, such as formatting. A run contains one or more `t` elements. The `t` element contains a range of text. The following code example shows the `WordprocessingML` markup for a document that contains the text "Example text."

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK, you can create document structure and content using strongly-typed classes that correspond to `WordprocessingML` elements. You will find these classes in the `DocumentFormat.OpenXml.Wordprocessing` namespace. The following table lists the class names of the classes that correspond to the `document`, `body`, `p`, `r`, and `t` elements.

| **WordprocessingML Element** | **Open XML SDK Class** | **Description** |
|---|---|---|
| `<document/>` | `DocumentFormat.OpenXml.Wordprocessing.Document` | The root element for the main document part. |
| `<body/>` | `DocumentFormat.OpenXml.Wordprocessing.Body` | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| `<p/>` | `DocumentFormat.OpenXml.Wordprocessing.Paragraph` | A paragraph. |
| `<r/>` | `DocumentFormat.OpenXml.Wordprocessing.Run` | A run. |
| `<t/>` | `DocumentFormat.OpenXml.Wordprocessing.Text` | A range of text. |

For more information about the overall structure of the parts and elements of a WordprocessingML document, see [Structure of a WordprocessingML document](structure-of-a-wordprocessingml-document.md).

---------------------------------------------------------------------------------
## Structure of the Vanish Element

The `vanish` element plays an important role in hiding the text in a
Word file. The `Hidden` formatting property is a toggle property,
which means that its behavior differs between using it within a style
definition and using it as direct formatting. When used as part of a
style definition, setting this property toggles its current state.
Setting it to `false` (or an equivalent)
results in keeping the current setting unchanged. However, when used as
direct formatting, setting it to `true` or
`false` sets the absolute state of the
resulting property.

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the `vanish` element.

> **vanish (Hidden Text)**
> 
> This element specifies whether the contents of this run shall be
> hidden from display at display time in a document. [*Note*: The
> setting should affect the normal display of text, but an application
> can have settings to force hidden text to be displayed. *end note*]
> 
> This formatting property is a *toggle property* (§17.7.3).
> 
> If this element is not present, the default value is to leave the
> formatting applied at previous level in the *style hierarchy*. If this
> element is never applied in the style hierarchy, then this text shall
> not be hidden when displayed in a document.
> 
> [*Example*: Consider a run of text which shall have the hidden text
> property turned on for the contents of the run. This constraint is
> specified using the following WordprocessingML:

```xml
    <w:rPr>
      <w:vanish />
    </w:rPr>
```

> This run declares that the **vanish** property is set for the contents
> of this run, so the contents of this run will be hidden when the
> document contents are displayed. *end example*]
> 
> © ISO/IEC 29500: 2016

The following XML schema segment defines the contents of this element.

```xml
    <complexType name="CT_OnOff">
       <attribute name="val" type="ST_OnOff"/>
    </complexType>
```

The `val` property in the code above is a binary value that can be
turned on or off. If given a value of `on`, `1`, or `true` the property is turned on. If given the
value `off`, `0`, or `false` the property
is turned off.

## How the Code Works

The `WDDeleteHiddenText` method works with the document you specify and removes all of the `run` elements that are hidden and removes extra `vanish` elements. The code starts by opening the
document, using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` method and indicating that the
document should be opened for read/write access (the final true
parameter). Given the open document, the code uses the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart` property to navigate to
the main document, storing the reference in a variable.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument doc = WordprocessingDocument.Open(docName, true))
    {
```
### [Visual Basic](#tab/vb-0)
```vb
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
```
***

## Get a List of Vanish Elements

The code first checks that `doc.MainDocumentPart` and `doc.MainDocumentPart.Document.Body` are not null and throws an exception if one is missing. Then uses the `DocumentFormat.OpenXml.OpenXmlElement.Descendants` passing it the `DocumentFormat.OpenXml.Wordprocessing.Vanish` type to get an `IEnumerable` of the `Vanish` elements and casts them to a list.

### [C#](#tab/cs-1)
```csharp
        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Get a list of all the Vanish elements
        List<Vanish> vanishes = doc.MainDocumentPart.Document.Body.Descendants<Vanish>().ToList();
```
### [Visual Basic](#tab/vb-1)
```vb
            If doc.MainDocumentPart Is Nothing Or doc.MainDocumentPart.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is Nothing.")
            End If

            'Get a list of all the Vanish elements
            Dim vanishes As List(Of Vanish) = doc.MainDocumentPart.Document.Body.Descendants(Of Vanish).ToList()
```
***

## Remove Runs with Hidden Text and Extra Vanish Elements

To remove the hidden text we next loop over the `List` of `Vanish` elements. The `Vanish` element is a child of the `DocumentFormat.OpenXml.Wordprocessing.RunProperties` but `RunProperties` can be a child of a `DocumentFormat.OpenXml.Wordprocessing.Run` or `DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties`, so we get the parent and grandparent of each `Vanish` and check its type. Then if the grandparent is a `Run` we remove that run and if not 
we we remove the `Vanish` child elements from the parent.

### [C#](#tab/cs-2)
```csharp
        // Loop over the list of Vanish elements
        foreach (Vanish vanish in vanishes)
        {
            var parent = vanish?.Parent;
            var grandparent = parent?.Parent;

            // If the grandparent is a Run remove it
            if (grandparent is Run)
            {
                grandparent.Remove();
            }
            // If it's not a run remove the Vanish
            else if (parent is not null)
            {
                parent.RemoveAllChildren<Vanish>();
            }
        }
```
### [Visual Basic](#tab/vb-2)
```vb
            ' Loop over the list of Vanish elements
            For Each vanish In vanishes
                Dim parent = vanish.Parent
                Dim grandparent = parent.Parent

                ' If the grandparent is a Run remove it
                If TypeOf grandparent Is Run Then
                    grandparent.Remove()

                    ' If it's not a run remove the Vanish
                ElseIf parent IsNot Nothing Then
                    parent.RemoveAllChildren(Of Vanish)()
                End If
            Next
```
***
--------------------------------------------------------------------------------
## Sample Code

> **Note**
> This example assumes that the file being opened contains some hidden text. In order to hide part of the file text, select it, and click CTRL+D to show the **Font** dialog box. Select the **Hidden** box and click **OK**.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs-3)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

static void WDDeleteHiddenText(string docName)
{
    // Given a document name, delete all the hidden text.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(docName, true))
    {
        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Get a list of all the Vanish elements
        List<Vanish> vanishes = doc.MainDocumentPart.Document.Body.Descendants<Vanish>().ToList();
        // Loop over the list of Vanish elements
        foreach (Vanish vanish in vanishes)
        {
            var parent = vanish?.Parent;
            var grandparent = parent?.Parent;

            // If the grandparent is a Run remove it
            if (grandparent is Run)
            {
                grandparent.Remove();
            }
            // If it's not a run remove the Vanish
            else if (parent is not null)
            {
                parent.RemoveAllChildren<Vanish>();
            }
        }
    }
}
```

### [Visual Basic](#tab/vb-3)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        Dim fileName As String = args(0)

        WDDeleteHiddenText(fileName)
    End Sub

    Public Sub WDDeleteHiddenText(ByVal fileName As String)
        ' Given a document name, delete all the hidden text.
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If doc.MainDocumentPart Is Nothing Or doc.MainDocumentPart.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is Nothing.")
            End If

            'Get a list of all the Vanish elements
            Dim vanishes As List(Of Vanish) = doc.MainDocumentPart.Document.Body.Descendants(Of Vanish).ToList()
            ' Loop over the list of Vanish elements
            For Each vanish In vanishes
                Dim parent = vanish.Parent
                Dim grandparent = parent.Parent

                ' If the grandparent is a Run remove it
                If TypeOf grandparent Is Run Then
                    grandparent.Remove()

                    ' If it's not a run remove the Vanish
                ElseIf parent IsNot Nothing Then
                    parent.RemoveAllChildren(Of Vanish)()
                End If
            Next
        End Using
    End Sub
End Module
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
