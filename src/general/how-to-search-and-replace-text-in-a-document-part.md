# Search and replace text in a document part

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically search and replace a text value in a word
processing document.

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

---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object

In the sample code, you start by opening the word processing file by
instantiating the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class as shown in
the following `using` statement. In the same
statement, you open the word processing file `document` by using the
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

--------------------------------------------------------------------------------
## Sample Code

The following example demonstrates a quick and easy way to search and
replace. It may not be reliable because it retrieves the XML document in
string format. Depending on the regular expression you might
unintentionally replace XML tags and corrupt the document. If you simply
want to search a document, but not replace the contents you can use
`MainDocumentPart.Document.InnerText`.

This example also shows how to use a regular expression to search and
replace the text value, "Hello World!" stored in a word processing file
with the value "Hi Everyone!". To call the method
`SearchAndReplace`, you can use the following
example.

### [C#](#tab/cs-2)
```csharp
SearchAndReplace(args[0]);
```

### [Visual Basic](#tab/vb-2)
```vb
        SearchAndReplace(args(0))
```
***

After running the program, you can inspect the file to see the change in
the text, "Hello world!"

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void SearchAndReplace(string document)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        string? docText = null;

        if (wordDoc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
        {
            docText = sr.ReadToEnd();
        }

        Regex regexText = new Regex("Hello World!");
        docText = regexText.Replace(docText, "Hi Everyone!");

        using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(docText);
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    Sub SearchAndReplace(document As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim docText As String = Nothing

            If wordDoc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Using sr As New StreamReader(wordDoc.MainDocumentPart.GetStream())
                docText = sr.ReadToEnd()
            End Using

            Dim regexText As New Regex("Hello World!")
            docText = regexText.Replace(docText, "Hi Everyone!")

            Using sw As New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))
                sw.Write(docText)
            End Using
        End Using
    End Sub
```
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)

- [Regular Expressions](https://learn.microsoft.com/dotnet/standard/base-types/regular-expressions)
