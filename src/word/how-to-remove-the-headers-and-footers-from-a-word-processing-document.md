# Remove the headers and footers from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically remove all headers and footers in a word
processing document. It contains an example `RemoveHeadersAndFooters` method to illustrate this
task.

## RemoveHeadersAndFooters Method

You can use the `RemoveHeadersAndFooters`
method to remove all header and footer information from a word
processing document. Be aware that you must not only delete the header
and footer parts from the document storage, you must also delete the
references to those parts from the document too. The sample code
demonstrates both steps in the operation. The `RemoveHeadersAndFooters` method accepts a single
parameter, a string that indicates the path of the file that you want to
modify.

### [C#](#tab/cs-0)
```csharp
static void RemoveHeadersAndFooters(string filename)
```
### [Visual Basic](#tab/vb-0)
```vb
    Sub RemoveHeadersAndFooters(filename As String)
```
***

The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document that you want to modify as shown
in the following code example.

### [C#](#tab/cs-1)
```csharp
string filename = args[0];

RemoveHeadersAndFooters(filename);
```
### [Visual Basic](#tab/vb-1)
```vb
        Dim filename As String = args(0)

        RemoveHeadersAndFooters(filename)
```
***

## How the Code Works

The `RemoveHeadersAndFooters` method works
with the document you specify, deleting all of the header and footer
parts and references to those parts. The code starts by opening the
document, using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` method and indicating that the
document should be opened for read/write access (the final true
parameter). Given the open document, the code uses the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart` property to navigate to
the main document, storing the reference in a variable named `docPart`.

### [C#](#tab/cs-2)
```csharp
    // Given a document name, remove all of the headers and footers
    // from the document.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
    {
        if (doc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Get a reference to the main document part.
        var docPart = doc.MainDocumentPart;
```
### [Visual Basic](#tab/vb-2)
```vb
        ' Given a document name, remove all of the headers and footers
        ' from the document.
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filename, True)

            If doc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            ' Get a reference to the main document part.
            Dim docPart = doc.MainDocumentPart
```
***

## Confirm Header/Footer Existence

Given a reference to the document part, the code next determines if it
has any work to do, i.e. if the document contains any headers or
footers. To decide, the code calls the `Count` method of both the `DocumentFormat.OpenXml.Packaging.MainDocumentPart.HeaderParts` and
`DocumentFormat.OpenXml.Packaging.MainDocumentPart.FooterParts` properties of the document
part, and if either returns a value greater than 0, the code continues.
Be aware that the `HeaderParts` and `FooterParts` properties each return an
`System.Collections.Generic.IEnumerable%601` of
`DocumentFormat.OpenXml.Packaging.HeaderPart` or `DocumentFormat.OpenXml.Packaging.FooterPart` objects, respectively.

### [C#](#tab/cs-3)
```csharp
        // Count the header and footer parts and continue if there 
        // are any.
        if (docPart.HeaderParts.Count() > 0 || docPart.FooterParts.Count() > 0)
        {
```
### [Visual Basic](#tab/vb-3)
```vb
            ' Count the header and footer parts and continue if there 
            ' are any.
            If docPart.HeaderParts.Count() > 0 OrElse docPart.FooterParts.Count() > 0 Then
```
***

## Remove the Header and Footer Parts

Given a collection of references to header and footer parts, you could
write code to delete each one individually, but that is not necessary
because of the Open XML SDK. Instead, you can call the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.DeleteParts%2A` method, passing in the
collection of parts to be deleted─this simple method provides a shortcut
for deleting a collection of parts. Therefore, the following few lines
of code take the place of the loop that you would otherwise have to
write yourself.

### [C#](#tab/cs-4)
```csharp
            // Remove the header and footer parts.
            docPart.DeleteParts(docPart.HeaderParts);
            docPart.DeleteParts(docPart.FooterParts);
```
### [Visual Basic](#tab/vb-4)
```vb
                ' Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts)
                docPart.DeleteParts(docPart.FooterParts)
```
***

## Delete the Header and Footer References

At this point, the code has deleted the header and footer parts, but the
document still contains orphaned references to those parts. Before the
orphaned references can be removed, the code must retrieve a reference
to the content of the document (that is, to the XML content contained
within the main document part).

To remove the stranded references, the code first retrieves a collection
of HeaderReference elements, converts the collection to a `List`, and then
loops through the collection, calling the `DocumentFormat.OpenXml.OpenXmlElement.Remove` method for each element found. Note
that the code converts the `IEnumerable`
returned by the `DocumentFormat.OpenXml.OpenXmlElement.Descendants` method into a `List` so that it
can delete items from the list, and that the `DocumentFormat.OpenXml.Wordprocessing.HeaderReference` type that is provided by
the Open XML SDK makes it easy to refer to elements of type `HeaderReference` in the XML content. (Without that
additional help, you would have to work with the details of the XML
content directly.) Once it has removed all the headers, the code repeats
the operation with the footer elements.

### [C#](#tab/cs-6)
```csharp
            // Get a reference to the root element of the main
            // document part.
            Document document = docPart.Document;

            // Remove all references to the headers and footers.

            // First, create a list of all descendants of type
            // HeaderReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var headers = document.Descendants<HeaderReference>().ToList();

            foreach (var header in headers)
            {
                header.Remove();
            }

            // First, create a list of all descendants of type
            // FooterReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var footers = document.Descendants<FooterReference>().ToList();

            foreach (var footer in footers)
            {
                footer.Remove();
            }
```
### [Visual Basic](#tab/vb-6)
```vb
                ' Get a reference to the root element of the main
                ' document part.
                Dim document As Document = docPart.Document

                ' Remove all references to the headers and footers.

                ' First, create a list of all descendants of type
                ' HeaderReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim headers = document.Descendants(Of HeaderReference)().ToList()

                For Each header In headers
                    header.Remove()
                Next

                ' First, create a list of all descendants of type
                ' FooterReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim footers = document.Descendants(Of FooterReference)().ToList()

                For Each footer In footers
                    footer.Remove()
                Next
```
***

## Sample Code

The following is the complete `RemoveHeadersAndFooters` code sample in C\# and
Visual Basic.

### [C#](#tab/cs)
```csharp
// Remove all of the headers and footers from a document.
static void RemoveHeadersAndFooters(string filename)
{
    // Given a document name, remove all of the headers and footers
    // from the document.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
    {
        if (doc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Get a reference to the main document part.
        var docPart = doc.MainDocumentPart;
        // Count the header and footer parts and continue if there 
        // are any.
        if (docPart.HeaderParts.Count() > 0 || docPart.FooterParts.Count() > 0)
        {
            // Remove the header and footer parts.
            docPart.DeleteParts(docPart.HeaderParts);
            docPart.DeleteParts(docPart.FooterParts);
            // Get a reference to the root element of the main
            // document part.
            Document document = docPart.Document;

            // Remove all references to the headers and footers.

            // First, create a list of all descendants of type
            // HeaderReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var headers = document.Descendants<HeaderReference>().ToList();

            foreach (var header in headers)
            {
                header.Remove();
            }

            // First, create a list of all descendants of type
            // FooterReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var footers = document.Descendants<FooterReference>().ToList();

            foreach (var footer in footers)
            {
                footer.Remove();
            }
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
    ' Remove all of the headers and footers from a document.
    Sub RemoveHeadersAndFooters(filename As String)
        ' Given a document name, remove all of the headers and footers
        ' from the document.
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filename, True)

            If doc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            ' Get a reference to the main document part.
            Dim docPart = doc.MainDocumentPart
            ' Count the header and footer parts and continue if there 
            ' are any.
            If docPart.HeaderParts.Count() > 0 OrElse docPart.FooterParts.Count() > 0 Then
                ' Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts)
                docPart.DeleteParts(docPart.FooterParts)
                ' Get a reference to the root element of the main
                ' document part.
                Dim document As Document = docPart.Document

                ' Remove all references to the headers and footers.

                ' First, create a list of all descendants of type
                ' HeaderReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim headers = document.Descendants(Of HeaderReference)().ToList()

                For Each header In headers
                    header.Remove()
                Next

                ' First, create a list of all descendants of type
                ' FooterReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim footers = document.Descendants(Of FooterReference)().ToList()

                For Each footer In footers
                    footer.Remove()
                Next
            End If
        End Using
    End Sub
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
