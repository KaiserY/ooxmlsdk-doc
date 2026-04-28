# Replace Text in a Word Document Using SAX (Simple API for XML)

This topic shows how to use the Open XML SDK to search and replace text in a Word document with the
Open XML SDK using the Simple API for XML (SAX) approach. For more information about the basic structure
of a `WordprocessingML` document, see [Structure of a WordprocessingML document](./structure-of-a-wordprocessingml-document.md).

## Why Use the SAX Approach?

The Open XML SDK provides two ways to parse Office Open XML files: the Document Object Model (DOM) and the Simple API for XML (SAX). The DOM approach is designed to make it easy to query and parse Open XML files by using strongly-typed classes. However, the DOM approach requires loading entire Open XML parts into memory, which can lead to slower processing and Out of Memory exceptions when working with very large parts. The SAX approach reads in the XML in an Open XML part one element at a time without reading in the entire part into memory giving noncached, forward-only access to the XML data, which makes it a better choice when reading very large parts.

## Accessing the MainDocumentPart

The text of a Word document is stored in the `DocumentFormat.OpenXml.Packaging.MainDocumentPart`, so the first step to
finding and replacing text is to access the Word document's `MainDocumentPart`. To do that we first use the `WordprocessingDocument.Open`
method passing in the path to the document as the first parameter and a second parameter `true` to indicate that we
are opening the file for editing. Then make sure that the `MainDocumentPart` is not null.

### [C#](#tab/cs-1)
```csharp
    // Open the WordprocessingDocument for editing
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, true))
    {
        // Access the MainDocumentPart and make sure it is not null
        MainDocumentPart? mainDocumentPart = wordprocessingDocument.MainDocumentPart;

        if (mainDocumentPart is not null)
```

### [Visual Basic](#tab/vb-1)
```vb
        ' Open the WordprocessingDocument for editing
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(path, True)
            ' Access the MainDocumentPart and make sure it is not null
            Dim mainDocumentPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            If mainDocumentPart IsNot Nothing Then
```
***

## Create Memory Stream, OpenXmlReader, and OpenXmlWriter

With the DOM approach to editing documents, the entire part is read into memory, so we can use the Open XML SDK's
strongly typed classes to access the `DocumentFormat.OpenXml.Wordprocessing.Text` class to access the
document's text and edit it. The SAX approach, however, uses the `DocumentFormat.OpenXml.OpenXmlPartReader`
and `DocumentFormat.OpenXml.OpenXmlPartWriter` classes, which access a part's stream with forward-only
access. The advantage of this is that the entire part does not need to be loaded into memory, which is faster
and uses less memory, but since the same part cannot be opened in multiple streams at the same time, we cannot create a
`DocumentFormat.OpenXml.OpenXmlReader` to read a part and a `DocumentFormat.OpenXml.OpenXmlWriter` to edit
the same part at the same time. The solution to this is to create an additional memory stream and write the
updated part to the new memory stream then use the stream to update the part when `OpenXmlReader` and `OpenXmlWriter`
have been disposed. In the code below we create the `MemoryStream` to store the updated part and create an
`OpenXmlReader` for the `MainDocumentPart` and a `OpenXmlWriter` to write to the `MemoryStream`

### [C#](#tab/cs-2)
```csharp
            // Create a MemoryStream to store the updated MainDocumentPart
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Create an OpenXmlReader to read the main document part
                // and an OpenXmlWriter to write to the MemoryStream
                using (OpenXmlReader reader = OpenXmlPartReader.Create(mainDocumentPart))
                using (OpenXmlWriter writer = OpenXmlPartWriter.Create(memoryStream))
```

### [Visual Basic](#tab/vb-2)
```vb
                ' Create a MemoryStream to store the updated MainDocumentPart
                Using memoryStream As New MemoryStream()
                    ' Create an OpenXmlReader to read the main document part
                    ' and an OpenXmlWriter to write to the MemoryStream
                    Using reader As OpenXmlReader = OpenXmlPartReader.Create(mainDocumentPart)
                        Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(memoryStream)
```
***

## Reading the Part and Writing to the New Stream

Now that we have an `OpenXmlReader` to read the part and an `OpenXmlWriter` to write to the new `MemoryStream`
we use the `DocumentFormat.OpenXml.OpenXmlReader.Read` method to read each element in the part. As
each element is read in we check if it is of type `Text` and if it is, we use the <xrefDocumentFormat.OpenXml.OpenXmlReader.GetText*>
method to access the text and use `System.String.Replace` to update the text. If it is not a
`Text` element, then we write it to the stream unchanged.

> [!Note]
> In a Word document text can be separated into multiple `Text` elements, so if you are replacing a
> phrase and not a single word, it's best to replace one word at a time.

### [C#](#tab/cs-3)
```csharp
                    // Write the XML declaration with the version "1.0".
                    writer.WriteStartDocument();
                    
                    // Read the elements from the MainDocumentPart
                    while (reader.Read())
                    {
                        // Check if the element is of type Text
                        if (reader.ElementType == typeof(Text))
                        {
                            // If it is the start of an element write the start element and the updated text
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);

                                string text = reader.GetText().Replace(textToReplace, replacementText);

                                writer.WriteString(text);

                            }
                            else
                            {
                                // Close the element
                                writer.WriteEndElement();
                            }
                        }
                        else
                        // Write the other XML elements without editing
                        {
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);
                            }
                            else if (reader.IsEndElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                    }
```

### [Visual Basic](#tab/vb-3)
```vb
                            ' Write the XML declaration with the version "1.0".
                            writer.WriteStartDocument()

                            ' Read the elements from the MainDocumentPart
                            While reader.Read()
                                ' Check if the element is of type Text
                                If reader.ElementType Is GetType(Text) Then
                                    ' If it is the start of an element write the start element and the updated text
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)

                                        Dim text As String = reader.GetText().Replace(textToReplace, replacementText)

                                        writer.WriteString(text)
                                    Else
                                        ' Close the element
                                        writer.WriteEndElement()
                                    End If
                                Else
                                    ' Write the other XML elements without editing
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)
                                    ElseIf reader.IsEndElement Then
                                        writer.WriteEndElement()
                                    End If
                                End If
                            End While
```
***

## Writing the New Stream to the MainDocumentPart

With the updated part written to the memory stream the last step is to set the `MemoryStream`'s
position to 0 and use the `DocumentFormat.OpenXml.Packaging.OpenXmlPart.FeedData` method
to replace the `MainDocumentPart` with the updated stream.

### [C#](#tab/cs-4)
```csharp
                // Set the MemoryStream's position to 0 and replace the MainDocumentPart
                memoryStream.Position = 0;
                mainDocumentPart.FeedData(memoryStream);
```

### [Visual Basic](#tab/vb-4)
```vb
                    ' Set the MemoryStream's position to 0 and replace the MainDocumentPart
                    memoryStream.Position = 0
                    mainDocumentPart.FeedData(memoryStream)
```
***

## Sample Code

Below is the complete sample code to replace text in a Word document using the SAX (Simple API for XML)
approach.

### [C#](#tab/cs-0)
```csharp
void ReplaceTextWithSAX(string path, string textToReplace, string replacementText)
{
    // Open the WordprocessingDocument for editing
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, true))
    {
        // Access the MainDocumentPart and make sure it is not null
        MainDocumentPart? mainDocumentPart = wordprocessingDocument.MainDocumentPart;

        if (mainDocumentPart is not null)
        {
            // Create a MemoryStream to store the updated MainDocumentPart
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Create an OpenXmlReader to read the main document part
                // and an OpenXmlWriter to write to the MemoryStream
                using (OpenXmlReader reader = OpenXmlPartReader.Create(mainDocumentPart))
                using (OpenXmlWriter writer = OpenXmlPartWriter.Create(memoryStream))
                {
                    // Write the XML declaration with the version "1.0".
                    writer.WriteStartDocument();
                    
                    // Read the elements from the MainDocumentPart
                    while (reader.Read())
                    {
                        // Check if the element is of type Text
                        if (reader.ElementType == typeof(Text))
                        {
                            // If it is the start of an element write the start element and the updated text
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);

                                string text = reader.GetText().Replace(textToReplace, replacementText);

                                writer.WriteString(text);

                            }
                            else
                            {
                                // Close the element
                                writer.WriteEndElement();
                            }
                        }
                        else
                        // Write the other XML elements without editing
                        {
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);
                            }
                            else if (reader.IsEndElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                    }
                }
                // Set the MemoryStream's position to 0 and replace the MainDocumentPart
                memoryStream.Position = 0;
                mainDocumentPart.FeedData(memoryStream);
            }
        }
    }
}
```

### [Visual Basic](#tab/vb-0)
```vb
    Sub ReplaceTextWithSAX(path As String, textToReplace As String, replacementText As String)
        ' Open the WordprocessingDocument for editing
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(path, True)
            ' Access the MainDocumentPart and make sure it is not null
            Dim mainDocumentPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            If mainDocumentPart IsNot Nothing Then
                ' Create a MemoryStream to store the updated MainDocumentPart
                Using memoryStream As New MemoryStream()
                    ' Create an OpenXmlReader to read the main document part
                    ' and an OpenXmlWriter to write to the MemoryStream
                    Using reader As OpenXmlReader = OpenXmlPartReader.Create(mainDocumentPart)
                        Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(memoryStream)
                            ' Write the XML declaration with the version "1.0".
                            writer.WriteStartDocument()

                            ' Read the elements from the MainDocumentPart
                            While reader.Read()
                                ' Check if the element is of type Text
                                If reader.ElementType Is GetType(Text) Then
                                    ' If it is the start of an element write the start element and the updated text
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)

                                        Dim text As String = reader.GetText().Replace(textToReplace, replacementText)

                                        writer.WriteString(text)
                                    Else
                                        ' Close the element
                                        writer.WriteEndElement()
                                    End If
                                Else
                                    ' Write the other XML elements without editing
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)
                                    ElseIf reader.IsEndElement Then
                                        writer.WriteEndElement()
                                    End If
                                End If
                            End While
                        End Using
                    End Using
                    ' Set the MemoryStream's position to 0 and replace the MainDocumentPart
                    memoryStream.Position = 0
                    mainDocumentPart.FeedData(memoryStream)
                End Using
            End If
        End Using
    End Sub
```
***
