# Copy a Worksheet Using SAX (Simple API for XML)

This topic shows how to use the the Open XML SDK for Office to programmatically copy a large worksheet
using SAX (Simple API for XML). For more information about the basic structure of a `SpreadsheetML`
document, see [Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md).

------------------------------------
## Why Use the SAX Approach?

The Open XML SDK provides two ways to parse Office Open XML files: the Document Object Model (DOM) and
the Simple API for XML (SAX). The DOM approach is designed to make it easy to query and parse Open XML
files by using strongly-typed classes. However, the DOM approach requires loading entire Open XML parts into
memory, which can lead to slower processing and `Out of Memory` exceptions when working with very large parts.
The SAX approach reads in the XML in an Open XML part one element at a time without reading in the entire part
into memory giving noncached, forward-only access to XML data, which makes it a better choice when reading
very large parts, such as a `DocumentFormat.OpenXml.Packaging.WorksheetPart` with hundreds of thousands of rows.

## Using the DOM Approach

Using the DOM approach, we can take advantage of the Open XML SDK's strongly typed classes. The first step
is to access the package's `WorksheetPart` and make sure that it is not null.

### [C#](#tab/cs-1)
```csharp
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
```

### [Visual Basic](#tab/vb-1)
```vb
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
```
***

Once it is determined that the `WorksheetPart` to be copied is not null, add a new `WorksheetPart` to copy it to.
Then clone the `WorksheetPart`'s `DocumentFormat.OpenXml.Spreadsheet.Worksheet` and assign the cloned
`Worksheet` to the new `WorksheetPart`'s Worksheet property.

### [C#](#tab/cs-2)
```csharp
            // Add a new WorksheetPart
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();

            // Make a copy of the original worksheet
            Worksheet newWorksheet = (Worksheet)worksheetPart.Worksheet.Clone();

            // Add the new worksheet to the new worksheet part
            newWorksheetPart.Worksheet = newWorksheet;
```

### [Visual Basic](#tab/vb-2)
```vb
                ' Add a new WorksheetPart
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

                ' Make a copy of the original worksheet
                Dim newWorksheet As Worksheet = CType(worksheetPart.Worksheet.Clone(), Worksheet)

                ' Add the new worksheet to the new worksheet part
                newWorksheetPart.Worksheet = newWorksheet
```
***

At this point, the new `WorksheetPart` has been added, but a new `DocumentFormat.OpenXml.Spreadsheet.Sheet`
element must be added to the  `WorkbookPart`'s `DocumentFormat.OpenXml.Spreadsheet.Sheets`'s
child elements for it to display. To do this, first find the new `WorksheetPart`'s Id and
create a new sheet Id by incrementing the `Sheets` count by one then append a new `Sheet`
child to the `Sheets` element. With this, the copied Worksheet is added to the file.

### [C#](#tab/cs-3)
```csharp
            // Find the new WorksheetPart's Id and create a new sheet id
            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            // Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
```

### [Visual Basic](#tab/vb-3)
```vb
                ' Find the new WorksheetPart's Id and create a new sheet id
                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                ' Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
```
***

## Using the SAX Approach

The SAX approach works on parts, so using the SAX approach, the first step is the same.
Access the package's `DocumentFormat.OpenXml.Packaging.WorksheetPart` and make sure
that it is not null.

### [C#](#tab/cs-4)
```csharp
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
```

### [Visual Basic](#tab/vb-4)
```vb
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
```
***

With SAX, we don't have access to the `DocumentFormat.OpenXml.OpenXmlElement.Clone`
method. So instead, start by adding a new `WorksheetPart` to the `WorkbookPart`.

### [C#](#tab/cs-5)
```csharp
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();
```

### [Visual Basic](#tab/vb-5)
```vb
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
```
***

Then create an instance of the `DocumentFormat.OpenXml.OpenXmlPartReader` with the
original worksheet part and an instance of the `DocumentFormat.OpenXml.OpenXmlPartWriter`
with the newly created worksheet part.

### [C#](#tab/cs-6)
```csharp
            using (OpenXmlReader reader = OpenXmlPartReader.Create(worksheetPart))
            using (OpenXmlWriter writer = OpenXmlPartWriter.Create(newWorksheetPart))
```

### [Visual Basic](#tab/vb-6)
```vb
                Using reader As OpenXmlReader = OpenXmlPartReader.Create(worksheetPart)
                    Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(newWorksheetPart)
```
***

Then read the elements one by one with the `DocumentFormat.OpenXml.OpenXmlPartReader.Read`
method. If the element is a `DocumentFormat.OpenXml.Spreadsheet.CellValue` the inner text
needs to be explicitly added using the `DocumentFormat.OpenXml.OpenXmlPartReader.GetText`
method to read the text, because the `DocumentFormat.OpenXml.OpenXmlPartWriter.WriteStartElement`
does not write the inner text of an element. For other elements we only need to use the `WriteStartElement`
method, because we don't need the other element's inner text.

### [C#](#tab/cs-7)
```csharp
                // Write the XML declaration with the version "1.0".
                writer.WriteStartDocument();

                // Read the elements from the original worksheet part
                while (reader.Read())
                {
                    // If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                    // or the CellValue element will be empty
                    if (reader.ElementType == typeof(CellValue))
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                            writer.WriteString(reader.GetText());
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                    // For other elements write the start and end elements
                    else
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

### [Visual Basic](#tab/vb-7)
```vb
                        ' Write the XML declaration with the version "1.0".
                        writer.WriteStartDocument()

                        ' Read the elements from the original worksheet part
                        While reader.Read()
                            ' If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                            ' or the CellValue element will be empty
                            If reader.ElementType Is GetType(CellValue) Then
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                    writer.WriteString(reader.GetText())
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                                ' For other elements write the start and end elements
                            Else
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                            End If
                        End While
```
***

At this point, the worksheet part has been copied to the newly added part, but as with the DOM
approach, we still need to add a `Sheet` to the `Workbook`'s `Sheets` element. Because
the SAX approach gives noncached, **forward-only** access to XML data, it is only possible to
prepend element children, which in this case would add the new worksheet to the beginning instead
of the end, changing the order of the worksheets. So the DOM approach is
necessary here, because we want to append not prepend the new `Sheet` and since the `WorkbookPart` is
not usually a large part, the performance gains would be minimal.

### [C#](#tab/cs-8)
```csharp
            Sheets? sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets is null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AddChild(new Sheets());
            }

            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
```

### [Visual Basic](#tab/vb-8)
```vb
                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

                If sheets Is Nothing Then
                    spreadsheetDocument.WorkbookPart.Workbook.AddChild(New Sheets())
                End If

                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
```
***

## Sample Code

Below is the sample code for both the DOM and SAX approaches to copying the data from one sheet
to a new one and adding it to the Spreadsheet document. While the DOM approach is simpler
and in many cases the preferred choice, with very large documents the SAX approach is better
given that it is faster and can prevent `Out of Memory` exceptions. To see the difference,
create a spreadsheet document with many (10,000+) rows and check the results of the
`System.Diagnostics.Stopwatch` to check the difference in execution time. Increase the
number of rows to 100,000+ to see even more significant performance gains.

### DOM Approach

### [C#](#tab/cs-0)
```csharp
void CopySheetDOM(string path)
{
    Console.WriteLine("Starting DOM method");

    Stopwatch sw = new();
    sw.Start();
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
        {
            // Add a new WorksheetPart
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();

            // Make a copy of the original worksheet
            Worksheet newWorksheet = (Worksheet)worksheetPart.Worksheet.Clone();

            // Add the new worksheet to the new worksheet part
            newWorksheetPart.Worksheet = newWorksheet;
            Sheets? sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets is null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AddChild(new Sheets());
            }
            // Find the new WorksheetPart's Id and create a new sheet id
            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            // Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
        }
    }

    sw.Stop();

    Console.WriteLine($"DOM method took {sw.Elapsed.TotalSeconds} seconds");
}
```

### [Visual Basic](#tab/vb-0)
```vb
    Sub CopySheetDOM(path As String)
        Console.WriteLine("Starting DOM method")

        Dim sw As Stopwatch = New Stopwatch()
        sw.Start()
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
                ' Add a new WorksheetPart
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

                ' Make a copy of the original worksheet
                Dim newWorksheet As Worksheet = CType(worksheetPart.Worksheet.Clone(), Worksheet)

                ' Add the new worksheet to the new worksheet part
                newWorksheetPart.Worksheet = newWorksheet
                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

                If sheets Is Nothing Then
                    spreadsheetDocument.WorkbookPart.Workbook.AddChild(New Sheets())
                End If
                ' Find the new WorksheetPart's Id and create a new sheet id
                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                ' Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
            End If
        End Using

        sw.Stop()

        Console.WriteLine($"DOM method took {sw.Elapsed.TotalSeconds} seconds")
    End Sub
```
***

### SAX Approach

### [C#](#tab/cs-99)
```csharp
void CopySheetSAX(string path)
{
    Console.WriteLine("Starting SAX method");

    Stopwatch sw = new();
    sw.Start();
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
        {
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();
            using (OpenXmlReader reader = OpenXmlPartReader.Create(worksheetPart))
            using (OpenXmlWriter writer = OpenXmlPartWriter.Create(newWorksheetPart))
            {
                // Write the XML declaration with the version "1.0".
                writer.WriteStartDocument();

                // Read the elements from the original worksheet part
                while (reader.Read())
                {
                    // If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                    // or the CellValue element will be empty
                    if (reader.ElementType == typeof(CellValue))
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                            writer.WriteString(reader.GetText());
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                    // For other elements write the start and end elements
                    else
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
            Sheets? sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets is null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AddChild(new Sheets());
            }

            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
            sw.Stop();

            Console.WriteLine($"SAX method took {sw.Elapsed.TotalSeconds} seconds");
        }
    }
}
```

### [Visual Basic](#tab/vb-99)
```vb
    Sub CopySheetSAX(path As String)
        Console.WriteLine("Starting SAX method")

        Dim sw As Stopwatch = New Stopwatch()
        sw.Start()
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
                Using reader As OpenXmlReader = OpenXmlPartReader.Create(worksheetPart)
                    Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(newWorksheetPart)
                        ' Write the XML declaration with the version "1.0".
                        writer.WriteStartDocument()

                        ' Read the elements from the original worksheet part
                        While reader.Read()
                            ' If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                            ' or the CellValue element will be empty
                            If reader.ElementType Is GetType(CellValue) Then
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                    writer.WriteString(reader.GetText())
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                                ' For other elements write the start and end elements
                            Else
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                            End If
                        End While
                    End Using
                End Using
                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

                If sheets Is Nothing Then
                    spreadsheetDocument.WorkbookPart.Workbook.AddChild(New Sheets())
                End If

                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
                sw.Stop()

                Console.WriteLine($"SAX method took {sw.Elapsed.TotalSeconds} seconds")
            End If
        End Using
    End Sub
```
***
