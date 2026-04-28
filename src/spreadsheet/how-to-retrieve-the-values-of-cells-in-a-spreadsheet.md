# Retrieve the values of cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve the values of cells in a spreadsheet
document. It contains an example `GetCellValue` method to illustrate
this task.

## GetCellValue Method

You can use the `GetCellValue` method to
retrieve the value of a cell in a workbook. The method requires the
following three parameters:

- A string that contains the name of the document to examine.

- A string that contains the name of the sheet to examine.

- A string that contains the cell address (such as A1, B12) from which
    to retrieve a value.

The method returns the value of the specified cell, if it could be
found. The following code example shows the method signature.

### [C#](#tab/cs-0)
```csharp
static string GetCellValue(string fileName, string sheetName, string addressName)
```

### [Visual Basic](#tab/vb-0)
```vb
    Function GetCellValue(fileName As String, sheetName As String, addressName As String) As String
```
***

## How the Code Works

The code starts by creating a variable to hold the return value, and
initializes it to null.

### [C#](#tab/cs-2)
```csharp
    string? value = null;
```

### [Visual Basic](#tab/vb-2)
```vb
        Dim value As String = Nothing
```
***

## Accessing the Cell

Next, the code opens the document by using the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open` method, indicating that the document
should be open for read-only access (the final `false` parameter). Next, the code retrieves a
reference to the workbook part by using the `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart` property of the document.

### [C#](#tab/cs-3)
```csharp
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart? wbPart = document.WorkbookPart;
```

### [Visual Basic](#tab/vb-3)
```vb
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart
```
***

To find the requested cell, the code must first retrieve a reference to
the sheet, given its name. The code must search all the sheet-type
descendants of the workbook part workbook element and examine the `DocumentFormat.OpenXml.Spreadsheet.Sheet.Name` property of each sheet that it finds.
Be aware that this search looks through the relations of the workbook,
and does not actually find a worksheet part. It finds a reference to a
`DocumentFormat.OpenXml.Spreadsheet.Sheet`, which contains information such as
the name and `DocumentFormat.OpenXml.Spreadsheet.Sheet.Id` of the sheet. The simplest way to do
this is to use a LINQ query, as shown in the following code example.

### [C#](#tab/cs-4)
```csharp
        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        // Throw an exception if there is no sheet.
        if (theSheet is null || theSheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }
```

### [Visual Basic](#tab/vb-4)
```vb
            ' Find the sheet with the supplied name, and then use that 
            ' Sheet object to retrieve a reference to the first worksheet.
            Dim theSheet As Sheet = wbPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = sheetName).FirstOrDefault()

            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing OrElse theSheet.Id Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If
```
***

Be aware that the `System.Linq.Enumerable.FirstOrDefault`
method returns either the first matching reference (a sheet, in this
case) or a null reference if no match was found. The code checks for the
null reference, and throws an exception if you passed in an invalid
sheet name.Now that you have information about the sheet, the code must
retrieve a reference to the corresponding worksheet part. The sheet
information that you already retrieved provides an `DocumentFormat.OpenXml.Spreadsheet.Sheet.Id` property, and given that **Id** property, the code can retrieve a reference to
the corresponding `DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart` by calling the workbook part
`DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.GetPartById` method.

### [C#](#tab/cs-5)
```csharp
        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
```

### [Visual Basic](#tab/vb-5)
```vb
            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
```
***

Just as when locating the named sheet, when locating the named cell, the
code uses the `DocumentFormat.OpenXml.OpenXmlElement.Descendants` method, searching for the first
match in which the `DocumentFormat.OpenXml.Spreadsheet.CellType.CellReference` property equals the specified
`addressName`
parameter. After this method call, the variable named `theCell` will either contain a reference to the cell,
or will contain a null reference.

### [C#](#tab/cs-6)
```csharp
        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
```

### [Visual Basic](#tab/vb-6)
```vb
            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet?.Descendants(Of Cell)().Where(Function(c) c.CellReference = addressName).FirstOrDefault()
```
***

## Retrieving the Value

At this point, the variable named `theCell`
contains either a null reference, or a reference to the cell that you
requested. If you examine the Open XML content (that is, `theCell.OuterXml`) for the cell, you will find XML
such as the following.

```xml
    <x:c r="A1">
        <x:v>12.345000000000001</x:v>
    </x:c>
```

The `DocumentFormat.OpenXml.OpenXmlElement.InnerText` property contains the content for
the cell, and so the next block of code retrieves this value.

### [C#](#tab/cs-7)
```csharp
        // If the cell does not exist, return an empty string.
        if (theCell is null || theCell.InnerText.Length < 0)
        {
            return string.Empty;
        }
        value = theCell.InnerText;
```

### [Visual Basic](#tab/vb-7)
```vb
            ' If the cell does not exist, return an empty string.
            If theCell Is Nothing OrElse theCell.InnerText.Length < 0 Then
                Return String.Empty
            End If
            value = theCell.InnerText
```
***

Now, the sample method must interpret the value. As it is, the code
handles numeric and date, string, and Boolean values. You can extend the
sample as necessary. The `DocumentFormat.OpenXml.Spreadsheet.Cell` type provides a
`DocumentFormat.OpenXml.Spreadsheet.CellType.DataType` property that indicates the type
of the data within the cell. The value of the `DataType` property is null for numeric and date
types. It contains the value `CellValues.SharedString` for strings, and `CellValues.Boolean` for Boolean values. If the
`DataType` property is null, the code returns
the value of the cell (it is a numeric value). Otherwise, the code
continues by branching based on the data type.

### [C#](#tab/cs-8)
```csharp
        // If the cell represents an integer number, you are done. 
        // For dates, this code returns the serialized value that 
        // represents the date. The code handles strings and 
        // Booleans individually. For shared strings, the code 
        // looks up the corresponding value in the shared string 
        // table. For Booleans, the code converts the value into 
        // the words TRUE or FALSE.
        if (theCell.DataType is not null)
        {
            if (theCell.DataType.Value == CellValues.SharedString)
            {
```

### [Visual Basic](#tab/vb-8)
```vb
            ' If the cell represents an integer number, you are done. 
            ' For dates, this code returns the serialized value that 
            ' represents the date. The code handles strings and 
            ' Booleans individually. For shared strings, the code 
            ' looks up the corresponding value in the shared string 
            ' table. For Booleans, the code converts the value into 
            ' the words TRUE or FALSE.
            If theCell.DataType IsNot Nothing Then
                If theCell.DataType.Value = CellValues.SharedString Then
```
***

If the `DataType` property contains `CellValues.SharedString`, the code must retrieve a
reference to the single `DocumentFormat.OpenXml.Packaging.WorkbookPart.SharedStringTablePart`.

### [C#](#tab/cs-9)
```csharp
                // For shared strings, look up the value in the
                // shared strings table.
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
```

### [Visual Basic](#tab/vb-9)
```vb
                    ' For shared strings, look up the value in the
                    ' shared strings table.
                    Dim stringTable = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
```
***

Next, if the string table exists (and if it does not, the workbook is
damaged and the sample code returns the index into the string table
instead of the string itself) the code returns the `InnerText` property of the element it finds at the
specified index (first converting the value property to an integer).

### [C#](#tab/cs-10)
```csharp
                // If the shared string table is missing, something 
                // is wrong. Return the index that is in
                // the cell. Otherwise, look up the correct text in 
                // the table.
                if (stringTable is not null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
```

### [Visual Basic](#tab/vb-10)
```vb
                    ' If the shared string table is missing, something 
                    ' is wrong. Return the index that is in
                    ' the cell. Otherwise, look up the correct text in 
                    ' the table.
                    If stringTable IsNot Nothing Then
                        value = stringTable.SharedStringTable.ElementAt(Integer.Parse(value)).InnerText
                    End If
```
***

If the `DataType` property contains `CellValues.Boolean`, the code converts the 0 or 1
it finds in the cell value into the appropriate text string.

### [C#](#tab/cs-11)
```csharp
                switch (value)
                {
                    case "0":
                        value = "FALSE";
                        break;
                    default:
                        value = "TRUE";
                        break;
                }
```

### [Visual Basic](#tab/vb-11)
```vb
                    Select Case value
                        Case "0"
                            value = "FALSE"
                        Case Else
                            value = "TRUE"
                    End Select
```
***

Finally, the procedure returns the variable `value`, which contains the requested information.

## Sample Code

The following is the complete `GetCellValue` code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static string GetCellValue(string fileName, string sheetName, string addressName)
{
    string? value = null;
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart? wbPart = document.WorkbookPart;
        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        // Throw an exception if there is no sheet.
        if (theSheet is null || theSheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }
        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
        // If the cell does not exist, return an empty string.
        if (theCell is null || theCell.InnerText.Length < 0)
        {
            return string.Empty;
        }
        value = theCell.InnerText;
        // If the cell represents an integer number, you are done. 
        // For dates, this code returns the serialized value that 
        // represents the date. The code handles strings and 
        // Booleans individually. For shared strings, the code 
        // looks up the corresponding value in the shared string 
        // table. For Booleans, the code converts the value into 
        // the words TRUE or FALSE.
        if (theCell.DataType is not null)
        {
            if (theCell.DataType.Value == CellValues.SharedString)
            {
                // For shared strings, look up the value in the
                // shared strings table.
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                // If the shared string table is missing, something 
                // is wrong. Return the index that is in
                // the cell. Otherwise, look up the correct text in 
                // the table.
                if (stringTable is not null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            else if (theCell.DataType.Value == CellValues.Boolean)
            {
                switch (value)
                {
                    case "0":
                        value = "FALSE";
                        break;
                    default:
                        value = "TRUE";
                        break;
                }
            }
        }
    }

    return value;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetCellValue(fileName As String, sheetName As String, addressName As String) As String
        Dim value As String = Nothing
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart
            ' Find the sheet with the supplied name, and then use that 
            ' Sheet object to retrieve a reference to the first worksheet.
            Dim theSheet As Sheet = wbPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = sheetName).FirstOrDefault()

            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing OrElse theSheet.Id Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If
            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet?.Descendants(Of Cell)().Where(Function(c) c.CellReference = addressName).FirstOrDefault()
            ' If the cell does not exist, return an empty string.
            If theCell Is Nothing OrElse theCell.InnerText.Length < 0 Then
                Return String.Empty
            End If
            value = theCell.InnerText
            ' If the cell represents an integer number, you are done. 
            ' For dates, this code returns the serialized value that 
            ' represents the date. The code handles strings and 
            ' Booleans individually. For shared strings, the code 
            ' looks up the corresponding value in the shared string 
            ' table. For Booleans, the code converts the value into 
            ' the words TRUE or FALSE.
            If theCell.DataType IsNot Nothing Then
                If theCell.DataType.Value = CellValues.SharedString Then
                    ' For shared strings, look up the value in the
                    ' shared strings table.
                    Dim stringTable = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                    ' If the shared string table is missing, something 
                    ' is wrong. Return the index that is in
                    ' the cell. Otherwise, look up the correct text in 
                    ' the table.
                    If stringTable IsNot Nothing Then
                        value = stringTable.SharedStringTable.ElementAt(Integer.Parse(value)).InnerText
                    End If
                ElseIf theCell.DataType.Value = CellValues.Boolean Then
                    Select Case value
                        Case "0"
                            value = "FALSE"
                        Case Else
                            value = "TRUE"
                    End Select
                End If
            End If
        End Using

        Return value
    End Function
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
