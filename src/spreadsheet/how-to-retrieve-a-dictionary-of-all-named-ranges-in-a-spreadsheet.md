# Retrieve a dictionary of all named ranges in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a dictionary that contains the names
and ranges of all defined names in a Microsoft Excel workbook. It contains an example **GetDefinedNames** method
to illustrate this task.

## GetDefinedNames Method

The **GetDefinedNames** method accepts a
single parameter that indicates the name of the document from which to
retrieve the defined names. The method returns an
`System.Collections.Generic.Dictionary`2`
instance that contains information about the defined names within the
specified workbook, which may be empty if there are no defined names.

## How the Code Works

The code opens the spreadsheet document, using the **Open** method, indicating that the
document should be open for read-only access with the final false parameter. Given the open workbook, the code uses the **WorkbookPart** property to navigate to the main workbook part. The code stores this reference in a variable named **wbPart**.

### [C#](#tab/cs-3)
```csharp
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        var wbPart = document.WorkbookPart;
```

### [Visual Basic](#tab/vb-3)
```vb
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart = document.WorkbookPart
```
***

## Retrieving the Defined Names

Given the workbook part, the next step is simple. The code uses the
**Workbook** property of the workbook part to retrieve a reference to the content of the workbook, and then retrieves the **DefinedNames** collection provided by the Open XML SDK. This property returns a collection of all of the
defined names that are contained within the workbook. If the property returns a non-null value, the code then iterates through the collection, retrieving information about each named part and adding the key  name) and value (range description) to the dictionary for each defined name.

### [C#](#tab/cs-4)
```csharp
        // Retrieve a reference to the defined names collection.
        DefinedNames? definedNames = wbPart?.Workbook?.DefinedNames;

        // If there are defined names, add them to the dictionary.
        if (definedNames is not null)
        {
            foreach (DefinedName dn in definedNames)
            {
                if (dn?.Name?.Value is not null && dn?.Text is not null)
                {
                    returnValue.Add(dn.Name.Value, dn.Text);
                }
            }
        }
```

### [Visual Basic](#tab/vb-4)
```vb
            ' Retrieve a reference to the defined names collection.
            Dim definedNames As DefinedNames = wbPart?.Workbook?.DefinedNames

            ' If there are defined names, add them to the dictionary.
            If definedNames IsNot Nothing Then
                For Each dn As DefinedName In definedNames
                    If dn?.Name?.Value IsNot Nothing AndAlso dn?.Text IsNot Nothing Then
                        returnValue.Add(dn.Name.Value, dn.Text)
                    End If
                Next
            End If
```
***

## Sample Code

The following is the complete **GetDefinedNames** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static Dictionary<String, String>GetDefinedNames(String fileName)
{
    // Given a workbook name, return a dictionary of defined names.
    // The pairs include the range name and a string representing the range.
    var returnValue = new Dictionary<String, String>();
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        var wbPart = document.WorkbookPart;
        // Retrieve a reference to the defined names collection.
        DefinedNames? definedNames = wbPart?.Workbook?.DefinedNames;

        // If there are defined names, add them to the dictionary.
        if (definedNames is not null)
        {
            foreach (DefinedName dn in definedNames)
            {
                if (dn?.Name?.Value is not null && dn?.Text is not null)
                {
                    returnValue.Add(dn.Name.Value, dn.Text);
                }
            }
        }
    }

    return returnValue;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetDefinedNames(fileName As String) As Dictionary(Of String, String)
        ' Given a workbook name, return a dictionary of defined names.
        ' The pairs include the range name and a string representing the range.
        Dim returnValue As New Dictionary(Of String, String)()
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart = document.WorkbookPart
            ' Retrieve a reference to the defined names collection.
            Dim definedNames As DefinedNames = wbPart?.Workbook?.DefinedNames

            ' If there are defined names, add them to the dictionary.
            If definedNames IsNot Nothing Then
                For Each dn As DefinedName In definedNames
                    If dn?.Name?.Value IsNot Nothing AndAlso dn?.Text IsNot Nothing Then
                        returnValue.Add(dn.Name.Value, dn.Text)
                    End If
                Next
            End If
        End Using

        Return returnValue
    End Function
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
