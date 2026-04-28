# Retrieve a list of the hidden worksheets in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve a list of hidden worksheets in a Microsoft Excel workbook, without loading the document into Excel. It contains an example `GetHiddenSheets` method to illustrate this task.

## GetHiddenSheets method

You can use the `GetHiddenSheets` method, to retrieve a list of the hidden worksheets in a workbook. The `GetHiddenSheets` method accepts a single parameter, a string that indicates the path of the file that you want to examine. The method works with the workbook you specify, filling a `System.Collections.Generic.List`1` instance with a reference to each hidden `Sheet` object.

## Retrieve the collection of worksheets

The `WorkbookPart` class provides a `Workbook` property, which in turn contains the XML content of the workbook. Although the Open XML SDK provides the `Sheets` property, which returns a collection of the `Sheet` parts, all the information that you need is provided by the `Sheet` elements within the `Workbook` XML content.
The following code uses the `Descendants` generic method of the `Workbook` object to retrieve a collection of `Sheet` objects that contain information about all the sheet child elements of the workbook's XML content.

### [C#](#tab/cs-4)
```csharp
        WorkbookPart? wbPart = document.WorkbookPart;

        if (wbPart is not null)
        {
            var sheets = wbPart.Workbook.Descendants<Sheet>();
```

### [Visual Basic](#tab/vb-4)
```vb
            Dim wbPart As WorkbookPart = document.WorkbookPart

            If wbPart IsNot Nothing Then
                Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()
```
***

## Retrieve hidden sheets

It's important to be aware that Excel supports two levels of worksheets. You can hide a worksheet by using the Excel user interface by right-clicking the worksheets tab and opting to hide the worksheet.
For these worksheets, the `State` property of the `Sheet` object contains an enumerated value of `Hidden`. You can also make a worksheet very hidden by writing code (either in VBA or in another language) that sets the sheet's `Visible` property to the enumerated value `xlSheetVeryHidden`. For worksheets hidden in this manner, the `State` property of the `Sheet` object contains the enumerated value `VeryHidden`.

Given the collection that contains information about all the sheets, the following code uses the `System.Linq.Enumerable.Where` function to filter the collection so that it contains only the sheets in which the `State` property is not null. If the `State` property is not null, the code looks for the `Sheet` objects in which the `State` property as a value, and where the value is either `SheetStateValues.Hidden` or `SheetStateValues.VeryHidden`.

### [C#](#tab/cs-5)
```csharp
            var hiddenSheets = sheets.Where((item) => item.State is not null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));
```

### [Visual Basic](#tab/vb-5)
```vb
                Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot Nothing AndAlso
                    item.State.HasValue AndAlso
                    (item.State.Value = SheetStateValues.Hidden OrElse
                    item.State.Value = SheetStateValues.VeryHidden))
```
***

## Sample code

The following is the complete `GetHiddenSheets` code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static List<Sheet> GetHiddenSheets(string fileName)
{
    List<Sheet> returnVal = new List<Sheet>();

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart? wbPart = document.WorkbookPart;

        if (wbPart is not null)
        {
            var sheets = wbPart.Workbook.Descendants<Sheet>();
            // Look for sheets where there is a State attribute defined, 
            // where the State has a value,
            // and where the value is either Hidden or VeryHidden.
            var hiddenSheets = sheets.Where((item) => item.State is not null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));
            returnVal = hiddenSheets.ToList();
        }
    }

    return returnVal;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function GetHiddenSheets(fileName As String) As List(Of Sheet)
        Dim returnVal As New List(Of Sheet)()

        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            Dim wbPart As WorkbookPart = document.WorkbookPart

            If wbPart IsNot Nothing Then
                Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()
                ' Look for sheets where there is a State attribute defined, 
                ' where the State has a value,
                ' and where the value is either Hidden or VeryHidden.
                Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot Nothing AndAlso
                    item.State.HasValue AndAlso
                    (item.State.Value = SheetStateValues.Hidden OrElse
                    item.State.Value = SheetStateValues.VeryHidden))
                returnVal = hiddenSheets.ToList()
            End If
        End Using

        Return returnVal
    End Function
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
