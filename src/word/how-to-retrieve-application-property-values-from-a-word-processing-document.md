# Retrieve application property values from a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve an application property from a Microsoft Word document, without loading the document into Word. It contains example code to illustrate this task.

## Retrieving Application Properties

To retrieve application document properties, you can retrieve the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart` property of a `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` object, and then retrieve the specific application property you need. To do this, you must first get a reference to the document, as shown in the following code.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
    {
```
### [Visual Basic](#tab/vb-0)
```vb
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)
```
***

Given the reference to the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` object, you can retrieve a reference to the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart` property of the document. This object provides its own properties, each of which exposes one of the application document properties.

### [C#](#tab/cs-1)
```csharp
        if (document.ExtendedFilePropertiesPart is null)
        {
            throw new ArgumentNullException("ExtendedFilePropertiesPart is null.");
        }

        var props = document.ExtendedFilePropertiesPart.Properties;
```
### [Visual Basic](#tab/vb-1)
```vb
            If document.ExtendedFilePropertiesPart Is Nothing Then
                Throw New ArgumentNullException("ExtendedFileProperties is Nothing")
            End If

            Dim props = document.ExtendedFilePropertiesPart.Properties
```
***

Once you have the reference to the properties of `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart`, you can then retrieve any of the application properties, using simple code such as that shown
in the next example. Note that the code must confirm that the reference to each property isn't `null` of `Nothing` before retrieving its `Text` property. Unlike core properties, document properties aren't available if you (or the application) haven't specifically given them a value.

### [C#](#tab/cs-2)
```csharp
        if (props.Company is not null)
            Console.WriteLine("Company = " + props.Company.Text);

        if (props.Lines is not null)
            Console.WriteLine("Lines = " + props.Lines.Text);

        if (props.Manager is not null)
            Console.WriteLine("Manager = " + props.Manager.Text);
```
### [Visual Basic](#tab/vb-2)
```vb
            If props.Company IsNot Nothing Then
                Console.WriteLine("Company = " & props.Company.Text)
            End If

            If props.Lines IsNot Nothing Then
                Console.WriteLine("Lines = " & props.Lines.Text)
            End If

            If props.Manager IsNot Nothing Then
                Console.WriteLine("Manager = " & props.Manager.Text)
            End If
```
***

## Sample Code

The following is the complete code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using System;
static void GetApplicationProperty(string fileName)
{
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
    {
        if (document.ExtendedFilePropertiesPart is null)
        {
            throw new ArgumentNullException("ExtendedFilePropertiesPart is null.");
        }

        var props = document.ExtendedFilePropertiesPart.Properties;
        if (props.Company is not null)
            Console.WriteLine("Company = " + props.Company.Text);

        if (props.Lines is not null)
            Console.WriteLine("Lines = " + props.Lines.Text);

        if (props.Manager is not null)
            Console.WriteLine("Manager = " + props.Manager.Text);
    }
}
```

### [Visual Basic](#tab/vb)
```vb
Imports System.Runtime.Serialization
Imports DocumentFormat.OpenXml.Packaging

Module Module1

    Sub Main(args As String())
        GetPropertyValues(args(0))
    End Sub

    Public Sub GetPropertyValues(ByVal fileName As String)
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)
            If document.ExtendedFilePropertiesPart Is Nothing Then
                Throw New ArgumentNullException("ExtendedFileProperties is Nothing")
            End If

            Dim props = document.ExtendedFilePropertiesPart.Properties
            If props.Company IsNot Nothing Then
                Console.WriteLine("Company = " & props.Company.Text)
            End If

            If props.Lines IsNot Nothing Then
                Console.WriteLine("Lines = " & props.Lines.Text)
            End If

            If props.Manager IsNot Nothing Then
                Console.WriteLine("Manager = " & props.Manager.Text)
            End If
        End Using
    End Sub
End Module
```

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
