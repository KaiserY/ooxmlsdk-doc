# Set a custom property in a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically set a custom property in a word processing document. It contains an example  `SetCustomProperty` method to illustrate this task.

The sample code also includes an enumeration that defines the possible types of custom properties. The `SetCustomProperty` method requires that you supply one of these values when you call the method.

### [C#](#tab/cs-0)
```csharp
enum PropertyTypes : int
{
    YesNo,
    Text,
    DateTime,
    NumberInteger,
    NumberDouble
}
```
### [Visual Basic](#tab/vb-0)
```vb
    Enum PropertyTypes As Integer
        YesNo
        Text
        [DateTime]
        NumberInteger
        NumberDouble
    End Enum
```
***

## How Custom Properties Are Stored

It is important to understand how custom properties are stored in a word
processing document. You can use the Productivity Tool for Microsoft
Office, shown in Figure 1, to discover how they are stored. This tool
enables you to open a document and view its parts and the hierarchy of
parts. Figure 1 shows a test document after you run the code in the
[Calling the SetCustomProperty Method](#calling-the-setcustomproperty-method) section of
this article. The tool displays in the right-hand panes both the XML for
the part and the reflected C\# code that you can use to generate the
contents of the part.

Figure 1. Open XML SDK Productivity Tool for Microsoft Office

 ![Open XML SDK Productivity Tool](../media/OpenXmlCon_HowToSetCustomProperty_Fig1.gif)
  
The relevant XML is also extracted and shown here for ease of reading.

```xml
    <op:Properties xmlns:vt="https://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns:op="https://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
      <op:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Manager">
        <vt:lpwstr>Mary</vt:lpwstr>
      </op:property>
      <op:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ReviewDate">
        <vt:filetime>2010-12-21T00:00:00Z</vt:filetime>
      </op:property>
    </op:Properties>
```

If you examine the XML content, you will find the following:

- Each property in the XML content consists of an XML element that includes the name and the value of the property.
- For each property, the XML content includes an `fmtid` attribute, which is always set to the same string value: `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`.
- Each property in the XML content includes a `pid` attribute, which must include an integer starting at 2 for the first property and incrementing for each successive property.
- Each property tracks its type (in the figure, the `vt:lpwstr` and `vt:filetime` element names define the types for each property).

The sample method that is provided here includes the code that is required to create or modify a custom document property in a Microsoft Word document. You can find the complete code listing for the method in the [Sample Code](#sample-code) section.

## SetCustomProperty Method

Use the `SetCustomProperty` method to set a custom property in a word processing document. The `SetCustomProperty` method accepts four parameters:

- The name of the document to modify (string).

- The name of the property to add or modify (string).

- The value of the property (object).

- The kind of property (one of the values in the `PropertyTypes` enumeration).

### [C#](#tab/cs-1)
```csharp
static string SetCustomProperty(
    string fileName,
    string propertyName,
    object propertyValue,
    PropertyTypes propertyType)
```
### [Visual Basic](#tab/vb-1)
```vb
    Function SetCustomProperty(fileName As String, propertyName As String, propertyValue As Object, propertyType As PropertyTypes) As String
```
***

## Calling the SetCustomProperty Method

The `SetCustomProperty` method enables you to set a custom property, and returns the current value of the property, if it exists. To call the sample method, pass the file name, property name, property value, and property type parameters. The following sample code shows an example.

### [C#](#tab/cs-2)
```csharp
string fileName = args[0];

Console.WriteLine(string.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Pedro", PropertyTypes.Text)));

Console.WriteLine(string.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Bonnie", PropertyTypes.Text)));

Console.WriteLine(string.Join("ReviewDate = ", SetCustomProperty(fileName, "ReviewDate", DateTime.Parse("01/26/2024"), PropertyTypes.DateTime)));
```
### [Visual Basic](#tab/vb-2)
```vb
    Sub Main(args As String())
        Dim fileName As String = args(0)

        Console.WriteLine(String.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Pedro", PropertyTypes.Text)))

        Console.WriteLine(String.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Shweta", PropertyTypes.Text)))

        Console.WriteLine(String.Join("ReviewDate = ", SetCustomProperty(fileName, "ReviewDate", DateTime.Parse("01/26/2024"), PropertyTypes.DateTime)))
    End Sub
```
***

After running this code, use the following procedure to view the custom properties from Word.

1. Open the .docx file in Word.
2. On the **File** tab, click **Info**.
3. Click **Properties**.
4. Click **Advanced Properties**.

The custom properties will display in the dialog box that appears, as shown in Figure 2.

Figure 2. Custom Properties in the Advanced Properties dialog box

 ![Advanced Properties dialog with custom properties](../media/custom-property-menu.png)

## How the Code Works

The `SetCustomProperty` method starts by setting up some internal variables. Next, it examines the information about the property, and creates a new `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty` based on the parameters that you have specified. The code also maintains a variable named `propSet` to indicate whether it successfully created the new property object. This code verifies the
type of the property value, and then converts the input to the correct type, setting the appropriate property of the `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty` object.

> **Note**
> The `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty` type works much like a VBA Variant type. It maintains separate placeholders as properties for the various types of data it might contain.

### [C#](#tab/cs-3)
```csharp
    string? returnValue = string.Empty;

    var newProp = new CustomDocumentProperty();
    bool propSet = false;

    string? propertyValueString = propertyValue.ToString() ?? throw new System.ArgumentNullException("propertyValue can't be converted to a string.");

    // Calculate the correct type.
    switch (propertyType)
    {
        case PropertyTypes.DateTime:

            // Be sure you were passed a real date, 
            // and if so, format in the correct way. 
            // The date/time value passed in should 
            // represent a UTC date/time.
            if ((propertyValue) is DateTime)
            {
                newProp.VTFileTime =
                    new VTFileTime(string.Format("{0:s}Z",
                        Convert.ToDateTime(propertyValue)));
                propSet = true;
            }

            break;

        case PropertyTypes.NumberInteger:
            if ((propertyValue) is int)
            {
                newProp.VTInt32 = new VTInt32(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.NumberDouble:
            if (propertyValue is double)
            {
                newProp.VTFloat = new VTFloat(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.Text:
            newProp.VTLPWSTR = new VTLPWSTR(propertyValueString);
            propSet = true;

            break;

        case PropertyTypes.YesNo:
            if (propertyValue is bool)
            {
                // Must be lowercase.
                newProp.VTBool = new VTBool(
                  Convert.ToBoolean(propertyValue).ToString().ToLower());
                propSet = true;
            }
            break;
    }

    if (!propSet)
    {
        // If the code was not able to convert the 
        // property to a valid value, throw an exception.
        throw new InvalidDataException("propertyValue");
    }
```
### [Visual Basic](#tab/vb-3)
```vb
        Dim returnValue As String = String.Empty

        Dim newProp As New CustomDocumentProperty()
        Dim propSet As Boolean = False

        Dim propertyValueString As String = propertyValue.ToString()
        If propertyValueString Is Nothing Then
            Throw New ArgumentNullException("propertyValue can't be converted to a string.")
        End If

        ' Calculate the correct type.
        Select Case propertyType
            Case PropertyTypes.DateTime
                ' Be sure you were passed a real date, 
                ' and if so, format in the correct way. 
                ' The date/time value passed in should 
                ' represent a UTC date/time.
                If TypeOf propertyValue Is DateTime Then
                    newProp.VTFileTime = New VTFileTime(String.Format("{0:s}Z", Convert.ToDateTime(propertyValue)))
                    propSet = True
                End If

            Case PropertyTypes.NumberInteger
                If TypeOf propertyValue Is Integer Then
                    newProp.VTInt32 = New VTInt32(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.NumberDouble
                If TypeOf propertyValue Is Double Then
                    newProp.VTFloat = New VTFloat(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.Text
                newProp.VTLPWSTR = New VTLPWSTR(propertyValueString)
                propSet = True

            Case PropertyTypes.YesNo
                If TypeOf propertyValue Is Boolean Then
                    ' Must be lowercase.
                    newProp.VTBool = New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                    propSet = True
                End If
        End Select

        If Not propSet Then
            ' If the code was not able to convert the 
            ' property to a valid value, throw an exception.
            Throw New InvalidDataException("propertyValue")
        End If
```
***

At this point, if the code has not thrown an exception, you can assume that the property is valid, and the code sets the `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty.FormatId` and `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty.Name` properties of the new custom property.

### [C#](#tab/cs-4)
```csharp
    // Now that you have handled the parameters, start
    // working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    newProp.Name = propertyName;
```
### [Visual Basic](#tab/vb-4)
```vb
        ' Now that you have handled the parameters, start
        ' working on the document.
        newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        newProp.Name = propertyName
```
***

## Working with the Document

Given the `DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty` object, the code next interacts with the document that you supplied in the parameters to the `SetCustomProperty` procedure. The code starts by opening the document in read/write mode by
using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A` method of the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument` class. The code attempts to retrieve a reference to the custom file properties part by using the `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.CustomFilePropertiesPart` property of the document.

### [C#](#tab/cs-5)
```csharp
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var customProps = document.CustomFilePropertiesPart;
```
### [Visual Basic](#tab/vb-5)
```vb
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
```
***

If the code cannot find a custom properties part, it creates a new part, and adds a new set of properties to the part.

### [C#](#tab/cs-6)
```csharp
        if (customProps is null)
        {
            // No custom properties? Add the part, and the
            // collection of properties now.
            customProps = document.AddCustomFilePropertiesPart();
            customProps.Properties = new Properties();
        }
```
### [Visual Basic](#tab/vb-6)
```vb
            If customProps Is Nothing Then
                ' No custom properties? Add the part, and the
                ' collection of properties now.
                customProps = document.AddCustomFilePropertiesPart()
                customProps.Properties = New Properties()
            End If
```
***

Next, the code retrieves a reference to the `DocumentFormat.OpenXml.Packaging.CustomFilePropertiesPart.Properties` property of the custom
properties part (that is, a reference to the properties themselves). If
the code had to create a new custom properties part, you know that this
reference is not null. However, for existing custom properties parts, it
is possible, although highly unlikely, that the `DocumentFormat.OpenXml.Packaging.CustomFilePropertiesPart.Properties` property will be null. If so, the code
cannot continue.

### [C#](#tab/cs-7)
```csharp
        var props = customProps.Properties;

        if (props is not null)
        {
```
### [Visual Basic](#tab/vb-7)
```vb
            Dim props = customProps.Properties

            If props IsNot Nothing Then
```
***

If the property already exists, the code retrieves its current value,
and then deletes the property. Why delete the property? If the new type
for the property matches the existing type for the property, the code
could set the value of the property to the new value. On the other hand,
if the new type does not match, the code must create a new element,
deleting the old one (it is the name of the element that defines its
type—for more information, see Figure 1). It is simpler to always delete
and then re-create the element. The code uses a simple LINQ query to
find the first match for the property name.

### [C#](#tab/cs-8)
```csharp
            var prop = props.FirstOrDefault(p => ((CustomDocumentProperty)p).Name!.Value == propertyName);

            // Does the property exist? If so, get the return value, 
            // and then delete the property.
            if (prop is not null)
            {
                returnValue = prop.InnerText;
                prop.Remove();
            }
```
### [Visual Basic](#tab/vb-8)
```vb
                Dim prop = props.FirstOrDefault(Function(p) CType(p, CustomDocumentProperty).Name.Value = propertyName)

                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop IsNot Nothing Then
                    returnValue = prop.InnerText
                    prop.Remove()
                End If
```
***

Now, you will know for sure that the custom property part exists, a property that has the same name as the new property does not exist, and that there may be other existing custom properties. The code performs the following steps:

1. Appends the new property as a child of the properties collection.

2. Loops through all the existing properties, and sets the <span class="keyword">`pid`</span> attribute to increasing values, starting at 2.

3. Saves the part.

### [C#](#tab/cs-9)
```csharp
            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            props.AppendChild(newProp);
            int pid = 2;
            foreach (CustomDocumentProperty item in props)
            {
                item.PropertyId = pid++;
            }
```
### [Visual Basic](#tab/vb-9)
```vb
                ' Append the new property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                props.AppendChild(newProp)
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
```
***

Finally, the code returns the stored original property value.

### [C#](#tab/cs-10)
```csharp
    return returnValue;
```
### [Visual Basic](#tab/vb-10)
```vb
        Return returnValue
```
***

## Sample Code

The following is the complete `SetCustomProperty` code sample in C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static string SetCustomProperty(
    string fileName,
    string propertyName,
    object propertyValue,
    PropertyTypes propertyType)
{
    // Given a document name, a property name/value, and the property type, 
    // add a custom property to a document. The method returns the original
    // value, if it existed.
    string? returnValue = string.Empty;

    var newProp = new CustomDocumentProperty();
    bool propSet = false;

    string? propertyValueString = propertyValue.ToString() ?? throw new System.ArgumentNullException("propertyValue can't be converted to a string.");

    // Calculate the correct type.
    switch (propertyType)
    {
        case PropertyTypes.DateTime:

            // Be sure you were passed a real date, 
            // and if so, format in the correct way. 
            // The date/time value passed in should 
            // represent a UTC date/time.
            if ((propertyValue) is DateTime)
            {
                newProp.VTFileTime =
                    new VTFileTime(string.Format("{0:s}Z",
                        Convert.ToDateTime(propertyValue)));
                propSet = true;
            }

            break;

        case PropertyTypes.NumberInteger:
            if ((propertyValue) is int)
            {
                newProp.VTInt32 = new VTInt32(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.NumberDouble:
            if (propertyValue is double)
            {
                newProp.VTFloat = new VTFloat(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.Text:
            newProp.VTLPWSTR = new VTLPWSTR(propertyValueString);
            propSet = true;

            break;

        case PropertyTypes.YesNo:
            if (propertyValue is bool)
            {
                // Must be lowercase.
                newProp.VTBool = new VTBool(
                  Convert.ToBoolean(propertyValue).ToString().ToLower());
                propSet = true;
            }
            break;
    }

    if (!propSet)
    {
        // If the code was not able to convert the 
        // property to a valid value, throw an exception.
        throw new InvalidDataException("propertyValue");
    }
    // Now that you have handled the parameters, start
    // working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    newProp.Name = propertyName;
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var customProps = document.CustomFilePropertiesPart;
        if (customProps is null)
        {
            // No custom properties? Add the part, and the
            // collection of properties now.
            customProps = document.AddCustomFilePropertiesPart();
            customProps.Properties = new Properties();
        }
        var props = customProps.Properties;

        if (props is not null)
        {
            // This will trigger an exception if the property's Name 
            // property is null, but if that happens, the property is damaged, 
            // and probably should raise an exception.
            var prop = props.FirstOrDefault(p => ((CustomDocumentProperty)p).Name!.Value == propertyName);

            // Does the property exist? If so, get the return value, 
            // and then delete the property.
            if (prop is not null)
            {
                returnValue = prop.InnerText;
                prop.Remove();
            }
            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            props.AppendChild(newProp);
            int pid = 2;
            foreach (CustomDocumentProperty item in props)
            {
                item.PropertyId = pid++;
            }
        }
    }
    return returnValue;
}
```

### [Visual Basic](#tab/vb)
```vb
    Function SetCustomProperty(fileName As String, propertyName As String, propertyValue As Object, propertyType As PropertyTypes) As String
        ' Given a document name, a property name/value, and the property type, 
        ' add a custom property to a document. The method returns the original
        ' value, if it existed.
        Dim returnValue As String = String.Empty

        Dim newProp As New CustomDocumentProperty()
        Dim propSet As Boolean = False

        Dim propertyValueString As String = propertyValue.ToString()
        If propertyValueString Is Nothing Then
            Throw New ArgumentNullException("propertyValue can't be converted to a string.")
        End If

        ' Calculate the correct type.
        Select Case propertyType
            Case PropertyTypes.DateTime
                ' Be sure you were passed a real date, 
                ' and if so, format in the correct way. 
                ' The date/time value passed in should 
                ' represent a UTC date/time.
                If TypeOf propertyValue Is DateTime Then
                    newProp.VTFileTime = New VTFileTime(String.Format("{0:s}Z", Convert.ToDateTime(propertyValue)))
                    propSet = True
                End If

            Case PropertyTypes.NumberInteger
                If TypeOf propertyValue Is Integer Then
                    newProp.VTInt32 = New VTInt32(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.NumberDouble
                If TypeOf propertyValue Is Double Then
                    newProp.VTFloat = New VTFloat(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.Text
                newProp.VTLPWSTR = New VTLPWSTR(propertyValueString)
                propSet = True

            Case PropertyTypes.YesNo
                If TypeOf propertyValue Is Boolean Then
                    ' Must be lowercase.
                    newProp.VTBool = New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                    propSet = True
                End If
        End Select

        If Not propSet Then
            ' If the code was not able to convert the 
            ' property to a valid value, throw an exception.
            Throw New InvalidDataException("propertyValue")
        End If
        ' Now that you have handled the parameters, start
        ' working on the document.
        newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        newProp.Name = propertyName
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            If customProps Is Nothing Then
                ' No custom properties? Add the part, and the
                ' collection of properties now.
                customProps = document.AddCustomFilePropertiesPart()
                customProps.Properties = New Properties()
            End If
            Dim props = customProps.Properties

            If props IsNot Nothing Then
                ' This will trigger an exception if the property's Name 
                ' property is null, but if that happens, the property is damaged, 
                ' and probably should raise an exception.
                Dim prop = props.FirstOrDefault(Function(p) CType(p, CustomDocumentProperty).Name.Value = propertyName)

                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop IsNot Nothing Then
                    returnValue = prop.InnerText
                    prop.Remove()
                End If
                ' Append the new property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                props.AppendChild(newProp)
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
            End If
        End Using
        Return returnValue
    End Function
```
***

## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
