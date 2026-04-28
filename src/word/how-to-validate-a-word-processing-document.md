# Validate a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically validate a word processing document.

--------------------------------------------------------------------------------
## How the Sample Code Works
This code example consists of two methods. The first method, **ValidateWordDocument**, is used to validate a
regular Word file. It doesn't throw any exceptions and closes the file
after running the validation check. The second method, **ValidateCorruptedWordDocument**, starts by
inserting some text into the body, which causes a schema error. It then
validates the Word file, in which case the method throws an exception on
trying to open the corrupted file. The validation is done by using the
`DocumentFormat.OpenXml.Validation.OpenXmlValidator.Validate` method. The code displays
information about any errors that are found, in addition to the count of
errors.

--------------------------------------------------------------------------------

> [!Important] 
> Notice that you cannot run the code twice after corrupting the file in the first run. You have to start with a new Word file.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void ValidateWordDocument(string filepath)
{
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        try
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;
            foreach (ValidationErrorInfo error in
                validator.Validate(wordprocessingDocument))
            {
                count++;
                Console.WriteLine("Error " + count);
                Console.WriteLine("Description: " + error.Description);
                Console.WriteLine("ErrorType: " + error.ErrorType);
                Console.WriteLine("Node: " + error.Node);
                if (error.Path is not null)
                {
                    Console.WriteLine("Path: " + error.Path.XPath);
                }
                if (error.Part is not null)
                {
                    Console.WriteLine("Part: " + error.Part.Uri);
                }
                Console.WriteLine("-------------------------------------------");
            }

            Console.WriteLine("count={0}", count);
        }

        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}

static void ValidateCorruptedWordDocument(string filepath)
{
    // Insert some text into the body, this would cause Schema Error
    using (WordprocessingDocument wordprocessingDocument =
    WordprocessingDocument.Open(filepath, true))
    {

        if (wordprocessingDocument.MainDocumentPart is null || wordprocessingDocument.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Insert some text into the body, this would cause Schema Error
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
        Run run = new Run(new Text("some text"));
        body.Append(run);

        try
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;
            foreach (ValidationErrorInfo error in
                validator.Validate(wordprocessingDocument))
            {
                count++;
                Console.WriteLine("Error " + count);
                Console.WriteLine("Description: " + error.Description);
                Console.WriteLine("ErrorType: " + error.ErrorType);
                Console.WriteLine("Node: " + error.Node);
                if (error.Path is not null)
                {
                    Console.WriteLine("Path: " + error.Path.XPath);
                }
                if (error.Part is not null)
                {
                    Console.WriteLine("Part: " + error.Part.Uri);
                }
                Console.WriteLine("-------------------------------------------");
            }

            Console.WriteLine("count={0}", count);
        }

        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}
```

### [Visual Basic](#tab/vb)
```vb
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Validation
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ValidateWordDocument(args(0))
        ValidateCorruptedWordDocument(args(0))
    End Sub

    Public Sub ValidateWordDocument(ByVal filepath As String)
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
    End Sub

    Public Sub ValidateCorruptedWordDocument(ByVal filepath As String)
        ' Insert some text into the body, this would cause Schema Error
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Insert some text into the body, this would cause Schema Error
            Dim body As Body = wordprocessingDocument__1.MainDocumentPart.Document.Body
            Dim run As New Run(New Text("some text"))
            body.Append(run)

            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
    End Sub
End Module
```

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
