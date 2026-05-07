# Create a package

Open XML files are ZIP-based packages. A valid package needs content types, relationships, and at least the required root part for the document category you are creating.

In `ooxmlsdk`, package read/write APIs are available through types such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument`. Use each document type's `create(...)` constructor to start a new package, add the required main part and child parts, then save to the file or writer your application owns.

## Rust workflow

The recommended documented workflow is:

1. Pick the package family and document type.
2. Create the package with `WordprocessingDocument::create`, `SpreadsheetDocument::create`, or `PresentationDocument::create`.
3. Add the required main part and any child parts.
4. Set part bytes or generated root elements.
5. Save the package to a writer or file.

The domain-specific create chapters show minimal package writers:

- [Create a word processing document by providing a file name](../word/how-to-create-a-word-processing-document-by-providing-a-file-name.md)
- [Create a spreadsheet document by providing a file name](../spreadsheet/how-to-create-a-spreadsheet-document-by-providing-a-file-name.md)
- [Create a presentation document by providing a file name](../presentation/how-to-create-a-presentation-document-by-providing-a-file-name.md)

```rust
{{#include ../../listings/getting-started/src/lib.rs:full_example}}
```

## Document type and extension

The document type controls the content type written for the main part. Keep it aligned with the extension you persist:

| Family | Normal | Template | Macro-enabled |
|---|---|---|---|
| WordprocessingML | `.docx` with `WordprocessingDocumentType::Document` | `.dotx` with `Template` | `.docm` / `.dotm` with macro-enabled variants |
| SpreadsheetML | `.xlsx` with `SpreadsheetDocumentType::Workbook` | `.xltx` with `Template` | `.xlsm`, `.xltm`, or `.xlam` with macro-enabled variants |
| PresentationML | `.pptx` with `PresentationDocumentType::Presentation` | `.potx` with `Template` | `.pptm`, `.potm`, `.ppsm`, or `.ppam` with macro-enabled variants |

Office applications can reject a package whose extension does not match its main part content type.

## Templates

Use `create_from_template` when a `.dotx`, `.xltx`, or `.potx` should become an editable regular document package. The method opens the template, changes the package document type to the default regular type for that family, and preserves the package content for further mutation and saving.

## WordprocessingML structure

The minimum main document part for a word-processing package is a `w:document` root element with a `w:body` child:

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>
```

The corresponding generated Rust types live under:

`ooxmlsdk::schemas::schemas_openxmlformats_org_wordprocessingml_2006_main`

For more about WordprocessingML package structure, see [Structure of a WordprocessingML document](../word/structure-of-a-wordprocessingml-document.md).
