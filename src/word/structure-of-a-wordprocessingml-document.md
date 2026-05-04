# Structure of a WordprocessingML document

A WordprocessingML file is an Open Packaging Convention package. The `.docx` file is a ZIP container whose parts are connected by relationship items. `ooxmlsdk` exposes that graph through `WordprocessingDocument` and generated part accessors.

## Package parts

| Package part | Root element | `ooxmlsdk` access |
|---|---|---|
| Main document | `<w:document/>` | `WordprocessingDocument::main_document_part()` |
| Styles | `<w:styles/>` | `MainDocumentPart::style_definitions_part(&document)` |
| Comments | `<w:comments/>` | `MainDocumentPart::wordprocessing_comments_part(&document)` |
| Settings | `<w:settings/>` | `MainDocumentPart::document_settings_part(&document)` |
| Numbering | `<w:numbering/>` | `MainDocumentPart::numbering_definitions_part(&document)` |
| Headers | `<w:hdr/>` | `MainDocumentPart::header_parts(&document)` |
| Footers | `<w:ftr/>` | `MainDocumentPart::footer_parts(&document)` |
| Extended properties | `<Properties/>` | `WordprocessingDocument::extended_file_properties_part()` |

## Main document XML

The main document part stores body content under `<w:body/>`.

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>
```

Open and inspect it through the package model:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

Use generated part accessors for relationships, and read XML from parts only when the chapter needs schema-level inspection.
