# Structure of a WordprocessingML document

A WordprocessingML file is an Open Packaging Convention package. The `.docx` file is a ZIP container whose parts are connected by relationship items. `ooxmlsdk` exposes that graph through `WordprocessingDocument` and generated part accessors.

The basic main-document structure is `document` -> `body` -> block-level elements such as paragraphs. A paragraph (`p`) contains one or more runs (`r`), and a run contains text (`t`) plus any run-level properties that apply to that range of text.

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
| Footnotes | `<w:footnotes/>` | `MainDocumentPart::footnotes_part(&document)` |
| Endnotes | `<w:endnotes/>` | `MainDocumentPart::endnotes_part(&document)` |
| Glossary document | `<w:glossaryDocument/>` | `MainDocumentPart::glossary_document_part(&document)` |
| Extended properties | `<Properties/>` | `WordprocessingDocument::extended_file_properties_part()` |

## Stories

WordprocessingML organizes content into stories. The main document story is required and lives in the main document part. Other stories, such as comments, headers, footers, footnotes, endnotes, glossary content, text boxes, and subdocuments, appear only when the package needs them.

A minimal valid `.docx` only needs the main document story. It does not need comments, styles, numbering, headers, or other supporting parts unless the content references them.

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

A typical document is larger than the minimum package. It often adds style definitions, section properties, headers and footers, numbering, comments, footnotes, endnotes, media, and document properties. Keep these as separate parts instead of inlining their XML into the main document part.
