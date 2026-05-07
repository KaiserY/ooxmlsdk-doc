# Create a word processing document by providing a file name

Creating a `.docx` from scratch requires package relationships, content types, and at least a main document part with valid WordprocessingML.

In `ooxmlsdk`, create the package with `WordprocessingDocument::create(WordprocessingDocumentType::Document)`, add the main document part, write valid WordprocessingML, and save the package to the file or writer your application owns.

Choose the package type and file extension together. A standard document should be saved as `.docx`; macro-enabled documents and templates require different content types and extensions. Use `WordprocessingDocumentType::MacroEnabledDocument`, `Template`, or `MacroEnabledTemplate` when the target package is not a normal `.docx`.

## Minimal package pieces

A minimal document includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `word/document.xml`,
- `word/document.xml`,
- optional supporting parts such as styles, settings, and app properties.

This minimal writer creates that structure in memory:

```rust
{{#include ../../listings/word/src/lib.rs:create_word_document}}
```

For template workflows, `WordprocessingDocument::create_from_template` opens a `.dotx` or `.dotm` as an editable regular document package. Use `document_type()` to inspect the current type and `change_document_type(...)` when converting the package content type deliberately.

The minimal main document XML is a `document` root with a `body`, usually containing at least one paragraph, run, and text element:

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Create text in body</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
```

In `ooxmlsdk`, generated schema types include `Document`, `Body`, `Paragraph`, `Run`, and `Text`. The example above writes XML bytes directly for compactness; a larger writer can instead build the generated root type and call `set_root_element`.
