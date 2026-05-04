# Create a word processing document by providing a file name

Creating a `.docx` from scratch requires package relationships, content types, and at least a main document part with valid WordprocessingML.

Choose the package type and file extension together. A standard document should be saved as `.docx`; macro-enabled documents and templates require different content types and extensions.

## Minimal package pieces

A minimal document includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `word/document.xml`,
- `word/document.xml`,
- optional supporting parts such as styles, settings, and app properties.

The `listings/word` fixture builds this structure so documented readers run against a real `.docx` package.

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a from-scratch writer. A final writer should validate relationships, content type overrides, main document XML, and save behavior together.

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

In ooxmlsdk 0.6.0, generated schema types include `Document`, `Body`, `Paragraph`, `Run`, and `Text`.
