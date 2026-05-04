# Working with paragraphs

Paragraphs are stored as `<w:p/>` elements in the main document body, comments, headers, footers, footnotes, and other WordprocessingML parts.

## Paragraph markup

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r><w:t>Heading text</w:t></w:r>
</w:p>
```

Paragraph properties are stored in `<w:pPr/>`. Text content is usually stored in runs under the paragraph.

## Rust workflow

Use the main document part to read paragraph text:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

If you need paragraph boundaries or properties, parse the main document XML and inspect `<w:p/>` nodes instead of flattening all text.
