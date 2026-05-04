# Working with WordprocessingML tables

Tables are stored in the main document XML as `<w:tbl/>`. Rows are `<w:tr/>`, cells are `<w:tc/>`, and cell content is usually paragraphs.

## Table markup

```xml
<w:tbl>
  <w:tr>
    <w:tc>
      <w:p><w:r><w:t>Cell text</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

Table properties, grid definitions, borders, widths, and styles are stored under table-level or cell-level property elements.

## Rust workflow

The document text helper includes text found inside table cells:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

For table-specific edits, parse `<w:tbl/>`, `<w:tr/>`, and `<w:tc/>` boundaries so updates do not affect unrelated body paragraphs.
