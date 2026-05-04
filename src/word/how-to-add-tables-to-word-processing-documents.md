# Add tables to word processing documents

Tables are inserted as `<w:tbl/>` elements in the main document body. A table contains rows, cells, and usually paragraphs inside each cell.

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

## Rust workflow

Read the current body content before inserting:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a table writer. A final implementation should create table properties, rows, cells, cell paragraphs, and insert the table without moving or losing `<w:sectPr/>`.
