# Add tables to word processing documents

Tables are inserted as `<w:tbl/>` elements in the main document body. A table contains rows, cells, and usually paragraphs inside each cell.

The caller typically supplies a rectangular collection of strings. A writer turns each outer item into a table row and each value into a table cell containing at least one paragraph and run.

## Table markup

```xml
<w:tbl>
  <w:tblPr>
    <w:tblBorders>
      <w:top w:val="single" w:sz="12"/>
      <w:bottom w:val="single" w:sz="12"/>
      <w:left w:val="single" w:sz="12"/>
      <w:right w:val="single" w:sz="12"/>
      <w:insideH w:val="single" w:sz="12"/>
      <w:insideV w:val="single" w:sz="12"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:type="auto"/></w:tcPr>
      <w:p><w:r><w:t>Cell text</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

In `ooxmlsdk`, generated schema types include `Table`, `TableProperties`, `TableBorders`, `TableRow`, `TableCell`, `TableCellProperties`, `TableCellWidth`, `Paragraph`, `Run`, and `Text`.

## Rust workflow

```rust
{{#include ../../listings/word/src/lib.rs:insert_table}}
```

The listing creates a simple table from string rows and cells, then inserts it before the body section properties (`<w:sectPr/>`) if that element is present at the end of the body.

A production table writer can extend this with table properties, borders, widths, grid columns, and style ids. Appending after section properties can produce invalid or surprising document structure.
