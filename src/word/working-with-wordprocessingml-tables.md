# Working with WordprocessingML tables

Tables are stored in the main document XML as `<w:tbl/>`. Rows are `<w:tr/>`, cells are `<w:tc/>`, and cell content is usually paragraphs.

Tables are block-level content, arranged as rows and columns. A table can contain paragraph content and other block-level content inside cells.

## Table markup

```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="5000" w:type="pct"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol/>
    <w:gridCol/>
    <w:gridCol/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr/>
      <w:p><w:r><w:t>1</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr/>
      <w:p><w:r><w:t>2</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr/>
      <w:p><w:r><w:t>3</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

`w:tblPr` is required for a `w:tbl` in the generated schema and defines table-wide properties, such as style and width. `w:tblGrid` is optional and defines grid layout through `w:gridCol` children. Each `w:tr` can have row properties, and each `w:tc` can optionally include `w:tcPr` for cell-specific settings such as width, borders, margins, and vertical alignment.

## Rust workflow

The document text helper includes text found inside table cells:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

For table-specific edits, parse `<w:tbl/>`, `<w:tr/>`, and `<w:tc/>` boundaries so updates do not affect unrelated body paragraphs.

In `ooxmlsdk`, generated schema types include `Table`, `TableProperties`, `TableGrid`, `GridColumn`, `TableRow`, `TableRowProperties`, `TableCell`, `TableCellProperties`, `TableStyle`, `TableWidth`, and `TableBorders`.
