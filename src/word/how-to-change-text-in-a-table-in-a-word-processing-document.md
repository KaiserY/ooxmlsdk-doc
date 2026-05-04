# Change text in a table in a word processing document

Table cell text is stored in paragraphs inside `<w:tc/>` cells. Changing table text requires locating the correct table, row, cell, paragraph, and run.

Tables are block-level content in the document body. A table (`tbl`) contains table properties, an optional table grid, rows (`tr`), cells (`tc`), and block-level content inside each cell. Even an otherwise empty cell normally contains a paragraph.

## Table structure

```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4"/>
      <w:left w:val="single" w:sz="4"/>
      <w:bottom w:val="single" w:sz="4"/>
      <w:right w:val="single" w:sz="4"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="10296"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p><w:r><w:t>Cell text</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

## Rust workflow

The document text helper includes table cell text:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a table text writer. Avoid broad string replacement across `document.xml`; parse table boundaries and update only the target cell content.

The upstream sample targets the first table, second row, and third cell, then replaces text in the first run of the first paragraph. A production Rust API should make table, row, and cell selection explicit, handle missing rows or cells as errors, and decide whether replacing text should preserve existing runs or rebuild the cell paragraph.

In ooxmlsdk 0.6.0, generated schema types include `Table`, `TableProperties`, `TableGrid`, `GridColumn`, `TableRow`, `TableCell`, `TableCellProperties`, `Paragraph`, `Run`, and `Text`.
