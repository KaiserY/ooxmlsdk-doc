# Insert a table into a word processing document

Inserting a table is a main document edit. The table must be placed in the body before section properties and should contain valid cell content.

Tables are represented by `tbl` elements. A table can contain table-wide properties (`tblPr`), a table grid (`tblGrid`), rows (`tr`), cells (`tc`), and paragraphs inside each cell.

## Table markup

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
      <w:p><w:r><w:t>Hello, World!</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

## Rust workflow

Use the same package traversal as other main-document operations:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter shares the same implementation boundary as [Add tables to word processing documents](how-to-add-tables-to-word-processing-documents.md): writer code should be added only with a fixture that verifies the saved package.

The sample flow creates table properties, adds border elements, creates a row, creates a cell with width properties and a paragraph/run/text child, then appends the table to the document body. If cloning cells, clone XML structure carefully so relationship-backed content is not duplicated incorrectly.

In ooxmlsdk 0.6.0, generated schema types include `Table`, `TableProperties`, `TableWidth`, `TableBorders`, `TopBorder`, `LeftBorder`, `BottomBorder`, `RightBorder`, `TableGrid`, `GridColumn`, `TableRow`, `TableCell`, `TableCellProperties`, `TableCellWidth`, `Paragraph`, `Run`, and `Text`.
