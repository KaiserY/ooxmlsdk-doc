# Copy a worksheet using streaming XML

Copying a worksheet with a streaming parser is useful for large sheets because the cell XML can be processed without loading a full worksheet object model.

## Package model

A copied worksheet needs more than copied XML. The workbook must get a new `<sheet/>` entry and relationship, and any worksheet-owned relationships may need to be copied or adjusted.

```xml
<sheet name="Copied Sheet" sheetId="3" r:id="rId3"/>
```

## Rust workflow

Use `ooxmlsdk` to locate the source worksheet part:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a copy writer. A safe implementation must copy worksheet XML, relationships, tables, drawings, comments, printer settings, and workbook metadata consistently. Add that implementation to `listings/spreadsheet` with a fixture before documenting final code.
