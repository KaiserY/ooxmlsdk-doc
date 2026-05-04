# Change text in a table in a word processing document

Table cell text is stored in paragraphs inside `<w:tc/>` cells. Changing table text requires locating the correct table, row, cell, paragraph, and run.

## Rust workflow

The document text helper includes table cell text:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a table text writer. Avoid broad string replacement across `document.xml`; parse table boundaries and update only the target cell content.
