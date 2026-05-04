# Insert a table into a word processing document

Inserting a table is a main document edit. The table must be placed in the body before section properties and should contain valid cell content.

## Rust workflow

Use the same package traversal as other main-document operations:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter shares the same implementation boundary as [Add tables to word processing documents](how-to-add-tables-to-word-processing-documents.md): writer code should be added only with a fixture that verifies the saved package.
