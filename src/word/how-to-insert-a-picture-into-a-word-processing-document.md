# Insert a picture into a word processing document

Pictures require an image part, a relationship from the main document part, and DrawingML markup in the document body.

## Package model

The body markup references an image relationship id. The relationship resolves to an image part under the package.

## Rust workflow

Use the main document part as the insertion point:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a picture writer. A final implementation should add the image part, create the relationship, insert valid drawing markup, and verify the saved document.
