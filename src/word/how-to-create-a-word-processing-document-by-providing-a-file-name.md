# Create a word processing document by providing a file name

Creating a `.docx` from scratch requires package relationships, content types, and at least a main document part with valid WordprocessingML.

## Minimal package pieces

A minimal document includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `word/document.xml`,
- `word/document.xml`,
- optional supporting parts such as styles, settings, and app properties.

The `listings/word` fixture builds this structure so documented readers run against a real `.docx` package.

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a from-scratch writer. A final writer should validate relationships, content type overrides, main document XML, and save behavior together.
