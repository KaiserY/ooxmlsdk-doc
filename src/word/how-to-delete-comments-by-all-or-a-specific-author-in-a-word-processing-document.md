# Delete comments by all or a specific author in a word processing document

Deleting comments requires editing both the comments part and references in the main document body.

## Rust workflow

Read the comments part first:

```rust
{{#include ../../listings/word/src/lib.rs:get_comments}}
```

This chapter does not yet publish a deletion writer. A complete implementation should remove matching `<w:comment/>` entries, remove corresponding range start/end and reference markers, and preserve unrelated comments.
