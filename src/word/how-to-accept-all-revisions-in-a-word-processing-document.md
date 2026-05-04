# Accept all revisions in a word processing document

Tracked revisions are stored as WordprocessingML markup in the main document and supporting parts. Accepting revisions means transforming inserted, deleted, moved, and property-change markup into the final accepted content.

## Rust workflow

Start with the main document XML:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish an accept-revisions writer. A complete implementation must handle insertions, deletions, move ranges, formatting changes, comments, and related parts with fixtures for each revision type.
