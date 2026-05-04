# Replace the styles parts in a word processing document

Replacing styles means copying or rewriting `word/styles.xml` and keeping the main document relationship intact.

## Rust workflow

Inspect the existing styles part first:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a styles replacement writer. A final implementation should handle missing styles parts, preserve relationship ids where possible, and verify that referenced style ids still exist.
