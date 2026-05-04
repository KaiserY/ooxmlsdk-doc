# Convert a word processing document from DOCM to DOCX

Converting from macro-enabled `.docm` to `.docx` requires removing macro-related parts and changing package content types so the package no longer advertises VBA content.

## Rust workflow

Open the package and inspect the main document part first:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a converter. A correct implementation should remove VBA project relationships and parts, update content type overrides, change the main document content type when needed, and verify the saved `.docx`.
