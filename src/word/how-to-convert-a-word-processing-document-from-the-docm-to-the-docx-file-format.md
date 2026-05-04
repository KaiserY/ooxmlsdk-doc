# Convert a word processing document from DOCM to DOCX

Converting from macro-enabled `.docm` to `.docx` requires removing macro-related parts and changing package content types so the package no longer advertises VBA content.

The macro project is stored as a binary `vbaProject` part. Removing only the `.docm` extension is not a conversion; the package must no longer contain or reference the VBA project, and the main document content type must match a standard `.docx` document.

## Rust workflow

Open the package and inspect the main document part first:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a converter. A correct implementation should remove VBA project relationships and parts, update content type overrides, change the main document content type when needed, and verify the saved `.docx`.

In ooxmlsdk 0.6.0, `MainDocumentPart::vba_project_part(&document)` can locate a VBA project part when one is related from the main document part. A full converter also needs package mutation support for deleting that part, removing its relationship, changing content types, and saving to the new `.docx` path.

If the source file has no VBA project part, the conversion can be a no-op at the package level, but code should still avoid overwriting an existing `.docx` output unexpectedly.
