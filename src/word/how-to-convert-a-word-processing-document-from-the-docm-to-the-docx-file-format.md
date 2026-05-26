# Convert a word processing document from DOCM to DOCX

Converting from macro-enabled `.docm` to `.docx` requires removing macro-related parts and changing package content types so the package no longer advertises VBA content.

The macro project is stored as a binary `vbaProject` part. Removing only the `.docm` extension is not a conversion; the package must no longer contain or reference the VBA project, and the main document content type must match a standard `.docx` document.

## Rust workflow

Open the package, delete the VBA project relationship when present, change the document type to the standard Wordprocessing document content type, and save the package to the caller's output path:

```rust
{{#include ../../listings/word/src/lib.rs:convert_docm_to_docx}}
```

This listing covers the package-level conversion path. A production converter should still decide how to handle macro-related custom UI, application-specific metadata, signatures, or external workflow rules before writing the output `.docx`.

In `ooxmlsdk`, `MainDocumentPart::vba_project_part(&document)` locates a VBA project part when one is related from the main document part. `delete_part` removes that relationship and target part from the package model, and `change_document_type(WordprocessingDocumentType::Document)` updates the main document content type.

If the source file has no VBA project part, the conversion can be a no-op at the package level, but code should still avoid overwriting an existing `.docx` output unexpectedly.
