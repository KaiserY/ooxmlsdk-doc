# Open a word processing document for read-only access

Use `WordprocessingDocument` to open a `.docx` package and inspect the main document part. In `ooxmlsdk`, opening a package does not modify it; changes are persisted only when you call a save method.

## Open and inspect paragraphs

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

The example uses lazy package opening:

- `OpenSettings { open_mode: PackageOpenMode::Lazy, ..Default::default() }`
- `WordprocessingDocument::new_from_file_with_settings`
- `main_document_part()`
- `data_as_str(&document)`

Lazy opening is useful for read-only inspection because it lets you navigate package parts without parsing every root element up front.
