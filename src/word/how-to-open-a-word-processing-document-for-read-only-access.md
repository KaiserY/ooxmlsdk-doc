# Open a word processing document for read-only access

Use `WordprocessingDocument` to open a `.docx` package and inspect the main document part. In `ooxmlsdk`, opening a package does not modify it; changes are persisted only when you call a save method.

Use read-only access when callers only need to inspect text, metadata, styles, comments, or relationships and the package should remain unchanged.

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

The same read-only pattern applies to path-based and stream-based inputs. A valid word processing package has at least a main document part; optional parts such as styles, comments, settings, headers, and footers may be absent and should be handled with `Option`-style control flow.

Do not call save methods from read-only helpers. If a helper attempts to mutate and save a package that was intended for inspection, treat that as a bug in the helper design.
