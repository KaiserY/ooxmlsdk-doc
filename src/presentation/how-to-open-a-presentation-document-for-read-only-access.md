# Open a presentation document for read-only access

Use `PresentationDocument` to open a `.pptx` package and inspect its parts. In `ooxmlsdk`, `new_from_file` and `new_from_file_with_settings` read from a path and return a package value; changes are only persisted when you call `save` or `copy_to`.

For read-only inspection, open the package and avoid saving it.

`PresentationDocument::new` also accepts a `Read + Seek` reader, such as a `std::io::Cursor<Vec<u8>>` or a seekable file stream. Use that form when the package bytes come from storage other than a local path.

## Open and inspect slide parts

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

The example uses lazy package opening:

- `OpenSettings { open_mode: PackageOpenMode::Lazy, ..Default::default() }`
- `PresentationDocument::new_from_file_with_settings`
- `presentation_part()`
- `slide_parts(&document)`

Lazy opening is useful for inspection helpers because it lets you navigate the package model without parsing every root element up front.

If a caller asks for a slide index that is out of range, return an error or an empty result deliberately. The upstream sample can throw an out-of-range exception; Rust examples should make that behavior explicit in their `Result` or `Option` shape.

## Presentation package structure

A PresentationML package stores the main presentation in `ppt/presentation.xml`. Slides are separate parts, usually under `ppt/slides/`, and the presentation part owns relationships to those slide parts.

Use relationships and typed part accessors instead of hard-coding ZIP paths whenever possible.
