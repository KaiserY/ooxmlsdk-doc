# Change the print orientation of a word processing document

Page orientation is stored in section properties, usually in `<w:pgSz/>` under `<w:sectPr/>`.

A document can have more than one section. Updating orientation for the whole document means walking every section properties element, not just the final `<w:sectPr/>` in the body.

## Orientation markup

```xml
<w:sectPr>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
  <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
</w:sectPr>
```

When changing between portrait and landscape, update `w:orient` and swap page width and height. If margins are present, rotate them to match the new page orientation. Some printers treat the rotation direction differently, so code that must preserve exact physical margins should make that policy explicit.

## Rust workflow

Open the main document part and inspect section properties:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish an orientation writer. A safe implementation must update the intended section, swap width and height when needed, and preserve other section settings.

In ooxmlsdk 0.6.0, generated schema types include `SectionProperties`, `PageSize`, and `PageMargin`. The `w:orient` attribute can be absent; absence normally behaves like portrait, so a writer should avoid rewriting sections whose effective orientation already matches the requested value.
