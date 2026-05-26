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

Open the main document part, locate section properties, and write a matching page size:

```rust
{{#include ../../listings/word/src/lib.rs:change_print_orientation}}
```

In `ooxmlsdk`, generated schema types include `SectionProperties`, `PageSize`, and `PageMargin`. The `w:orient` attribute can be absent; absence normally behaves like portrait, so a writer should avoid rewriting sections whose effective orientation already matches the requested value.
