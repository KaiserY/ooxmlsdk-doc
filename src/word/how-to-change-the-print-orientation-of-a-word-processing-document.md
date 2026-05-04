# Change the print orientation of a word processing document

Page orientation is stored in section properties, usually in `<w:pgSz/>` under `<w:sectPr/>`.

## Orientation markup

```xml
<w:sectPr>
  <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
</w:sectPr>
```

## Rust workflow

Open the main document part and inspect section properties:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish an orientation writer. A safe implementation must update the intended section, swap width and height when needed, and preserve other section settings.
