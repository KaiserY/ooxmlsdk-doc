# Set the font for a text run

Run font settings are stored in run properties with `<w:rFonts/>`.

## Font markup

```xml
<w:rPr>
  <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
</w:rPr>
```

## Rust workflow

Read text runs before selecting the target:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a font writer. A correct implementation should preserve existing run properties and update only the selected run.
