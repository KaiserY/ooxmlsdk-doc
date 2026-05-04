# Set the font for a text run

Run font settings are stored in run properties with `<w:rFonts/>`.

Run fonts can specify different faces for different character classes: ASCII, high ANSI, complex script, and East Asian text. The effective font depends on the Unicode characters in the run unless overridden by related properties.

## Font markup

```xml
<w:rPr>
  <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Times New Roman"/>
</w:rPr>
```

For a mixed English and Arabic run, for example, ASCII characters can use the ASCII font while Arabic characters use the complex-script font.

## Rust workflow

Read text runs before selecting the target:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a font writer. A correct implementation should preserve existing run properties and update only the selected run.

If the selected run has no run properties, create `rPr` and prepend it before run content. If it already has run properties, update or add only the `rFonts` child so other properties such as bold, italic, color, and style remain intact.

In ooxmlsdk 0.6.0, generated schema types include `RunProperties`, `RunFonts`, `Run`, and `Text`.
