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

```rust
{{#include ../../listings/word/src/lib.rs:set_first_run_font}}
```

The listing updates the first run. If the selected run has no run properties, it creates `rPr` and prepends it before run content. If it already has run properties, a complete writer should update or add only the `rFonts` child so other properties such as bold, italic, color, and style remain intact.

In `ooxmlsdk`, generated schema types include `RunProperties`, `RunFonts`, `Run`, and `Text`.
