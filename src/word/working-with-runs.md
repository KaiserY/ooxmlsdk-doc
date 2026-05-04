# Working with runs

Runs are inline text containers inside paragraphs. A run can carry formatting such as bold, italic, color, font, or field-related markup.

A run defines a region of content with a common set of properties. Besides text, runs can contain breaks, tabs, drawings, field-related markup, comments references, and other inline content.

## Run markup

```xml
<w:r>
  <w:rPr>
    <w:b/>
  </w:rPr>
  <w:t>Bold text</w:t>
</w:r>
```

Run properties are stored in `<w:rPr/>`; visible text is stored in `<w:t/>`. Text for a sentence can be split across many runs, especially after editing in Word.

If present, `rPr` must appear before the run content it formats. Common run properties include bold, border, character style, color, font, font size, italic, kerning, spelling/grammar suppression, shading, small caps, strikethrough, text direction, and underline.

## Rust workflow

The text helper extracts `<w:t/>` values from the main document:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

For formatting-sensitive tasks, preserve run boundaries and update only the target run or run property subtree.

In ooxmlsdk 0.6.0, generated schema types include `Run`, `RunProperties`, `Text`, `Bold`, `Italic`, `Color`, `RunFonts`, `FontSize`, `Underline`, `Strike`, and `SmallCaps`.
