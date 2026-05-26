# Create and add a character style to a word processing document

Character styles are stored in the styles part with `w:type="character"` and are referenced from run properties.

Style ids are the internal identifiers used in WordprocessingML. Display names and aliases are UI-facing metadata and can differ from the `styleId`.

## Style markup

```xml
<w:style w:type="character" w:styleId="Emphasis">
  <w:aliases w:val="Important, Highlight"/>
  <w:name w:val="Emphasis"/>
  <w:rPr>
    <w:rFonts w:ascii="Tahoma"/>
    <w:sz w:val="48"/>
    <w:color w:themeColor="accent2"/>
    <w:b/>
    <w:i/>
  </w:rPr>
</w:style>
```

Character styles apply to runs, not whole paragraphs. Reference the style from run properties with `<w:rStyle w:val="Emphasis"/>`.

WordprocessingML supports paragraph, character, linked, table, numbering, and default paragraph/character property styles. Character styles should contain character-level run properties (`rPr`) and should not be used as paragraph styles.

## Rust workflow

Create or reuse the styles part, append a character style definition, and avoid duplicating an existing style id:

```rust
{{#include ../../listings/word/src/lib.rs:create_character_style}}
```

In `ooxmlsdk`, generated schema types include `Styles`, `Style`, `Aliases`, `StyleName`, `StyleRunProperties`, `RunStyle`, `RunFonts`, `FontSize`, `Color`, `Bold`, and `Italic`.
