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

Read existing styles first:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a style writer. A final implementation should create the styles part when missing, avoid duplicate style ids, and save valid style XML.

In ooxmlsdk 0.6.0, generated schema types include `Styles`, `Style`, `Aliases`, `StyleName`, `StyleRunProperties`, `RunStyle`, `RunFonts`, `FontSize`, `Color`, `Bold`, and `Italic`.
