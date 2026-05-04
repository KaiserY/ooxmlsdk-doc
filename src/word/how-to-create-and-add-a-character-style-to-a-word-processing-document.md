# Create and add a character style to a word processing document

Character styles are stored in the styles part with `w:type="character"` and are referenced from run properties.

## Style markup

```xml
<w:style w:type="character" w:styleId="Emphasis">
  <w:name w:val="Emphasis"/>
  <w:rPr><w:i/></w:rPr>
</w:style>
```

## Rust workflow

Read existing styles first:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a style writer. A final implementation should create the styles part when missing, avoid duplicate style ids, and save valid style XML.
