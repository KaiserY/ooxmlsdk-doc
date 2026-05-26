# Create and add a paragraph style to a word processing document

Paragraph styles are stored in the styles part with `w:type="paragraph"`.

A paragraph style applies to an entire paragraph and its paragraph mark. It can define paragraph-level properties (`pPr`) and character-level properties (`rPr`) that apply to runs in that paragraph.

## Style markup

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="Normal"/>
  <w:link w:val="Heading1Char"/>
  <w:pPr/>
  <w:rPr/>
</w:style>
```

The `next` element controls the style applied to the next paragraph when editing. The `link` element can associate a paragraph style with a character style. The `styleId` is the value referenced by paragraph properties:

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>
</w:pPr>
```

## Rust workflow

Create or reuse the styles part, append a paragraph style definition, and preserve existing styles:

```rust
{{#include ../../listings/word/src/lib.rs:create_paragraph_style}}
```

In `ooxmlsdk`, generated schema types include `Styles`, `Style`, `Aliases`, `StyleName`, `BasedOn`, `NextParagraphStyle`, `LinkedStyle`, `PrimaryStyle`, `StyleParagraphProperties`, `StyleRunProperties`, and `ParagraphStyleId`.
