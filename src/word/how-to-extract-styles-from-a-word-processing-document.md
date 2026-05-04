# Extract styles from a word processing document

Styles are stored in the style definitions part, usually `word/styles.xml`. Paragraphs and runs refer to styles by id.

Some documents also contain a `stylesWithEffects` part. Word 2013 and later can keep both the original styles part and a styles-with-effects part so the document can round-trip through older Word versions. Callers need to decide which part they want to inspect.

## Read style ids

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

The helper follows the main document's styles relationship and extracts `w:styleId` values.

## Style markup

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/>
</w:style>
```

Styles can define paragraph, character, table, and numbering behavior. A style id is the stable value referenced from document content.

For full extraction, return the complete XML for `style_definitions_part` or `styles_with_effects_part` rather than only style ids. If the requested part is absent, return `None` or an empty result explicitly.

In ooxmlsdk 0.6.0, `MainDocumentPart::style_definitions_part(&document)` locates `word/styles.xml`, and `MainDocumentPart::styles_with_effects_part(&document)` locates the styles-with-effects part when present. Generated schema types include `Styles` and `Style`.
