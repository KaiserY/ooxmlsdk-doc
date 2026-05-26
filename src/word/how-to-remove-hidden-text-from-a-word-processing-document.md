# Remove hidden text from a word processing document

Hidden text is represented by run properties, usually `<w:vanish/>`, on runs or inherited styles.

`vanish` is a toggle property. As direct run formatting, values such as `on`, `1`, or `true` hide the run, and `off`, `0`, or `false` turn direct hidden formatting off. In styles, toggle behavior interacts with the style hierarchy, so direct markup alone is not always enough to know the effective result.

## Hidden text markup

```xml
<w:r>
  <w:rPr><w:vanish/></w:rPr>
  <w:t>Hidden text</w:t>
</w:r>
```

## Rust workflow

Remove runs whose direct run properties contain active `vanish` formatting:

```rust
{{#include ../../listings/word/src/lib.rs:remove_hidden_text}}
```

The upstream sample removes runs whose run properties contain `vanish`, and removes extra `vanish` elements when they are not under a run. That is a useful cleanup pattern for direct formatting, but a complete Rust implementation should also account for hidden text inherited from character or paragraph styles.

In `ooxmlsdk`, generated schema types include `Vanish`, `RunProperties`, `Run`, `ParagraphProperties`, and `Text`.
