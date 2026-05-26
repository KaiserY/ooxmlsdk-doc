# Open and add text to a word processing document

Adding text means editing the main document XML, usually by inserting a new paragraph or run under `<w:body/>`.

The main document part contains the text of the document as WordprocessingML. Opening a package for editing is only the first step; a writer must ensure the document root and body exist before appending content.

## Text markup

```xml
<w:p>
  <w:r><w:t>New text</w:t></w:r>
</w:p>
```

## Rust workflow

```rust
{{#include ../../listings/word/src/lib.rs:add_paragraph_text}}
```

The listing appends a paragraph before a trailing `<w:sectPr/>` if the body has section properties. A broader writer should also handle documents with missing or unusual body structure, and can build the paragraph from generated schema types instead of raw XML strings.

Unlike the upstream .NET SDK's AutoSave behavior, this book shows explicit save behavior through `document.save(...)`.
