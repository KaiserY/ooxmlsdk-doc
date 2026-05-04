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

Read the current document text first:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a text writer. A safe implementation should parse the body XML, insert a valid paragraph or run, preserve section properties, and save the package through `ooxmlsdk`.

When appending a paragraph to the document body, insert it before a trailing `<w:sectPr/>` if the body has section properties. The paragraph itself should be built from `Paragraph`, `Run`, and `Text` equivalents, with escaping handled by XML serialization rather than string concatenation.

Unlike the upstream .NET SDK's AutoSave behavior, this book should show explicit save behavior once a writer listing is added.
