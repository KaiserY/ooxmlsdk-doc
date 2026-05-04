# Accept all revisions in a word processing document

Tracked revisions are stored as WordprocessingML markup in the main document and supporting parts. Accepting revisions means transforming inserted, deleted, moved, and property-change markup into the final accepted content.

The document structure still follows the normal WordprocessingML hierarchy: `document`, `body`, paragraphs (`p`), runs (`r`), and text (`t`). Revision elements wrap or annotate that content and must be resolved without breaking the surrounding paragraph, run, and table structure.

## Revision markup

Common revision elements include:

| Element | Meaning when accepting |
|---|---|
| `<w:pPrChange/>` | Keep the current paragraph properties and remove the tracked previous state. |
| `<w:del/>` | Remove deleted content or paragraph mark revision metadata, depending on where it appears. |
| `<w:ins/>` | Keep inserted content and remove insertion metadata. |
| `<w:moveFrom/>` / ranges | Remove the move source content. |
| `<w:moveTo/>` / ranges | Keep the move destination content and remove move metadata. |

In ooxmlsdk 0.6.0, generated schema types for these elements include `ParagraphPropertiesChange`, `Deleted`, `Inserted`, `MoveFrom`, `MoveTo`, `MoveFromRangeStart`, `MoveFromRangeEnd`, `MoveToRangeStart`, and `MoveToRangeEnd`.

## Rust workflow

Start with the main document XML:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish an accept-revisions writer. A complete implementation must handle insertions, deletions, move ranges, formatting changes, comments, and related parts with fixtures for each revision type.

Do not treat revision acceptance as a simple string replacement. Some revisions live in paragraph properties, table row properties, run content, comments, headers, footers, footnotes, or endnotes. A complete writer should apply the same acceptance rules to every story that can contain tracked changes.
