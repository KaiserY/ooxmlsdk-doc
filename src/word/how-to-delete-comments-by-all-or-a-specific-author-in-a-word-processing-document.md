# Delete comments by all or a specific author in a word processing document

Deleting comments requires editing both the comments part and references in the main document body.

The comments part stores `<w:comment/>` elements with ids and author metadata. The main document story stores matching range and reference markers. Filtering by author should happen in the comments part first; then use the matching ids to remove references from document content.

## Rust workflow

Read the comments part first:

```rust
{{#include ../../listings/word/src/lib.rs:get_comments}}
```

This chapter does not yet publish a deletion writer. A complete implementation should remove matching `<w:comment/>` entries, remove corresponding range start/end and reference markers, and preserve unrelated comments.

For each deleted comment id, remove all matching `commentRangeStart`, `commentRangeEnd`, and `commentReference` elements from the main document. If a package has comments in headers, footers, footnotes, or endnotes, apply the same marker cleanup in those stories too.

In ooxmlsdk 0.6.0, generated schema types include `Comments`, `Comment`, `CommentRangeStart`, `CommentRangeEnd`, and `CommentReference`. `MainDocumentPart::wordprocessing_comments_part(&document)` locates the comments part when it exists.
