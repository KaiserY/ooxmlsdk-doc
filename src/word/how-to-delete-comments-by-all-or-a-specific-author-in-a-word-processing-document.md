# Delete comments by all or a specific author in a word processing document

Deleting comments requires editing both the comments part and references in the main document body.

The comments part stores `<w:comment/>` elements with ids and author metadata. The main document story stores matching range and reference markers. Filtering by author should happen in the comments part first; then use the matching ids to remove references from document content.

## Rust workflow

Filter the comments part by author, then remove the corresponding body markers:

```rust
{{#include ../../listings/word/src/lib.rs:delete_comments_by_author}}
```

For each deleted comment id, remove all matching `commentRangeStart`, `commentRangeEnd`, and `commentReference` elements from the main document. If a package has comments in headers, footers, footnotes, or endnotes, apply the same marker cleanup in those stories too.

In `ooxmlsdk`, generated schema types include `Comments`, `Comment`, `CommentRangeStart`, `CommentRangeEnd`, and `CommentReference`. `MainDocumentPart::wordprocessing_comments_part(&document)` locates the comments part when it exists.
