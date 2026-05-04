# Insert a comment into a word processing document

Inserting a comment requires updating the comments part and adding matching comment range markers or references in the main document body.

The upstream sample attaches a comment to the first paragraph. A general Rust API should make the target range explicit and should fail clearly if the selected paragraph or range does not exist.

## Comment markup

```xml
<w:comment w:id="0" w:author="Ada">
  <w:p><w:r><w:t>Review this paragraph</w:t></w:r></w:p>
</w:comment>
```

The same id must be used in the comment and in the document markers:

```xml
<w:commentRangeStart w:id="0"/>
<w:r><w:t>Commented text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:commentReference w:id="0"/></w:r>
```

## Rust workflow

Read existing comments before allocating ids:

```rust
{{#include ../../listings/word/src/lib.rs:get_comments}}
```

This chapter does not yet publish a comment writer. A safe implementation must create or update `word/comments.xml`, add body references, preserve existing ids, and save the package.

Allocate a new comment id by scanning existing comments and adding one to the maximum id. If the comments part is absent, create it with a `comments` root before appending the new `comment`.

In ooxmlsdk 0.6.0, generated schema types include `Comments`, `Comment`, `CommentRangeStart`, `CommentRangeEnd`, and `CommentReference`.
