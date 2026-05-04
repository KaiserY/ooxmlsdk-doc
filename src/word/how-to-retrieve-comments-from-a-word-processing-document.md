# Retrieve comments from a word processing document

Word comments are stored in a comments part related from the main document part. The body XML contains comment references; the comment text lives in `word/comments.xml`.

Open the package for read-only access when retrieving comments. If the main document part has no comments relationship, the document has no comment entries to return.

## Read comments

```rust
{{#include ../../listings/word/src/lib.rs:get_comments}}
```

The helper opens the main document part, follows the comments relationship when present, and extracts text from the comments XML.

## Comment markup

```xml
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Ada">
    <w:p><w:r><w:t>Review this paragraph</w:t></w:r></w:p>
  </w:comment>
</w:comments>
```

Use the comment id to connect body references with comment entries.

The `comments` element is the root of the comments part and contains zero or more `comment` children. A `comment` can contain block-level WordprocessingML content such as paragraphs. If a comment id is not referenced by a matching `commentReference` in document content, a consuming application may ignore it; if more than one comment has the same id, only one may be loaded.

In ooxmlsdk 0.6.0, generated schema types include `Comments`, `Comment`, and `CommentReference`.
