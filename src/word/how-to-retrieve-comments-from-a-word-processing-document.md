# Retrieve comments from a word processing document

Word comments are stored in a comments part related from the main document part. The body XML contains comment references; the comment text lives in `word/comments.xml`.

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
