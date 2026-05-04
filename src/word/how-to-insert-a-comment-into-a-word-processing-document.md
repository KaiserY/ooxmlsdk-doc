# Insert a comment into a word processing document

Inserting a comment requires updating the comments part and adding matching comment range markers or references in the main document body.

## Comment markup

```xml
<w:comment w:id="0" w:author="Ada">
  <w:p><w:r><w:t>Review this paragraph</w:t></w:r></w:p>
</w:comment>
```

## Rust workflow

Read existing comments before allocating ids:

```rust
{{#include ../../listings/word/src/lib.rs:get_comments}}
```

This chapter does not yet publish a comment writer. A safe implementation must create or update `word/comments.xml`, add body references, preserve existing ids, and save the package.
