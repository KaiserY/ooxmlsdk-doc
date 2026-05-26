# Reply to a comment in a presentation

Comment replies are more complex than simple comment insertion because PowerPoint files may use modern comment extension parts in addition to classic PresentationML comment lists.

## Model notes

Classic comments are stored as `<p:cm/>` entries in a slide comment part. Modern comments and replies can involve Office extension namespaces and additional relationship targets. A reply must preserve author identity, timestamps, threading metadata, and the relationship between the slide and its comment parts.

## Reply workflow

The upstream modern-comments sample follows this shape:

1. Open the presentation and find or create the comment authors part.
2. Match the reply author by name and initials, or create a new author record.
3. Locate the first or target slide part and read its comment parts.
4. Iterate existing comments and choose the comment that should receive a reply.
5. Find or create that comment's reply list.
6. Append the reply text with the author ID and timestamp.

## Rust workflow

Classic PresentationML comments do not have a reply subtree. For classic comments, add another comment by the reply author and preserve the author/comment ids:

```rust
{{#include ../../listings/presentation/src/lib.rs:add_comment_to_slide}}
```

For modern threaded comments, use the modern PowerPoint comment and author parts (`PowerPointCommentPart` and `PowerPointAuthorsPart`) and preserve the threading metadata. Prefer modifying a fixture that already contains a modern comment and reply so the package structure is known-good.
