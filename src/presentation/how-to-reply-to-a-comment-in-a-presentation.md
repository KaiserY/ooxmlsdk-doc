# Reply to a comment in a presentation

Comment replies are more complex than simple comment insertion because PowerPoint files may use modern comment extension parts in addition to classic PresentationML comment lists.

## Model notes

Classic comments are stored as `<p:cm/>` entries in a slide comment part. Modern comments and replies can involve Office extension namespaces and additional relationship targets. A reply must preserve author identity, timestamps, threading metadata, and the relationship between the slide and its comment parts.

## Rust workflow

Start by opening the presentation and locating the target slide:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

This chapter does not yet include a tested Rust reply writer for `ooxmlsdk 0.6.0`. Before documenting one, add fixture coverage for:

- presentations with classic comments,
- presentations with modern threaded comments,
- author lookup or creation,
- reply id allocation,
- round-trip package save.

For the first implementation, prefer modifying a fixture that already contains comments and replies so the package structure is known-good.
