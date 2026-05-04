# Working with comments

Presentation comments are stored outside the slide XML. A slide can have a comments part, and the presentation can have a comment authors part that identifies authors used by comments.

## Comment structure

Classic PresentationML comments use a comment list rooted at `<p:cmLst/>`. Each comment has an author id, creation time, position, and text.

```xml
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-05-04T10:00:00Z" idx="1">
    <p:pos x="10" y="20"/>
    <p:text>Review this slide.</p:text>
  </p:cm>
</p:cmLst>
```

The author list is a separate presentation-level part:

```xml
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Ada" initials="AL" lastIdx="1" clrIdx="0"/>
</p:cmAuthorLst>
```

## Rust workflow

Use package relationships to find the relevant parts. Start with the presentation part for author data, and with each slide part for slide-specific comments.

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

`ooxmlsdk 0.6.0` exposes the package and part graph, but this chapter does not yet include a tested comment writer. Adding or replying to comments must coordinate comment ids, author ids, slide relationships, and in newer PowerPoint files, modern comment parts. Keep writer examples out of the docs until a fixture covers the whole graph.

## Practical guidance

For read-only tooling, inspect existing comment parts and author parts. For editing, prefer starting from a real presentation that already contains comments, then make a small schema-aware change and verify that PowerPoint can open the result.
