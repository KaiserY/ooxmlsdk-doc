# Add a comment to a slide in a presentation

Comments are stored in comment parts related to slides, with author information stored separately at the presentation level. Adding a comment is therefore a package graph update, not just a text edit in the slide XML.

## Comment parts

A comment list can look like this:

```xml
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-05-04T10:00:00Z" idx="1">
    <p:pos x="10" y="20"/>
    <p:text>Review this slide.</p:text>
  </p:cm>
</p:cmLst>
```

The matching author entry is stored in a presentation-level comment authors part.

## Rust workflow

Navigate from the presentation to the target slide first:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

`ooxmlsdk 0.6.0` exposes the package graph needed for comment parts, but this page does not yet publish a writer. A tested comment insertion example must cover:

- creating or finding the comment authors part,
- creating or finding the slide comment part,
- assigning author and comment ids,
- preserving existing comments,
- saving and validating the package.

Modern PowerPoint comments can also use newer Office extension parts, so this chapter keeps the write path as future fixture-backed work.
