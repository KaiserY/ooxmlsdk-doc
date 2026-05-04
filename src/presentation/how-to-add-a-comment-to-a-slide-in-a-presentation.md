# Add a comment to a slide in a presentation

Comments are stored in comment parts related to slides, with author information stored separately at the presentation level. Adding a comment is therefore a package graph update, not just a text edit in the slide XML.

This page describes modern PowerPoint comments. Classic comments use a different, older package shape and should be tested separately before reusing the same writer logic.

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

## Add-comment workflow

The package update has these steps:

1. Open the presentation and get the presentation part.
2. Find or create the presentation-level comment authors part.
3. Find an existing author by name and initials, or append a new author and assign an ID.
4. Locate the target slide by slide ID or slide relationship.
5. Find or create the slide's PowerPoint comments part.
6. Append the comment with author ID, timestamp, position, and text.
7. Ensure the slide has the extension list entries required by modern comments.

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

When matching author data, use the exact name and initials PowerPoint stores in the file, usually the values from PowerPoint Options under the General tab.
