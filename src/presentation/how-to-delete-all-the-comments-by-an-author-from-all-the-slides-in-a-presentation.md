# Delete all the comments by an author from all the slides in a presentation

Deleting comments by author requires scanning slide comment parts and matching each comment's author id against the presentation comment authors part.

This page describes modern PowerPoint comments. Classic comments have a different archived package shape and should be handled by a separate tested fixture.

## Package model

A presentation comment is a text note attached to a slide. It stores unformatted text, author information, and a slide position. Comments can be visible while editing the presentation, but they are not part of the slide show; the viewing application decides when and how to display them.

The author list maps author ids to names and initials:

```xml
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Ada" initials="AL" lastIdx="3" clrIdx="0"/>
</p:cmAuthorLst>
```

Slide comment parts then use `authorId`:

```xml
<p:cm authorId="0" idx="1">
  <p:text>Review this slide.</p:text>
</p:cm>
```

The author name must match the user name stored by PowerPoint. In the PowerPoint UI, that value is shown under File, Options, General.

## Delete workflow

The upstream modern-comments sample follows this package traversal:

1. Open the presentation for editing and get the presentation part.
2. Read the comment authors part and find authors whose `name` matches the requested author.
3. Iterate every slide part in the presentation.
4. For each slide comment part, remove comments whose `authorId` matches one of those author IDs.
5. If a slide comment part becomes empty, remove that comment part relationship.
6. Remove the matched author entries from the comment authors part.

## Rust workflow

Match classic comment author entries by name, scan slide comments parts, and remove comments with matching author ids:

```rust
{{#include ../../listings/presentation/src/lib.rs:delete_comments_by_author}}
```

Modern threaded comments use different parts and metadata. Apply the same author-id filtering idea there only after adding fixtures for those modern parts.
