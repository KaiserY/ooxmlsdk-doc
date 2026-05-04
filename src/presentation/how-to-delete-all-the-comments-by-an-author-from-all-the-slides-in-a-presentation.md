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

Use the presentation part to enumerate slides and locate comment-related parts:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

This chapter does not yet publish a deletion writer. A complete tested example should:

- read the author list,
- match the requested author name to ids,
- scan every slide comment part,
- remove only matching comments,
- preserve unrelated comments and modern comment metadata,
- remove empty comment parts only after confirming no comments remain,
- remove author entries only after their comments are gone,
- save and validate the package.

Until then, treat this as a package traversal task and keep implementation experiments in `listings/` with fixtures.
