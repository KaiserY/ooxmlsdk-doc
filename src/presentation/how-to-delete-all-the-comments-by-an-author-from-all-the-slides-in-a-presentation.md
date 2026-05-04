# Delete all the comments by an author from all the slides in a presentation

Deleting comments by author requires scanning slide comment parts and matching each comment's author id against the presentation comment authors part.

## Package model

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
- save and validate the package.

Until then, treat this as a package traversal task and keep implementation experiments in `listings/` with fixtures.
