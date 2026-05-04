# Delete a slide from a presentation

Deleting a slide is a coordinated edit to `ppt/presentation.xml`, the presentation relationships, and possibly related notes, comments, media, and custom show data. Removing only the slide XML file leaves dangling relationships or slide ids.

## What must change

A slide is listed in `<p:sldIdLst/>`:

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId1"/>
  <p:sldId id="257" r:id="rId2"/>
</p:sldIdLst>
```

The relationship id points to a slide part:

```xml
<Relationship
  Id="rId2"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  Target="slides/slide2.xml"/>
```

A safe deletion removes the slide id entry and the relationship, then checks whether related slide resources are still referenced.

## Rust workflow

Use read-only traversal to identify the target slide before writing:

```rust
{{#include ../../listings/presentation/src/lib.rs:count_slides}}
```

This chapter does not yet include a tested deletion writer for `ooxmlsdk 0.6.0`. A final example should cover visible and hidden slides, custom shows, notes slides, comments, media, and relationship cleanup.
