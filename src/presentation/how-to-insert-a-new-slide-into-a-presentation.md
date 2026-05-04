# Insert a new slide into a presentation

Inserting a slide requires creating a slide part, adding a presentation relationship, inserting a `<p:sldId/>` entry at the desired position, and linking the slide to a layout. A valid slide usually depends on an existing slide master and slide layout.

## Relationship model

The presentation part orders slides through `<p:sldIdLst/>`:

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId1"/>
  <p:sldId id="257" r:id="rId2"/>
</p:sldIdLst>
```

The `id` value is a presentation slide id. The `r:id` value resolves to a slide part relationship.

## Rust workflow

Before inserting, inspect the current slide count and available layouts:

```rust
{{#include ../../listings/presentation/src/lib.rs:count_slides}}
```

`ooxmlsdk 0.6.0` exposes the package graph and generated part types, but this chapter does not yet publish an insertion writer. A tested example must verify:

- new slide part creation,
- relationship id allocation,
- slide id allocation,
- layout relationship creation,
- insertion order,
- saved package compatibility.

For practical implementations, copy an existing slide and adjust its content before attempting a fully from-scratch slide.
