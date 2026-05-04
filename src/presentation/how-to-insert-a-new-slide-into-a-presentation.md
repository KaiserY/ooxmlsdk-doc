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

## Insert workflow

The upstream sample takes a file path, a zero-based insertion index, and a title string. It then:

1. Creates a new `p:sld` root and common slide data.
2. Adds a title shape and sets its text.
3. Adds a body shape and sets its text or placeholder properties.
4. Creates a new slide part.
5. Inserts a new `p:sldId` entry at the requested index.
6. Assigns the new slide root to the new slide part.

In a complete package, the new slide should also have a slide layout relationship. Copying the layout relationship from a nearby slide is often safer than inventing one from scratch.

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
