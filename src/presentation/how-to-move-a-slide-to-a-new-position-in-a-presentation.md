# Move a slide to a new position in a presentation

Slide order is controlled by the order of `<p:sldId/>` elements in the presentation part. Moving a slide does not require changing the slide XML; it requires reordering the slide id list while preserving the same relationship ids.

## Slide order markup

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId1"/>
  <p:sldId id="257" r:id="rId2"/>
  <p:sldId id="258" r:id="rId3"/>
</p:sldIdLst>
```

Moving the third slide to the first position means moving the `rId3` entry before `rId1`.

## Rust workflow

Use the package model to inspect the current order and slide count:

```rust
{{#include ../../listings/presentation/src/lib.rs:count_slides}}
```

This chapter does not yet include a tested writer. A final implementation should parse `ppt/presentation.xml`, reorder only `<p:sldId/>` children, preserve ids and namespaces, save the package, and verify that slide relationships still resolve.
