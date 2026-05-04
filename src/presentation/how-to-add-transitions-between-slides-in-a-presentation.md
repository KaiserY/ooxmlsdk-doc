# Add a transition to a slide in a presentation

Slide transitions are stored on the slide that appears after the transition. The transition markup is a child of the slide root.

## Transition markup

```xml
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:transition spd="fast">
    <p:fade/>
  </p:transition>
  <p:cSld/>
</p:sld>
```

The transition element can specify speed, advance timing, sound, and transition-specific child elements such as fade or push.

## Rust workflow

Open the presentation and locate the target slide through the package graph:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

This page does not yet include a tested transition writer. A safe implementation should parse the slide XML, insert or replace only the `<p:transition/>` element, preserve the rest of the slide, and verify round-trip save behavior.
