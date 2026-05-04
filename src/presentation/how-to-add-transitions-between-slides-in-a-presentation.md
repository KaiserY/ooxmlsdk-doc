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

Important attributes include:

| Attribute | Meaning |
|---|---|
| `advClick` | Whether a mouse click advances the slide. If omitted, the default is true. |
| `advTm` | Auto-advance time in milliseconds. If omitted, no auto-advance time is set. |
| `spd` | Transition speed, commonly `slow`, `med`, or `fast`. |

For example, a random-bar transition can advance after three seconds:

```xml
<p:transition spd="slow" advClick="1" advTm="3000">
  <p:randomBar dir="horz"/>
</p:transition>
```

## Markup compatibility

Some transition features are stored with newer Office namespaces. For example, PowerPoint 2010 introduced a transition duration attribute in the `p14` namespace. A compatible package can wrap that richer transition in `mc:AlternateContent`, with a fallback transition that omits the newer attribute.

When documenting a writer for this page, include both the richer choice and a fallback only after the fixture proves that `ooxmlsdk` preserves or processes the compatibility markup correctly for the selected feature flags.

## Rust workflow

Open the presentation and locate the target slide through the package graph:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

This page does not yet include a tested transition writer. A safe implementation should parse the slide XML, insert or replace only the `<p:transition/>` element, preserve the rest of the slide, and verify round-trip save behavior.
