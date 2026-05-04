# Working with animation

PresentationML stores animation data in slide timing markup. The core element is `<p:timing/>`, which contains timing nodes and behavior elements such as `<p:anim/>`.

The animation model is loosely based on SMIL. Slide animations are time-based effects applied to slide objects or text. Slide transitions are related, but they are stored in `<p:transition/>` and occur before the slide's own animation timeline.

## Animation structure

An animation behavior usually points at a target element on the slide and defines how a value changes over time.

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never"/>
    </p:par>
  </p:tnLst>
</p:timing>
```

Important animation-related elements include:

| PresentationML element | Purpose |
|---|---|
| `<p:timing/>` | Container for slide timing and animation data |
| `<p:tnLst/>` | Time node list |
| `<p:anim/>` | Value animation behavior |
| `<p:cBhvr/>` | Common behavior settings |
| `<p:tgtEl/>` | Target element for the behavior |
| `<p:tavLst/>` | Time-animated value list |

Important `<p:anim/>` attributes include:

| Attribute | Meaning |
|---|---|
| `by` | Relative offset from the starting value |
| `from` | Starting value |
| `to` | Ending value |
| `calcmode` | Interpolation mode |
| `valueType` | Type of the animated property value |

In `ooxmlsdk`, the corresponding generated types live under `ooxmlsdk::schemas::p`, including `Animate`, `CommonBehavior`, `TimeAnimateValueList`, `Timing`, and `TargetElement`.

## Rust workflow

`ooxmlsdk 0.6.0` can open the presentation package and read slide part XML. A conservative animation inspection workflow is:

1. Open the `.pptx` with `PresentationDocument`.
2. Iterate slide parts with `presentation_part.slide_parts(&document)`.
3. Read each slide with `data_as_str(&document)`.
4. Parse or inspect the `<p:timing/>` subtree.

Use the slide traversal pattern from the text extraction example:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

This chapter does not yet include a tested animation writer. Animation markup is sensitive to ids and target references, so editing should be added only with a fixture that round-trips through the package model.
