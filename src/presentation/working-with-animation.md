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

Open the target slide part and add or replace a minimal timing tree:

```rust
{{#include ../../listings/presentation/src/lib.rs:add_basic_animation_timing}}
```

Animation markup is sensitive to ids and target references. The example targets shape id `2` from the simple fixture; in a general writer, discover the target shape id from the slide XML and keep timing node ids unique.
