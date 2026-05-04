# Retrieve the number of slides in a presentation document

This example counts slide parts in a PresentationML package. It can include all slides or skip slides marked hidden with `show="0"` or `show="false"` in the slide XML.

## Count slides

```rust
{{#include ../../listings/presentation/src/lib.rs:count_slides}}
```

The function opens the package lazily, gets the presentation part, and iterates `presentation_part.slide_parts(&document)`.

When `include_hidden` is `true`, the function returns the number of slide parts. When `include_hidden` is `false`, it reads each slide part as XML and skips slides whose root has a hidden `show` value.

## Notes

This example uses raw slide XML to check the `show` attribute. That keeps the example focused on package navigation and avoids loading every slide root. If you are already parsing slide roots for other edits, use the generated slide schema type instead.
