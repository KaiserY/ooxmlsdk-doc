# Apply a theme to a presentation

A presentation theme is stored in a theme part rooted at `<a:theme/>`. Applying a theme means copying or replacing the target presentation's theme part and keeping relationships valid.

## Theme structure

The theme root contains theme elements, optional object defaults, extra color schemes, custom colors, and extensions.

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office"/>
    <a:fontScheme name="Office"/>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>
```

In a PresentationML package, the presentation part usually owns the theme relationship.

## Rust workflow

Open the source presentation read-only, open or copy the target package, and use the presentation part's theme accessor to read the theme data. The general package navigation pattern is:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

This chapter intentionally does not include a presentation theme writer yet. A complete example should verify:

- source theme part exists,
- target theme relationship is added or replaced,
- dependent slide masters and layouts still resolve their theme references,
- the saved package opens in PowerPoint or another strict consumer.

The WordprocessingML theme replacement example in the General section shows the package-level idea; PresentationML needs its own tested fixture before publishing final code.
