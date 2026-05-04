# Change the fill color of a shape in a presentation

Shape fill color is stored in the slide XML, usually under a shape's `<p:spPr/>` properties. A solid fill uses DrawingML color markup such as `<a:solidFill/>`.

## Shape fill markup

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Accent shape"/>
  </p:nvSpPr>
  <p:spPr>
    <a:solidFill>
      <a:srgbClr val="FF0000"/>
    </a:solidFill>
  </p:spPr>
</p:sp>
```

The `val` attribute stores the RGB color as six hexadecimal digits.

## Rust workflow

Use `PresentationDocument` to find the slide, then read the slide XML:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

For a real writer, do not use broad text replacement across the whole slide. Parse the slide XML, locate the intended shape by id or name, update only its fill subtree, and then write the part back through the package.

This chapter does not yet include a tested writer because the safe behavior depends on shape selection, existing fill variants, and XML namespace preservation. Add the implementation under `listings/` with fixture coverage before documenting it as final API.
