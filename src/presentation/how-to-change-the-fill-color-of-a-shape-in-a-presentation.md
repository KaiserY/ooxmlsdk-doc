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

## Shape tree

Slide content lives under the shape tree (`p:spTree`). It contains the non-visual group properties, group shape properties, and then zero or more drawing objects:

| Element | Meaning |
|---|---|
| `p:sp` | Shape |
| `p:grpSp` | Group shape |
| `p:graphicFrame` | Graphic frame |
| `p:cxnSp` | Connection shape |
| `p:pic` | Picture |
| `p:extLst` | Extension list |

The upstream sample changes the first shape on the first slide, so the test file must contain at least one shape. A production writer should select by a stable shape ID or name instead.

## Rust workflow

```rust
{{#include ../../listings/presentation/src/lib.rs:change_first_shape_fill_color}}
```

The listing updates the first shape on the selected slide. For a broader writer, do not use broad text replacement across the whole slide. Parse the slide XML, locate the intended shape by id or name, update only its fill subtree, and then write the part back through the package.
