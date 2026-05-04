# Get the contents of a document part from a package

This example reads the raw XML payload of the Wordprocessing comments part from a `.docx` package.

The comments part is an optional child of the main document part. In Rust, that optional relationship is represented as `Option`.

## Comments element

The comments part root is `w:comments`. It contains the comments defined in the current WordprocessingML document. A typical part contains zero or more `w:comment` children:

```xml
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Author">
    <!-- comment content -->
  </w:comment>
</w:comments>
```

The schema shape is a sequence of `comment` elements with `minOccurs="0"` and `maxOccurs="unbounded"`.

## Read the comments part

```rust
{{#include ../../listings/getting-started/src/lib.rs:read_comments_part}}
```

The function returns:

- `Ok(Some(xml))` when the package has a comments part and the part data is valid UTF-8.
- `Ok(None)` when the main document part has no comments part relationship.
- `Err(_)` when the package cannot be opened, the main document part is missing, or the part data cannot be read as text.

## Raw part data vs. root elements

Use raw part data when you need to inspect or copy XML exactly as stored in the package.

Use `root_element` when you want `ooxmlsdk` to parse the part into the generated schema type for that part. Parsing is better when you want typed access to elements and attributes; raw data is better for simple pass-through and diagnostics.
