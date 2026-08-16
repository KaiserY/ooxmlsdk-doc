# Remove a document part from a package

This example removes the Wordprocessing settings part from a `.docx` package.

In Open XML, removing a child Part starts by removing a relationship from the parent. `ooxmlsdk` deletes the target payload only when it is no longer reachable through another package or Part relationship. The typed `delete_part` helper is convenient for a known unique child.

## Settings element

The document settings part root is `w:settings`. It stores settings that apply to the WordprocessingML document, such as default tab stops or character spacing behavior:

```xml
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="dontCompress"/>
</w:settings>
```

Removing the settings part removes the part relationship and the part payload from the saved package. It does not rewrite document body content.

## Remove the settings part

```rust
{{#include ../../listings/getting-started/src/lib.rs:remove_settings_part}}
```

The function:

1. Opens a WordprocessingML package.
2. Gets the main document part.
3. Checks whether the optional settings part exists.
4. Deletes it from the main document part if present.
5. Saves the updated package to memory.

If the settings part is not present, the function leaves the package unchanged and still returns saved package bytes.

When several relationship IDs target the same Part, select the intended edge
with `delete_part_by_id`. A target Part does not own one relationship ID, and
removing one edge must not silently remove its remaining references.

## Optional parts

Many Open XML parts are optional. In `ooxmlsdk`, optional child-part accessors return `Option<T>`, so callers should handle both the present and absent cases explicitly.
