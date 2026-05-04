# Remove a document part from a package

This example removes the Wordprocessing settings part from a `.docx` package.

In Open XML, removing a child part means removing the relationship from the parent part and marking the target part as deleted in the package model. `ooxmlsdk` handles that through `delete_part`.

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

## Optional parts

Many Open XML parts are optional. In `ooxmlsdk`, optional child-part accessors return `Option<T>`, so callers should handle both the present and absent cases explicitly.
