# Create a package

Open XML files are ZIP-based packages. A valid package needs content types, relationships, and at least the required root part for the document category you are creating.

In `ooxmlsdk 0.6.0`, package read/write APIs are available through types such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument`. The crate can add parts and save packages, but it does not currently expose a high-level convenience constructor equivalent to "create a blank `.docx` from a path".

## Current Rust workflow

For now, the recommended documented workflow is:

1. Start from an existing valid package or template file.
2. Open it with the appropriate package type.
3. Modify parts and root elements through `ooxmlsdk`.
4. Save the package to a writer or file.

The getting-started example demonstrates the open, inspect, and save path using `WordprocessingDocument`.

```rust
{{#include ../../listings/getting-started/src/lib.rs:full_example}}
```

## Creating from scratch

Creating a package from scratch is possible at the file-format level, but the initial OPC seed package must contain valid `[Content_Types].xml` and relationship parts before `ooxmlsdk` can open it as a package.

Until a higher-level creation helper is available in the crate, prefer using a template package for documentation examples. This keeps examples focused on `ooxmlsdk` APIs instead of teaching raw ZIP and OPC bootstrapping.

## WordprocessingML structure

The minimum main document part for a word-processing package is a `w:document` root element with a `w:body` child:

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>
```

The corresponding generated Rust types live under:

`ooxmlsdk::schemas::schemas_openxmlformats_org_wordprocessingml_2006_main`

For more about WordprocessingML package structure, see [Structure of a WordprocessingML document](../word/structure-of-a-wordprocessingml-document.md).
