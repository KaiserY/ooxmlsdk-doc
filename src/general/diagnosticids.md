# Diagnostic IDs

This page originally documented compiler diagnostic IDs used by the upstream SDK.

`ooxmlsdk 0.6.0` does not currently define custom Rust diagnostic IDs for obsolete or experimental APIs. Rust users should rely on the normal Rust compiler diagnostics, Cargo feature errors, crate documentation, and `ooxmlsdk::common::SdkError` values returned at runtime.

## Common Rust diagnostics

Typical issues while using `ooxmlsdk` include:

- Missing Cargo features, such as using `ooxmlsdk::parts` without enabling `parts`.
- Calling an MCE process mode without enabling the `mce` feature.
- Treating optional parts as always present instead of handling `Option`.
- Ignoring `Result` from package open, parse, write, and save operations.

## Runtime errors

Package operations return `Result` and may fail because:

- The input file is not a valid ZIP package.
- `[Content_Types].xml` is missing or invalid.
- A required relationship target is missing.
- XML does not match the generated schema type being parsed.
- Output writing fails.

Handle these cases with normal Rust error propagation:

```rust
fn open_docx(path: &std::path::Path) -> Result<(), Box<dyn std::error::Error>> {
  let _document = ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument::new_from_file(path)?;
  Ok(())
}
```
