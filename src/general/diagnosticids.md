# Diagnostic IDs

This page originally documented compiler diagnostic IDs used by the upstream SDK.

`ooxmlsdk` does not currently define custom Rust diagnostic IDs for obsolete or experimental APIs. Rust users should rely on the normal Rust compiler diagnostics, Cargo feature errors, crate documentation, and `ooxmlsdk::common::SdkError` values returned at runtime.

The upstream `OOXML0001` diagnostic is specific to an experimental .NET package abstraction and does not apply to this Rust crate.

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

The 0.13 package-bound Part model also reports ownership and relationship
failures explicitly:

| Error | Meaning |
| --- | --- |
| `ForeignPart` | The Part handle belongs to another package |
| `StalePart` | The Part has been deleted from its package |
| `PartNotReferenced` | The selected relationship source does not reference the Part |
| `PartRelationshipNotFound` | The requested relationship ID is absent |
| `AmbiguousPartRelationship` | An operation requiring one edge found several IDs targeting the Part |

Handle these cases with normal Rust error propagation, usually by returning `Result` from your own helper and using `?` on package operations.

## Validator errors

Where validator APIs are available, they return `ValidationErrorInfo` values. Those values include a validation category and, where available, an ID derived from the failed validator, such as `required`, `enum`, or `field_value`. Treat these as runtime validation data rather than compiler diagnostics.
