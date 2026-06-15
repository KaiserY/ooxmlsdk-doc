# Validate a word processing document

Validation can mean different things: package relationships resolve, required parts exist, XML is well-formed, and the content follows WordprocessingML schema rules. In `ooxmlsdk` 0.10.2, schema validation is available through the optional `validators` feature.

The upstream Open XML SDK sample separates two cases: validating a normal document and validating a deliberately corrupted document that contains schema-invalid content. In Rust, keep the same distinction: a valid document should return an empty diagnostics list, and an invalid document should return structured `ValidationErrorInfo` values that the caller can inspect or print.

## Validate the package

Enable the feature in `Cargo.toml`:

```toml
[dependencies]
ooxmlsdk = { version = "0.10.2", features = ["validators"] }
```

Then open the package and call `validate`. The package-level validator loads the known root elements for the document and reports schema diagnostics with part context when available.

```rust
{{#include ../../listings/word/src/lib.rs:validate_word_document}}
```

An empty vector means no validation errors were reported by the generated validators:

```rust
let errors = validate_word_document(path)?;
if errors.is_empty() {
  println!("document is valid");
}
```

Each `ValidationErrorInfo` includes an error category, a readable description, and optional fields such as `id`, `type_name`, `field_name`, and `part_uri`. Treat those diagnostics as runtime validation data; do not rely on opening a package alone as a substitute for schema validation.
