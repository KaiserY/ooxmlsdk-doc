# Migration notes

This documentation targets `ooxmlsdk 0.6.0`.

The imported Microsoft documentation originally included upstream SDK migration guidance such as v2.x to v3.0. Those notes do not apply to the Rust crate. For Rust users, the relevant migration concern in this documentation set is moving examples and manifests to `ooxmlsdk 0.6.0`.

## From 0.5.x to 0.6.0

Update your Cargo manifest:

```toml
[dependencies]
ooxmlsdk = "0.6.0"
```

The default feature set now enables `parts`, which keeps the package APIs available for most users.

If your manifest explicitly enabled the old `microsoft365` feature, remove it:

```toml
[dependencies]
ooxmlsdk = "0.6.0"
```

Newer generated schema and part coverage remains included in the runtime. It is no longer exposed as a public `microsoft365` feature switch.

## Explicit feature selection

For package APIs without extra feature behavior:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["parts"] }
```

For Flat OPC helpers:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["flat-opc"] }
```

For Markup Compatibility and Extensibility processing:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["mce"] }
```

For validation APIs:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", features = ["validators"] }
```

## Upstream v3 concepts in Rust

The upstream v3.0 migration guide describes architectural changes for .NET consumers. The same concepts usually appear differently in `ooxmlsdk`:

- Package APIs are already part of the Rust crate behind the `parts` feature. There is no separate framework package to add.
- Package and part constructors are Rust associated functions such as `new`, `new_with_settings`, `new_from_file`, and `new_from_file_with_settings`.
- Package lifetime follows Rust ownership. There is no explicit `Close` method; let values drop after saving or after all reads are complete.
- Part creation uses typed methods and generated relationship metadata, including helpers such as `add_new_part`, `add_image_part`, `add_custom_xml_part`, and content-type-specific variants where supported.
- Runtime configuration uses Cargo features plus `OpenSettings`. There is no dynamic feature-collection API equivalent.
- Markup Compatibility and Extensibility processing is feature gated by `mce`. Enable it only for examples or applications that need compatibility processing while loading package roots.

## Documentation migration

When porting a page from the original upstream documentation:

- Remove examples in other programming languages.
- Replace upstream package names with Rust crate paths such as `ooxmlsdk::parts`, `ooxmlsdk::schemas`, and `ooxmlsdk::sdk`.
- Prefer examples stored under `listings/` and included into Markdown with mdBook include directives.
- Keep each example covered by `cargo test --workspace`.
- Use Rust terms such as `Result`, `Option`, iterators, borrowed paths, and Cargo features.
