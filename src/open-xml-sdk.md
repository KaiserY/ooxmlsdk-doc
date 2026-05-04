# Welcome to ooxmlsdk

`ooxmlsdk` is a Rust library for reading, writing, and round-tripping Office Open XML packages such as `.docx`, `.xlsx`, and `.pptx`.

Office Open XML is standardized by [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) and [ISO/IEC 29500](https://www.iso.org/standard/71691.html). The file formats are ZIP packages containing XML parts and explicit relationships, which makes them suitable for Rust tooling that needs deterministic package inspection or transformation.

The crate provides generated schema types, serializers, deserializers, and strongly typed package parts. Its API is Rust-native: methods return `Result`, package and part types are regular Rust structs, and optional functionality is controlled by Cargo features.

`ooxmlsdk` works at both layers of the format: the package graph and the XML schema data. You can open an existing document package, follow typed relationships to well-known parts, read or replace part data, load generated root elements, and save the package back to a writer.

## Start here

- [Getting started](getting-started.md)
- [About ooxmlsdk](about-the-open-xml-sdk.md)
- [What's new in ooxmlsdk](what-s-new-in-the-open-xml-sdk.md)
- [Design considerations](open-xml-sdk-design-considerations.md)

## Migrating

- [Migration notes](migration/migrate-v2-to-v3.md)

## Working with packages

- [General package APIs](general/overview.md)
- [Presentations](presentation/overview.md)
- [Spreadsheets](spreadsheet/overview.md)
- [Word processing](word/overview.md)

## References

- [`ooxmlsdk` on crates.io](https://crates.io/crates/ooxmlsdk)
- [`ooxmlsdk` API documentation](https://docs.rs/ooxmlsdk)
- [`ooxmlsdk` source repository](https://github.com/KaiserY/ooxmlsdk)
