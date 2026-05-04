# Welcome to ooxmlsdk

`ooxmlsdk` is a Rust library for reading, writing, and round-tripping Office Open XML packages such as `.docx`, `.xlsx`, and `.pptx`.

Office Open XML is standardized by [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) and [ISO/IEC 29500](https://www.iso.org/standard/71691.html). The file formats are useful for Rust applications because they are ZIP packages containing XML parts and explicit relationships, rather than opaque application files.

The crate provides generated schema types, serializers, deserializers, and strongly typed package parts. Its package model follows the same container concepts as the upstream SDK, but the API is Rust-native: methods return `Result`, package and part types are regular Rust structs, and optional crate functionality is controlled by Cargo features.

## Start here

- [Getting started](getting-started.md)
- [About ooxmlsdk](about-the-open-xml-sdk.md)
- [What's new in ooxmlsdk](what-s-new-in-the-open-xml-sdk.md)
- [Design considerations](open-xml-sdk-design-considerations.md)
- [Migration notes](migration/migrate-v2-to-v3.md)

## Working with packages

- [General package APIs](general/overview.md)
- [Presentations](presentation/overview.md)
- [Spreadsheets](spreadsheet/overview.md)
- [Word processing](word/overview.md)

## References

- [`ooxmlsdk` on crates.io](https://crates.io/crates/ooxmlsdk)
- [`ooxmlsdk` API documentation](https://docs.rs/ooxmlsdk)
- [Open XML SDK upstream metadata and behavior reference](https://github.com/dotnet/Open-XML-SDK)

<sup>1</sup> Some chapters discuss concepts from ISO/IEC 29500. When text is still derived from the imported Microsoft documentation baseline, the notices in [Preface](preface.md) apply.
