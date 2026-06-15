# ooxmlsdk documentation

This repository contains the mdBook source for the `ooxmlsdk` documentation.

Published documentation: <https://kaisery.github.io/ooxmlsdk-doc/>

Related links:

- `ooxmlsdk` source: <https://github.com/KaiserY/ooxmlsdk>
- crates.io: <https://crates.io/crates/ooxmlsdk>
- API docs: <https://docs.rs/ooxmlsdk/latest/ooxmlsdk/>

The book currently targets `ooxmlsdk` 0.10.2. Unless a page explicitly names another version, examples and API notes refer to 0.10.2.

## Build

```bash
mdbook build
```

Rust listings used by the book are kept under `listings/` and can be checked with:

```bash
cargo test --workspace
```

## License

New `ooxmlsdk` documentation and Rust examples in this repository are licensed under MIT OR Apache-2.0.
