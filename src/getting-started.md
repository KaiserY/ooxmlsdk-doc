# Getting started with the Open XML SDK

The Open XML SDK simplifies the task of manipulating Open XML packages and the underlying Open XML schema elements within a package. The classes in the Open XML SDK encapsulate many common tasks that developers perform on Open XML packages, so that you can perform complex operations with just a few lines of code.

## Rust package

Add `ooxmlsdk` to your Cargo project:

```toml
[dependencies]
ooxmlsdk = "0.5.1"
```

The documentation examples in this book are backed by real Rust files under `listings/` and are checked with `cargo test --workspace`.

For example, this function creates a minimal WordprocessingML package, adds the main document part, and writes the package to memory:

```rust
{{#include ../listings/getting-started/src/lib.rs:full_example}}
```

## Available packages

The SDK is available as a collection of NuGet packages that support .NET 3.5+, .NET Standard 2.0, .NET 6+, and [other supported platforms](https://learn.microsoft.com/dotnet/standard/net-standard) for those targets. For information about installing packages, please see [the NuGet documentation](https://learn.microsoft.com/nuget/quickstart/install-and-use-a-package-in-visual-studio). The following are the available packages:

- [`DocumentFormat.OpenXml.Framework`](https://www.nuget.org/packages/DocumentFormat.OpenXml.Framework): This package contains the foundational framework that enables the SDK. This is a new package starting with v3.0 and contains many types that previously were included in `DocumentFormat.OpenXml`.
- [`DocumentFormat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml): This package contains all of the strongly typed classes for parts and elements.
- [`DocumentFormat.OpenXml.Features`](https://www.nuget.org/packages/DocumentFormat.OpenXml.Features): This package contains additional functionality that enables some opt-in features.
- [`DocumentFormat.OpenXml.Linq`](https://www.nuget.org/packages/DocumentFormat.OpenXml.Linq): This package contains a collection of all the fully qualified names for parts and elements to enable more efficient `Linq` usage.
