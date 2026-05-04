# Repository Guidelines

## Project Structure & Module Organization

This repository is an mdBook for `ooxmlsdk` documentation. Book source lives in `src/`, with navigation in `src/SUMMARY.md`. Topic folders mirror document domains: `src/general/`, `src/word/`, `src/spreadsheet/`, and `src/presentation/`. Static images are in `src/media/`. Tested Rust examples live under `listings/`; each listing is a Cargo workspace member and should be included into Markdown with mdBook `{{#include ...}}` anchors. Generated book output goes to `book/` and build artifacts to `target/`.

## Build, Test, and Development Commands

- `cargo fmt --all`: format all Rust listings.
- `cargo test --workspace`: compile and run every tested documentation listing.
- `cargo clippy --workspace --all-targets -- -D warnings`: lint Rust listings and fail on warnings.
- `mdbook build`: render the documentation into `book/`.
- `git diff --check`: catch trailing whitespace and patch formatting issues before commit.

Run the full set before committing changes that touch examples or Markdown includes.

## Coding Style & Naming Conventions

Rust code uses `rustfmt` with the repository settings in `.rustfmt.toml`. Prefer small, focused listing crates under `listings/<chapter-or-topic>/`; use clear function names such as `read_comments_part` or `replace_theme_part`. Keep Markdown headings sentence-like and specific. Avoid C#, Visual Basic, NuGet, or .NET-specific wording unless a page explicitly explains migration context.

## Testing Guidelines

All Rust shown in documentation must come from `listings/` and pass `cargo test --workspace`. Use mdBook anchors to expose only the documentation-ready portion of a Rust file; keep fixtures and tests outside the included anchor. If an upstream Open XML SDK scenario is not supported by `ooxmlsdk 0.6.0`, document the limitation instead of inventing an API.

## Reference Sources

Use upstream documentation and crate source to keep ports accurate. Prefer local checkouts when available: `../open-xml-docs/` for the source Microsoft documentation and `../ooxmlsdk/` for the Rust crate implementation and tests. If those paths are missing, use https://github.com/OfficeDev/open-xml-docs and https://github.com/KaiserY/ooxmlsdk. Treat upstream pages as source material; rewrite them into Rust-oriented `ooxmlsdk` documentation rather than translating line by line.

## Commit & Pull Request Guidelines

Recent commits use short imperative or conventional-style messages, for example `docs: port initial chapters to ooxmlsdk 0.6.0` and `add listings framework`. Prefer `docs: ...` for documentation ports and `listings: ...` for example infrastructure. Pull requests should summarize changed chapters, list added/updated listing crates, and include verification commands run.

## Agent-Specific Instructions

Start from local evidence. Use `rg` or `rg --files` first, then read only the files needed for the task. Keep summaries diff-based and avoid pasting large generated snippets or broad search output unless requested.

Run commands from the repository root. Cargo generation, formatting, testing, clippy, and bench commands must run sequentially in the default `target/` directory; do not set `CARGO_TARGET_DIR`. If Cargo reports a target lock, wait for Cargo rather than probing processes.

Preserve user work in the tree. Do not revert unrelated changes. Keep licensing context in `src/preface.md` intact.
