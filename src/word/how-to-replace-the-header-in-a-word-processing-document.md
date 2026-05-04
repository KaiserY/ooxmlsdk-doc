# Replace the header in a word processing document

Headers are stored in separate header parts and referenced from section properties. Replacing a header can mean editing the existing header part or creating a new part and updating the section reference.

## Rust workflow

Use the main document part to locate section references:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a header writer. A safe implementation should preserve relationship ids where possible and handle default, first-page, and even-page headers separately.
