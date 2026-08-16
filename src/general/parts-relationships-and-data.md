# Parts, relationships, and data ownership

An Open XML document is an Open Packaging Conventions package. The package
contains parts, and relationship edges connect the package and its parts. The
main package relationship determines whether the document is WordprocessingML,
SpreadsheetML, or PresentationML. `ooxmlsdk` keeps that structure explicit
instead of presenting a Word-, Excel-, or PowerPoint-style object model.

## Package-bound Part handles

The document package owns storage. Typed values such as `MainDocumentPart`,
`WorksheetPart`, and `SlidePart` are lightweight handles into that storage.
Package types are not `Clone`; Part handles are cheap to clone, but every handle
remains bound to the package that created it.

Pass the package whenever a Part operation needs to resolve storage:

- read metadata or payloads with `&document`;
- change payloads, relationships, or typed roots with `&mut document`;
- use `add_part_from_package` when copying a Part between packages instead of
  passing a source handle to the destination package.

Using a handle with the wrong package returns `SdkError::ForeignPart`. A handle
to a Part that has since been deleted returns `SdkError::StalePart`. These
checks replace the public numeric Part IDs used by older releases.

## Relationship IDs identify edges

A relationship ID belongs to its source relationship set, not to the target
Part. Two different parents can use different IDs for the same target, and one
parent can contain multiple relationship edges to that target.

Use a `RelatedPart<T>` iterator when the edge identity matters. It keeps the
relationship ID and type beside the typed target Part:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheet_relationship_ids}}
```

Choose an API according to the identity you have:

| Need | API shape |
| --- | --- |
| Typed targets only | Generated plural accessors or `get_parts_of_type` |
| Every matching edge and ID | `related_parts_of_type` or another plural `RelatedPart` API |
| Target for a known ID | `get_part_by_id` or `try_get_part_by_id` |
| First ID that targets a Part | `get_id_of_part` |
| Mutate one specific edge | An explicit relationship-ID API such as `change_relationship_id` or `delete_part_by_id` |

`get_id_of_part` returns the first matching ID in source order and reports
foreign, stale, or unreferenced Parts through `Result`. Do not use it to imply
that a target can have only one relationship ID. Operations that require a
unique target relationship, such as `change_id_of_part`, return
`SdkError::AmbiguousPartRelationship` when several edges match.

## Borrowed, shared, and copied payloads

For short-lived inspection, `try_data` returns a borrowed byte slice and
`data_as_str` validates UTF-8 text. When payload ownership must outlive the
package borrow, use `try_data_bytes`:

```rust
{{#include ../../listings/getting-started/src/lib.rs:read_main_part_bytes}}
```

`try_data_bytes` returns `bytes::Bytes`. `Bytes` comes from the
[`bytes`](https://crates.io/crates/bytes) crate rather than `std`; cloning it
shares the immutable payload instead of copying the contents. Add `bytes = "1"`
as a direct dependency when an application names `Bytes` in its own public
types.

Use the other payload APIs deliberately:

| API | Behavior |
| --- | --- |
| `try_data` | Borrow the payload and preserve package/Part errors |
| `try_data_bytes` | Own a shared immutable view; cheap to clone |
| `data_to_vec` | Copy the payload into a new `Vec<u8>` |
| `data_as_str` | Borrow the payload as validated UTF-8 |
| `write_data_to` | Stream the current payload into a `Write` target |
| `set_data` / `feed_data` | Replace the payload and invalidate any cached typed root |

## Lazy typed roots and saving

`PackageOpenMode::Lazy` is the default. Package metadata and relationships are
available after open, while a generated XML root is parsed when
`root_element(&document)` first requests it. Use
`root_element_mut(&mut document)` for schema-aware changes, or
`set_root_element` to replace the root.

Saving follows the state actually observed by the application:

- an untouched lazy Part can reuse its original payload;
- once a typed root has been loaded, saving serializes that root;
- `set_data` and `feed_data` replace raw bytes and unload the cached root.

This makes a load-save-reopen cycle exercise the generated parser and
serializer instead of silently copying source XML after it has been parsed.

## Reader paths

Package constructors such as `WordprocessingDocument::new` accept a
`Read + Seek` source by value because ZIP package access requires seeking. The
generic reader is consumed during open and copied into the package's shared
in-memory archive backing; it is not the streaming XML path. Path constructors
manage a seekable file backing internally.

Generated root schema types separately expose `SdkType::from_bytes` for a
borrowed byte slice and `SdkType::from_reader` for a `BufRead` source. The
reader form parses from the supplied stream; it does not first convert the
entire XML input into a `Vec<u8>`.
