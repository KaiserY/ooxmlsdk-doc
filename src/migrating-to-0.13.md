# Migrating to ooxmlsdk 0.13

Versions 0.11 through 0.13 tightened the generated schema model and replaced
numeric package identities with package-bound typed Part handles. Most typed
child accessor names and the public XML entry points remain unchanged, but
applications upgrading from 0.10.x should review the following areas.

## Package and Part ownership

- `WordprocessingDocument`, `SpreadsheetDocument`, and
  `PresentationDocument` are no longer `Clone`. Keep one package owner and pass
  shared or mutable references to helpers.
- Typed Part handles remain cloneable, but they can only be resolved against
  their owning package.
- Public numeric `PartId`, low-level `SdkPart` Part-ID methods,
  `MediaDataPart::part_id`, and relationship `target_part_id` accessors were
  removed. Navigate with typed Parts and relationship APIs instead.
- Handle misuse now reports structured errors such as `ForeignPart`,
  `StalePart`, `PartNotReferenced`, `PartRelationshipNotFound`, and
  `AmbiguousPartRelationship`.

See [Parts, relationships, and data ownership](general/parts-relationships-and-data.md)
for the current model and tested examples.

## Relationship identity

A target Part does not own one relationship ID. Use plural `RelatedPart` APIs
when every edge matters, and use an explicit relationship-ID API when changing
or deleting a particular edge. `get_id_of_part` now returns `Result<&str,
SdkError>` and selects the first matching edge in source order.

## Payloads and typed roots

- `try_data_bytes` returns a shared owned `bytes::Bytes` payload without making
  every caller copy into a new `Vec<u8>`.
- Typed roots are lazy by default. Untouched Parts can retain their original
  payload, while loaded roots are serialized when the package is saved.
- Replacing raw payload data unloads the cached typed root so raw and typed
  representations cannot silently diverge.

## Generated schema changes

- Some required child fields and small choice payloads are now stored inline
  rather than behind `Box`.
- Generated root structs no longer contain an `xml_header` field. Use the
  existing `SdkType::write_to` or `to_xml` entry points; root metadata controls
  the XML declaration.
- Generated schema types no longer implement the redundant `AsRef<Self>`.
  Borrow them directly with `&value`.
- Catch-all variants and dynamic attribute storage were removed where they were
  not declared by the static schema metadata. Use the generated typed fields
  and schema-declared wildcard content that remains available.
- Validation is behind the `validators` feature. Generated schema types and
  package types expose inherent `validate` methods; there is no generated
  `SdkValidator` trait to import.

## Namespaces and MCE

The 0.13 reader recognizes supported namespace aliases, including known Strict
and Transitional URIs and non-canonical prefixes. The writer emits canonical
OOXML prefixes. Prefixed attributes and MCE namespace lists are resolved by
namespace identity rather than by assuming a literal source prefix.

Supported `mc:AlternateContent` positions can be read into typed schema fields.
Enable the `mce` feature when the application also needs active compatibility
branch selection and filtering for a target Office version.

## Upgrade checklist

1. Change the dependency to `ooxmlsdk = "0.13.0"` and keep only the features
   the application uses.
2. Remove package clones and numeric Part-ID plumbing.
3. Replace target-owned relationship assumptions with `RelatedPart` or
   relationship-ID APIs.
4. Choose borrowed bytes, shared `Bytes`, or copied `Vec<u8>` intentionally.
5. Update generated schema construction for inline fields and removed
   catch-alls.
6. Run save-reopen tests for representative `.docx`, `.xlsx`, and `.pptx`
   packages.
