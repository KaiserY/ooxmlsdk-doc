# Spreadsheets

This section covers SpreadsheetML packages (`.xlsx`, `.xlsm`, `.xltx`) with `ooxmlsdk`.

Spreadsheet packages are made of a workbook part, worksheet parts, optional shared strings, styles, tables, charts, pivot caches, drawings, and relationships between those parts. In `ooxmlsdk`, the entry point is usually `ooxmlsdk::parts::spreadsheet_document::SpreadsheetDocument`.

Use the `parts` feature, enabled by default, to open and save packages. Examples in this section are backed by tested Rust code in `listings/spreadsheet`.

## In this section

- [Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md)
- [Open a spreadsheet document for read-only access](how-to-open-a-spreadsheet-document-for-read-only-access.md)
- [Retrieve a list of the worksheets in a spreadsheet document](how-to-retrieve-a-list-of-the-worksheets-in-a-spreadsheet.md)
- [Get worksheet information from a package](how-to-get-worksheet-information-from-a-package.md)
- [Retrieve the values of cells in a spreadsheet](how-to-retrieve-the-values-of-cells-in-a-spreadsheet.md)
- [Working with sheets](working-with-sheets.md)
- [Working with the shared string table](working-with-the-shared-string-table.md)
- [Working with formulas](working-with-formulas.md)
- [Working with the calculation chain](working-with-the-calculation-chain.md)
- [Working with conditional formatting](working-with-conditional-formatting.md)
- [Working with PivotTables](working-with-pivottables.md)
- [Working with SpreadsheetML tables](working-with-tables.md)

Writer-focused chapters are being ported only when the code has a fixture in `listings/` and passes `cargo test --workspace`.

## Related sections

- [Getting started](../getting-started.md)
- [General package operations](../general/overview.md)
