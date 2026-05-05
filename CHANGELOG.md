# Changelog

All notable changes to OfficeWeaver are documented here.

## 0.2.0 - 2026-05-06

### Added

- Added ONLYOFFICE Word/text-document support through the new `Doc.*` helper namespace.
- Added `buildTextDocumentMacroCommand` and `buildWordMacroCommand` for text-document macros.
- Added traced Word helpers for paragraphs, headings, text styling, search/replace, lists, and tables.
- Added Word RAG manuals under `manuals/onlyoffice/9.3/word/`.
- Added helper-first spreadsheet manuals for values, formulas, formatting, alignment, borders, rows, columns, sheets, and freeze panes.
- Added Korean and Japanese README coverage for the Word/text-document workflow.

### Changed

- Updated the package description and manuals manifest for `0.2.0`.
- Changed RAG guidance so common mapped operations prefer `Sheet.*` and `Doc.*` helpers instead of raw ONLYOFFICE API calls.
- Reworked spreadsheet manuals so helper-mapped operations no longer expose raw API examples.
- Kept raw API documentation only for operations that are not yet wrapped, such as advanced charts, conditional formatting, AutoFilter, table formatting, and range reads.
- Limited the documented ONLYOFFICE support target to the 9.3 version family.

### Removed

- Removed the placeholder ONLYOFFICE 9.4 adapter metadata.
- Removed helper-mapped raw spreadsheet manuals such as raw value, font/fill, number format, border, alignment, range/sheet, and freeze-pane pages.
- Removed helper-mapped raw Word manuals such as raw paragraph, range, table, list, and search/replace pages.

### Migration Notes

- If your app embeds OfficeWeaver manuals into a vector database, rebuild the OfficeWeaver RAG index after upgrading to `0.2.0`.
- Filter retrieved manuals by `engine`, `version_family`, and `kind`. For this release, the supported ONLYOFFICE version family is `9.3`, and supported kinds are `spreadsheet` and `word`.
- Generated spreadsheet macros should use `Sheet.*` helpers first.
- Generated Word/text-document macros should use `Doc.*` helpers first.
- Use `Sheet.raw(...)` or `Doc.raw(...)` only when the requested editor operation is not available as a helper.

## 0.1.0 - 2026-05-05

### Added

- Initial public npm release.
- Added the OfficeWeaver runtime for AI-friendly ONLYOFFICE spreadsheet automation.
- Added traced `Sheet.*` helpers for values, formulas, formatting, rows, columns, sheets, freeze panes, and raw escape-hatch operations.
- Added RAG-ready spreadsheet manuals for ONLYOFFICE 9.3.
- Added English, Korean, and Japanese README files.
