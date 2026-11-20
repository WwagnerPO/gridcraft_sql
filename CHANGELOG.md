# Changelog

Changes of this fork (GridCraft SQL+VBA) on top of upstream GridCraft. Upstream's own changes are
in its [releases](https://github.com/storytold/gridcraft/releases). The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions are named after the upstream
release the fork is based on.

## [Unreleased] — based on upstream v0.19.1

### Changed
- Merged upstream v0.14.2, v0.15.0, v0.16.0, v0.16.1, v0.17.0, v0.18.0, v0.19.0 and v0.19.1,
  one merge per release, `cargo xtask ci` green after each.
- The VBA interpreter stays separate from upstream's new macro engine (command lists, 0.14) and
  shares its document trust; see [`docs/vba.md`](docs/vba.md).
- The VBA macros dialog is now `vbaMacros`: Alt+F8 opens it in a workbook with a VBA project
  (with a button to GridCraft's macros); every other way to Macros opens upstream's dialog.
- Form buttons that run a VBA macro are upstream's form controls (0.17): `control.press` and a
  click in the window run the VBA macro, with `Application.Caller` set to the button.
- The fork's VBA documentation moved from `docs/macros.md` (now upstream's) to `docs/vba.md`.

### Fixed
- Form buttons assigned to a VBA macro stopped working after upstream 0.17 read form controls
  itself ("There is no macro called …"); they run again.
- The SQL Server query's table hooks (rename, convert to range) were in a file upstream 0.14.2
  no longer compiles; moved where they run.

### Removed
- The ArtCraft brand files (`docs/brand/`): their licence asks forks to remove them.
- The fork's own reader and painter of VML form buttons (`xlsx/form_controls.rs`,
  `ui-egui/controls.rs` painting): upstream reads and draws them.

## [0.13.0-sql] — 2026-10-30, based on upstream v0.13.0

### Added
- SQL Server connections saved on this computer as Office Data Connection files (`.odc`) in
  `Documents/My Data Sources`: Export Connection File, Get Data › Existing Connections, and
  `data.exportConnectionFile`, `data.existingConnections`, `data.fromConnectionFile`.
- VBA events `Worksheet_Change` and `Workbook_SheetChange`.
- VBA sequential text files: `Open … For Input | Output | Append`, `Print #`, `Write #`,
  `Input #`, `Line Input #`, `Close`, `Reset`, `EOF`, `LOF`, `FreeFile`.
- `vba.trust` and `trusted: true` on `vba.run` / `vba.click`.
- `docs/vba.md` (then `docs/macros.md`).

### Changed
- Macros of a workbook opened from a file run only once enabled, on every path (Alt+F8, form
  buttons, CLI, MCP, control channel); before, `vba.run` and `vba.click` ran them unasked.
- `NOW()`, `TODAY()`, Ctrl+;, "today" conditional formats and the `&[Date]` header field read
  the local time instead of UTC.

### Fixed
- Sheet code names (`<sheetPr codeName>`) were dropped on save.
- `Workbook_Open` only ran from a module called `ThisWorkbook` (not `EsteLibro`…).
- Macros could crash or hang the app: deep nesting overflowed the stack, `Like` with several
  `*` backtracked exponentially, huge strings and arrays exhausted memory, and a panic on the
  macro thread lost the workbook.
- `data.queriesConnections` reported the source table and the worksheet table under the same key.

### Security
- `webpki-root-certs` 1.0.11 (RUSTSEC-2026-0383).

[Unreleased]: https://github.com/WwagnerPO/gridcraft_sql/compare/848e4584...feat/merge-upstream-0.19.1
[0.13.0-sql]: https://github.com/WwagnerPO/gridcraft_sql/compare/39d60d1e...848e4584
