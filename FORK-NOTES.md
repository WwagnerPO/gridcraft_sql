# Fork notes

How this fork (GridCraft SQL+VBA, `WwagnerPO/gridcraft_sql`) relates to upstream
(`storytold/gridcraft`), and how to carry its changes onto a new upstream base. Day-to-day rules
and the table of hook lines in upstream files are in [`FORK.md`](FORK.md) (Spanish); the user-facing
summary is in the [README](README.md#what-this-fork-adds) and [`CHANGELOG.md`](CHANGELOG.md).

Current base: **upstream v0.19.1** (merged 2026-11-20 on `feat/merge-upstream-0.19.1`).

## Commits of the fork

Everything else in the history is upstream's, or a merge of an upstream release.

### SQL Server (13–14 Oct 2026)

| Commit | What |
|---|---|
| `53917334` | Get Data from SQL Server Database (Navigator, preview, load, refresh) |
| `6e493ded` | Plain reasons for refused server certificates; trust the certificate and connect |
| `0269d191` | Write to SQL Server: Save Changes, Export, Run SQL |
| `b49973d9` | Windows sign-in (current user and another account) |
| `ca150cee` | Remember database logins in the system credential store |
| `0b221326` | Excel Power Query over SQL Server opens as a refreshable query |

### VBA (14–15 Oct 2026)

| Commit | What |
|---|---|
| `e3bb1148` | Read VBA projects (MS-OVBA) and list their modules |
| `f85accee` | VBA language interpreter (`gridcraft-vba` crate) |
| `3a757b31` | Excel object model for VBA macros |
| `9f7f15d1` | Run VBA macros from the window |
| `23ec3f62` | Macro workbooks open with the security bar |
| `241c31cc` | Form buttons run their macros; `Application.Caller`, Shapes, threaded comments, `ADODB.Stream` |

### Fork plumbing (15 Oct 2026)

| Commit | What |
|---|---|
| `58c53994` | Keep the database and macro code out of shared files (hook lines only) |
| `b6192e07` | Cargo.lock follows the release for the vba and sqlserver crates |
| `e53e3f1d` | GitHub workflow that builds and publishes the Windows MSI (`fork-windows.yml`) |
| `8556a613`, `4e985ce6`, `253dde25` | `FORK.md` |

### 30 Oct 2026 (on upstream v0.13.0)

| Commit | What |
|---|---|
| `b6dedf4d` | Connections saved on this computer (`.odc`, My Data Sources) |
| `7b34d625` | Harden macros against hostile code; fix the query report |
| `9dd82325` | `webpki-root-certs` 1.0.11 (RUSTSEC-2026-0383) |
| `a79427c8` | Macros from a file run only once enabled, on every path |
| `d779c587` | VBA events `Worksheet_Change`, `Workbook_SheetChange`; keep sheet code names |
| `62c4ef5a` | VBA text files (`Open`, `Print #`, `Write #`, `Input #`, `Line Input #`, `Close`) |
| `848e4584` | `NOW()` and `TODAY()` read the local time, not UTC |

### Sync to v0.19.1 (20 Nov 2026)

| Commit | What |
|---|---|
| `99993db9` | Merge v0.14.2 (VBA dialog becomes `vbaMacros`; `docs/macros.md` → `docs/vba.md`; table hooks to `insert.rs`) |
| `5b0961ac` | Decision: keep the VBA interpreter beside upstream's macro engine |
| merges | v0.15.0, v0.16.0, v0.16.1, v0.17.0, v0.18.0, v0.19.0, v0.19.1 (each message lists its conflicts) |
| in the v0.17.0 merge | VBA form buttons rebuilt on upstream's form controls |
| `54414767` | Cargo.lock follows v0.16.1 for the vba and sqlserver crates |
| `228ace1b`, `1115b354`, this file | README, CHANGELOG, fork notes |

## Two fixes that belong upstream

They are upstream bugs fixed here; reporting them upstream would remove two hook points:

- `NOW()` / `TODAY()` read UTC: `crates/core/src/date.rs` (`local_now_ms`) and
  `crates/calc/src/recalc.rs` (`now_serial`).
- Sheet code names dropped on save: `crates/model/src/sheet.rs` (`code_name`),
  `crates/xlsx/src/sheet_read.rs` and `sheet_write.rs` (`<sheetPr codeName>`).

## Re-applying the fork on a new upstream base

The usual way is a merge (see the policy in the README): `rerere` is on and remembers how each
conflict was resolved. If the history ever has to be rebuilt on a fresh upstream (a rebase, or a
new repository), don't cherry-pick 40 commits one by one; carry the result instead:

1. Start a branch at the new upstream tag: `git checkout -b fork-on-X.Y.Z vX.Y.Z`.
2. Copy the fork's own files from the current fork (they never conflict, upstream doesn't have
   them): the "Archivos propios" table of [`FORK.md`](FORK.md), plus `docs/vba.md`,
   `docs/sql-server.md`, `README.md`, `CHANGELOG.md`, `FORK-NOTES.md`, `FORK.md`,
   `PENDIENTES.md`, `.github/workflows/fork-windows.yml`:
   `git checkout <old-fork> -- crates/vba crates/sqlserver …`
3. Re-add each hook line of the "Archivos compartidos" table of `FORK.md` in the upstream files.
   `git diff <old-upstream-base> <old-fork> -- <file>` shows exactly which lines are the fork's.
   Upstream often moves code between files (`cmd/insert.rs` ↔ `cmd/insert/*.rs`, `dialogs.rs` →
   `dialogs/show.rs`): put the hook where the function lives now, and check no hook landed in a
   file that isn't compiled (a file without a `mod` line).
4. `cargo update -w`, `cargo xtask keytips`, then `cargo xtask ci` with `CRAFT_FONTS_DIR` set to a
   `storytold/craft-fonts` checkout (two UI tests need its fonts).
5. Check the fork's features still pass their tests:

   ```sh
   cargo test -p gridcraft-xlsx odc          # .odc files
   cargo test -p gridcraft-xlsx code_names   # sheet code names round trip
   cargo test -p gridcraft-engine connection # SQL Server queries and .odc commands
   cargo test -p gridcraft-engine vba_host   # events, text files, trust, buttons
   cargo test -p gridcraft-vba               # the interpreter, hostile code
   cargo test -p gridcraft-core local_now    # NOW()/TODAY() local time
   cargo test -p gridcraft-ui-egui --test macros --test get_data
   ```

## Open TODOs

- **TODO** Report the two upstream bugs above to `storytold/gridcraft`.
- **TODO** The ArtCraft brand files in `docs/brand/` are still in the repository (only the README
  stopped showing them). Their licence asks forks to remove them; `xtask/src/assets.rs` checks
  them, so removing them needs that check (and `ATTRIBUTION.md`) adjusted. Decide and do it.
- **TODO** Upstream `main` is past v0.19.1; the next sync is the next upstream minor release.
- Upstream v0.14.2 left `crates/engine/src/cmd/insert/*.rs` uncompiled (no `mod` line): not the
  fork's doing, nothing to fix here, but don't put hooks there.
