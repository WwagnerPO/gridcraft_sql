# VBA macros

> **Last reviewed:** 2026-10-30 · **Last updated:** 2026-10-30 · **Change:** major (created) · **Target:** Microsoft Excel (Microsoft 365)

GridCraft reads the VBA project of a macro-enabled workbook (`.xlsm`, `vbaProject.bin`,
[MS-OVBA]) and runs its macros with a clean-room interpreter of the language ([MS-VBAL]) and an
Excel object model (`Range`, `Cells`, `Worksheets`, `Application`, `MsgBox`, `InputBox`…). The
project is kept unchanged on save.

## Security

As in Excel, a workbook opened from a file (or holding content pasted from one) starts with its
macros **disabled**: the security bar says so and **Enable Content** turns them on for this
session, then runs `Workbook_Open` (or `Auto_Open`). A workbook made in this session can run its
macros at once. Disabled macros don't run from Alt+F8, form buttons, the CLI, MCP or the control
channel either: `vba.run` and `vba.click` refuse with "Macros in this workbook are disabled" until
`vba.trust` (or `trusted: true` on the call).

Enabled macros can do what Excel's can: change the workbook, read and write files (`Dir`, `Kill`,
`FileCopy`, `MkDir`, `RmDir`, `ADODB.Stream`), read environment variables and use the database
connections. Enable content only for workbooks whose origin you trust.

## In the app

- **Alt+F8** (Developer › Macros) lists the runnable macros (public Subs without arguments) and
  runs one. Message and input boxes appear as dialogs; **Stop** interrupts a running macro.
- Form-control buttons on a sheet run their macro when clicked (`Application.Caller` is the
  button).
- A macro's changes undo as one step.

## Commands (CLI, MCP, control channel)

| Command | Params |
|---|---|
| `vba.macros` | `{}` → the runnable macros |
| `vba.run` | `{macro, answers?: [button codes], inputs?: [text], trusted?}` → `{result, messages, output}` |
| `vba.click` | `{button, sheet?, answers?, inputs?, trusted?}` runs a form button's macro |
| `vba.buttons` | `{sheet?}` → the sheet's form buttons |
| `vba.trust` | `{trusted?: true \| false}` enables (or disables) the active workbook's macros |

Message boxes answer their first button (or the next of `answers`), input boxes their default
(or the next of `inputs`).

## Limits

- Runs the source, not the compiled p-code; `Declare` (Windows API calls), `GoSub`, `AddressOf`,
  UserForms and file statements (`Open … As #n`, `Print #`, `Line Input`…) aren't supported yet:
  a macro using them stops with a message naming what it used.
- Only `Workbook_Open` / `Auto_Open` run as events.
- Hostile code is stopped, never the app: nesting deeper than 128 levels is "Expression too
  complex", runaway recursion "Out of stack space", strings past 64M characters and arrays past
  16M elements "Out of memory", and a run of more than 500 million steps is stopped.

[MS-OVBA]: https://learn.microsoft.com/openspecs/office_file_formats/ms-ovba/
[MS-VBAL]: https://learn.microsoft.com/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-30 | major | Created: security (Enable Content, `vba.trust`), commands, limits |
