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

## Text files

`Open path For Input | Output | Append As #n` (with `FreeFile`), `Print #`, `Write #`, `Input #`,
`Line Input #`, `Close`, `Reset`, `EOF` and `LOF` read and write text files as VBA does: `Print #`
puts a space before and after numbers and moves to the next 14-column zone after a comma; `Write #`
quotes strings and writes `#TRUE#`, `#NULL#`, `#2026-10-30#`. Files are written in ANSI
(Windows-1252) with `\r\n` line ends; reading takes UTF-8 or UTF-16 when the file is, ANSI
otherwise. A file opened for output is written when it is closed, or when the macro ends.

## Events

When macros are enabled:

- `Workbook_Open` (or `Auto_Open`) runs on Enable Content.
- `Worksheet_Change(ByVal Target As Range)` in a sheet's code module runs when cells of that sheet
  change (typing, paste, fill, commands…), with `Target` the changed cells (their bounding range
  when several areas changed).
- `Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)` in the workbook's module runs
  for a change on any sheet.

The modules are found by the sheet's code name (`Hoja1`, `Sheet1`…, kept with the file) and the
workbook module whatever its language (`ThisWorkbook`, `EsteLibro`…). In the app an event runs
like a macro, with real message boxes; from the CLI, MCP or the control channel it runs at the end
of the command, message boxes answering their first button. Changes a macro makes raise no events
(as with `Application.EnableEvents = False` in Excel), so a handler that writes cells can't loop.

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
  UserForms, and binary or random files (`For Binary`, `For Random`, `Get`, `Put`, `Seek`) aren't
  supported yet: a macro using them stops with a message naming what it used.
- Other events (`Worksheet_SelectionChange`, `Worksheet_Activate`, `Workbook_BeforeSave`,
  `Workbook_BeforeClose`…) don't run yet.
- Hostile code is stopped, never the app: nesting deeper than 128 levels is "Expression too
  complex", runaway recursion "Out of stack space", strings past 64M characters and arrays past
  16M elements "Out of memory", and a run of more than 500 million steps is stopped.

[MS-OVBA]: https://learn.microsoft.com/openspecs/office_file_formats/ms-ovba/
[MS-VBAL]: https://learn.microsoft.com/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-30 | major | Created: security (Enable Content, `vba.trust`), commands, limits |
| 2026-10-30 | minor | Events: `Worksheet_Change`, `Workbook_SheetChange`; sheet code names kept |
| 2026-10-30 | minor | Text files: `Open`, `Print #`, `Write #`, `Input #`, `Line Input #`, `Close`, `EOF`, `LOF`, `FreeFile` |
