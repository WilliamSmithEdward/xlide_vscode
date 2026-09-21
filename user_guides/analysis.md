# XLIDE Analysis And Ignores

XLIDE analysis is designed to be useful without becoming noisy. Red diagnostics
are reserved for VBA compile errors or deterministic runtime failures that XLIDE
can prove. Warnings and information diagnostics are guidance, tracking, or softer
signals.

When XLIDE cannot prove a problem, it prefers no hard diagnostic over a guessed
error.

## Run Analysis

Use one of these entry points:

- Live diagnostics in open `xlide-vba` editors.
- **Analyze Current Module** from the editor context menu or command palette.
- **Analyze Project** from the file tree or XLIDE Activity Bar/sidebar.
- `xlide_analyzeProject` from an AI-agent workflow.

Live diagnostics cover the modules the XLIDE tree lists: the ones inside a
workbook, document, presentation or database, and the files a VB6 project
names. A `.bas`, `.cls` or `.frm` on disk that no project claims is left
alone. That is usually an exported copy of a module the tree already analyzes,
and analyzing it too would show every finding twice in the Problems panel.
Completion, hover and navigation still work in those files. To analyze them as
standalone modules, turn off `xlide.analysis.ignoreFilesOutsideTree`, shown as
**Ignore Files Outside The XLIDE Tree** in Global Settings.

File analysis opens a dedicated results panel. It groups findings by module,
shows counts, supports severity filters, can show suppressed diagnostics, and
links each finding back to the module and source line.

## Understand Results

Each finding has a stable diagnostic code. Use the code when configuring
tracking, severity overrides, or source suppressions.

Common result meanings:

- **Error** - XLIDE can prove a compile-equivalent error or deterministic runtime
  failure.
- **Warning** - useful guidance, a non-compile safety issue, or a softer project
  signal.
- **Information** - low-severity guidance or tracking context.
- **Suppressed** - the diagnostic exists, but a source suppression comment hides
  it from the active result set.
- **Untracked** - the rule is intentionally hidden from tracking globally or for
  this file.

## Another Application's Library

A project that references another Office application can name its types. In a
Word document that references Excel, `Dim xl As Excel.Application` compiles,
and XLIDE checks `xl.Workbooks.Add` against Excel's object model, completes
Excel's members after `xl.`, and knows Excel's spellings. Excel, Word,
PowerPoint and Access are covered in every combination.

Without the reference the VBE refuses the declaration - "User-defined type not
defined" - and the project stops compiling. XLIDE reports that as
`missing-library-reference`, once per library per module, with a quick fix that
writes the reference into the project. **Add Project Reference** on a project
in the XLIDE tree does the same without waiting for an error.

Late binding needs no reference, and is never reported:

```vba
Dim xl As Object
Set xl = CreateObject("Excel.Application")
```

It names nothing from the library, which is the whole point of it.

## Dead Code

Four rules point at code the module does not need. They report as
Information, and the editor fades the range instead of underlining it.

| Code | Reports | Never reports |
|---|---|---|
| `unused-variable` | A local, or a module-level `Private`/`Dim` variable or constant, that nothing in its scope names. | `Public` and `WithEvents` declarations, and variables with attributes. |
| `variable-never-read` | A variable whose every mention assigns to it. | A `For` counter, an array written by element, a `ReDim Preserve`, a `Mid` target, or a variable passed to a procedure. Those count as reads. |
| `unused-procedure` | A `Private` Sub, Function or Property that nothing in its module calls and no string in the project names. | `Public` procedures, event handlers, interface members, `Auto_Open` and friends, and procedures with attributes. |
| `unreachable-code` | Statements after `Exit Sub`, `Exit Do`, `GoTo`, `Resume`, `End` or `Return` in the same block. | Anything after a label, a line number, a `Case` arm or a `#If`, which give control somewhere to land. |

Public procedures are left alone on purpose. A button, a shape, the ribbon, a
hotkey, `OnTime` or `Application.Run` in another file can all call one, and
XLIDE cannot see any of them. A Private procedure can only be reached from its
own module, or by name in a string, so every string literal in the project is
searched before it is reported.

`unused-variable` offers a quick fix that removes the declaration, and
`unreachable-code` one that removes the lines. To silence a rule, set it to
`off` in `xlide.analysis.ruleSeverityOverrides`, or untrack it from the
analysis results panel.

## Doc Comments

When a procedure's `'''` doc comment is written in XML, six rules check it
against the declaration. Every parameter needs a `<param>`, a Function needs
a `<returns>`, and no tag may describe something the declaration does not
have, repeat, or be left open. They report as warnings, and most come with a
quick fix. The codes and what each accepts are in
[vba-doc-comments.md](vba-doc-comments.md#what-xlide-checks).

## Filter And Track Rules

The analysis results panel is the safest place to change analysis
visibility. It preserves setting provenance and writes file-specific choices
to the file's sidecar only when you choose a file-scoped action.

Use these actions when available:

- **Untrack In Project** - writes the selected diagnostic code to
  `<file>.xlide_settings.json`.
- **Untrack Globally** - writes the selected diagnostic code to the global
  machine/profile setting `xlide.analysis.untrackedRules`.
- Severity controls - use guarded `xlide.analysis.ruleSeverityOverrides` values
  where the rule allows it.

Global defaults live in VS Code machine/profile settings. Per-file overrides
live beside the file in:

```text
<file>.xlide_settings.json
```

Example per-file analysis settings:

```json
{
  "analysis": {
    "visibleSeverities": ["error", "warning", "information"],
    "untrackedRules": ["option-explicit-missing"],
    "ruleSeverityOverrides": {
      "unknown-call": "warning"
    }
  }
}
```

XLIDE validates settings strictly. If the sidecar has invalid JSON, unknown
keys, invalid severities, unknown diagnostic codes, or disallowed severity
overrides, XLIDE reports a settings error instead of silently falling back.

## Suppress In Source

Use source suppressions when the code itself should carry the reason. Suppression
comments are ordinary VBA comments. They affect XLIDE analysis only; they do not
change how the VBE compiles or the application runs the project.

Prefer suppressing one diagnostic code over suppressing `all`.

Suppress the next line:

```vba
' @xlide-analysis-disable-next-line argument-count -- legacy host callback
RunReport reportId, , True
```

Suppress the current physical line:

```vba
total = "100" ' @xlide-analysis-disable-line assignment-type-mismatch
```

A statement continued with ` _` counts as one line for both directives, so
put `disable-next-line` above its first line.

Suppress the next member:

```vba
' @xlide-analysis-disable-next-member non-callable-call -- external macro entry
Public Sub Workbook_Open()
    Start
End Sub
```

Suppress a whole module:

```vba
' @xlide-analysis-disable-file all -- generated module
Option Explicit
```

Suppress a block:

```vba
' @xlide-analysis-disable-block argument-type-mismatch -- imported legacy block
total = InvoiceTotal("100", 0.08)
caption = BuildCaption(1)
' @xlide-analysis-enable-block argument-type-mismatch
```

Supported directive forms:

- `@xlide-analysis-disable-file`
- `@xlide-analysis-disable-next-member`
- `@xlide-analysis-disable-line`
- `@xlide-analysis-disable-next-line`
- `@xlide-analysis-disable-block`
- `@xlide-analysis-enable-block`

Each directive accepts either `all`, a comma-separated list of diagnostic codes,
or no code list. No code list means `all`.

Documentation comments beginning with `'''` are never treated as suppression
directives.

## Choose The Right Ignore

Use **filters** when you only want to inspect part of a result set.

Use **Untrack In Project** when a rule is not useful for one file.

Use **Untrack Globally** when a rule is not useful for your machine/profile.

Use **source suppressions** when the exception belongs to the code and should be
visible to teammates and future analysis runs.

Use **severity overrides** only when the rule allows the requested downgrade or
disable behavior. XLIDE does not allow settings to turn uncertain behavior into a
red diagnostic.

## Troubleshooting

If a diagnostic does not disappear:

- Check that the suppression comment uses a single apostrophe comment, not `Rem`.
- Check the diagnostic code spelling in the results panel.
- Check whether the directive applies to the right scope.
- Check the analysis results panel for directive diagnostics.
- Fix malformed `<file>.xlide_settings.json` files before expecting per-file
  overrides to apply.

If a `.bas`, `.cls` or `.frm` file shows no diagnostics at all, check whether
it belongs to a project in the XLIDE tree. A file no project claims is not
analyzed while `xlide.analysis.ignoreFilesOutsideTree` is on, which is the
default.

If you are unsure whether to suppress a finding, leave it visible and use
tracking/settings first. Source suppressions are best when the reason is stable
and worth documenting in the VBA source.
