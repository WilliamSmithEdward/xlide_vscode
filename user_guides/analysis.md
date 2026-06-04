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
- **Analyze Workbook** from the workbook tree or XLIDE Activity Bar/sidebar.
- `xlide_analyzeWorkbook` from an AI-agent workflow.

Workbook analysis opens a dedicated results panel. It groups findings by module,
shows counts, supports severity filters, can show suppressed diagnostics, and
links each finding back to the workbook module and source line.

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
  this workbook.

## Filter And Track Rules

The analysis results panel is the safest place to change workbook analysis
visibility. It preserves setting provenance and writes workbook-specific choices
to the workbook sidecar only when you choose a workbook-scoped action.

Use these actions when available:

- **Untrack In Workbook** - writes the selected diagnostic code to
  `<workbook>.xlide_settings.json`.
- **Untrack Globally** - writes the selected diagnostic code to the global
  machine/profile setting `xlide.analysis.untrackedRules`.
- Severity controls - use guarded `xlide.analysis.ruleSeverityOverrides` values
  where the rule allows it.

Global defaults live in VS Code machine/profile settings. Workbook overrides
live beside the workbook in:

```text
<workbook>.xlide_settings.json
```

Example workbook analysis settings:

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
change how Excel or VBE compiles and runs the project.

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

Use **Untrack In Workbook** when a rule is not useful for one workbook.

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
- Fix malformed `<workbook>.xlide_settings.json` files before expecting workbook
  overrides to apply.

If you are unsure whether to suppress a finding, leave it visible and use
tracking/settings first. Source suppressions are best when the reason is stable
and worth documenting in the VBA source.
