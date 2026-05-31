# Excel/VBE Oracle Harness

This folder contains an optional Windows-only harness that asks a real Excel/VBE
instance whether small VBA snippets are accepted or rejected.

It is a developer audit tool only. XLIDE production code must stay deterministic
and must not depend on Excel, COM, Office, or VBE automation at runtime.

## Run

```powershell
npm run test:oracle:vbe
```

Useful filters:

```powershell
python syntax_corpus/oracle/run_excel_vbe_oracle.py --case missing_trailing_required_argument
python syntax_corpus/oracle/run_excel_vbe_oracle.py --timeout 30 --json
python syntax_corpus/oracle/run_excel_vbe_oracle.py --strict
```

## Requirements

- Windows
- Microsoft Excel installed
- Trust access to the VBA project object model enabled
- No unrelated modal Excel/VBE prompts blocking automation

The runner starts an unsaved disposable Excel workbook per case. Keeping the
workbook in memory avoids local macro-security policy blocking generated test
macros. If a case hangs, the Python coordinator times out the worker and attempts
to kill only the Excel process that the worker recorded.

By default the runner is observational and exits successfully even when a
fixture expectation differs from the observed result. Use `--strict` when the
local Excel/VBE automation path is stable enough to enforce expectations.

## Fixture Policy

Fixtures record empirical VBE behavior. They should be used to verify edge cases
and update deterministic analyzer rules, not to add guessed behavior.

The default fixture mode is `compile`. Compile fixtures insert the fixture source,
focus the disposable module in VBE, then invoke VBE's Debug > Compile VBAProject
UI path (`Alt+D`, then `L`). A `Compile error:` dialog is treated as compile
rejection. If no compile dialog appears before the popup watch period ends, the
case is recorded as compile-accepted. `run` mode is reserved for future
runtime-behavior probes and should be used sparingly because modal runtime errors
are harder to automate safely.

During compile fixtures, the worker starts a Win32 dialog watcher before invoking
VBE Compile. A dialog owned by the disposable Excel process with title
`Microsoft Visual Basic for Applications`, class `#32770`, and child text
containing `Compile error:` is recorded as `outcome: "rejected"`; the watcher
captures the dialog text and dismisses it with the OK command. After any VBE
dialog is dismissed, the watcher posts VBE's Reset command (`CommandBar` id 228)
and the Debug > Reset accelerator (`Alt+D`, then `R`) to the disposable VBE main
window so Excel does not remain paused on the highlighted source line. Timeouts
remain a fallback for hangs or unrelated modal prompts, not the primary rejection
signal.

Use `expected: "observe"` when the case exists to collect behavior but the repo
does not yet assert a specific result.

Every oracle case must also declare `provenance`:

- `vbe-oracle-verified` for accepted/rejected expectations that are intended to
  assert Excel/VBE compile behavior.
- `observed-not-asserted` for `expected: "observe"` cases.

Oracle provenance is case-level. The file-level entry in
`syntax_corpus/corpus_provenance.json` only records that this fixture file is
owned by the oracle workflow.
