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
python syntax_corpus/oracle/run_excel_vbe_oracle.py --timeout-retries 2
python syntax_corpus/oracle/run_excel_vbe_oracle.py --case string_scalar_member_access_compile --dialog-hold-seconds 20
python syntax_corpus/oracle/run_excel_vbe_oracle.py --strict
```

Promotion workflow for an observe-only case:

```powershell
python syntax_corpus/oracle/run_excel_vbe_oracle.py --case local_string_with_argument_statement --promote-observed
```

Promotion is deliberately narrow. It requires at least one `--case`, refuses
fixtures whose `expected` value is already asserted, and writes only
`accepted`/`rejected` outcomes back to `vbe_oracle_cases.json` with
`provenance: "vbe-oracle-verified"`, `evidencePhase`, and
`diagnosticMeaning`. Timeouts and worker errors remain unpromotable.

## Requirements

- Windows
- Microsoft Excel installed
- Trust access to the VBA project object model enabled
- No unrelated modal Excel/VBE prompts blocking automation

The runner starts an unsaved disposable Excel workbook per case. Keeping the
workbook in memory avoids local macro-security policy blocking generated test
macros. If a case hangs, the Python coordinator times out the worker and attempts
to kill only the Excel process that the worker recorded. Timeout is not VBA
evidence: it is treated as oracle infrastructure health. The coordinator retries
the same case up to `--timeout-retries`; if every attempt times out, it aborts
the remaining run as `outcome: "oracle_failure"` and exits non-zero even without
`--strict`. Investigate the harness/Excel modal state before running additional
oracle cases.

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
case is recorded as compile-accepted.

The worker keeps Excel visible and explicitly restores/focuses the disposable
Excel/VBE windows before UI-sensitive actions. This is important when the
developer is alt-tabbed into another application while the oracle runs. Dialog
detection is still owned by Win32 window enumeration for the recorded Excel
process; timeouts remain only a fallback and never satisfy an expected or
observe-only fixture result.

For harness debugging, `--dialog-hold-seconds N` keeps the first detected VBE
dialog visible for `N` seconds before dismissal so a developer can inspect or
screenshot it. This option is not evidence-changing; it only delays cleanup.
Use a `--timeout` comfortably larger than the hold window so the worker has time
to detect the dialog and clean up afterward.

`run` mode is for focused runtime-behavior probes only. Runtime fixtures must
name an `entryPoint`. The worker starts the same Win32 dialog watcher before
running the macro; a `Microsoft Visual Basic` dialog with `Run-time error` text
is recorded as `outcome: "rejected"` and dismissed with the dialog's `End`
button. Use runtime fixtures sparingly, only when a deterministic runtime error
changes diagnostic severity or corpus truth.

During compile fixtures, the worker starts a Win32 dialog watcher before invoking
VBE Compile. A dialog owned by the disposable Excel process with title
`Microsoft Visual Basic for Applications`, class `#32770`, and child text
containing `Compile error:` is recorded as `outcome: "rejected"`; the watcher
captures the dialog text, dismisses it with direct Win32 button/window commands,
verifies the dialog is no longer visible, and only then writes the oracle result
file. It does not reopen the Debug menu after dismissal; the disposable Excel
instance is closed immediately afterward. Timeouts remain a fallback for hangs
or unrelated modal prompts, not the primary rejection signal. A timeout never
means accepted or rejected; after the retry budget is exhausted, it means the
oracle itself needs attention.

Use `expected: "observe"` when the case exists to collect behavior but the repo
does not yet assert a specific result.

Every oracle case must also declare `provenance`:

- `vbe-oracle-verified` for accepted/rejected expectations that are intended to
  assert Excel/VBE behavior.
- `observed-not-asserted` for `expected: "observe"` cases.

Every oracle case must declare the phase it proves and the meaning of that
observation:

- `evidencePhase: "compile"` for Debug > Compile evidence.
- `evidencePhase: "runtime"` for `excel.Run` evidence.
- `diagnosticMeaning: "compile-error"` when VBE Compile rejects the source.
- `diagnosticMeaning: "runtime-error"` when the source compiles but running the
  entry point raises a deterministic VBE runtime dialog.
- `diagnosticMeaning: "compile-valid"` or `"runtime-valid"` for controls that
  prove accepted behavior.
- `diagnosticMeaning: "observation"` for observe-only fixtures.

Oracle provenance is case-level. The file-level entry in
`syntax_corpus/corpus_provenance.json` only records that this fixture file is
owned by the oracle workflow.
