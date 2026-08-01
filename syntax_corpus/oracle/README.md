# Excel/VBE Oracle Harness

This folder contains an optional Windows-only harness that asks a real Excel/VBE
instance whether small VBA snippets are accepted or rejected.

It is a developer audit tool only. XLIDE production code must stay deterministic
and must not depend on Excel, COM, Office, or VBE automation at runtime.

## Run

```powershell
npm run test:oracle:vbe
```

Run oracle commands sequentially. Do not run multiple `test:oracle:vbe` or
`run_excel_vbe_oracle.mjs` invocations in parallel, even for different cases.
The harness drives Excel/VBE through a single-user COM/UI automation surface;
parallel runs can contend for VBE focus, command bars, modal dialogs, and Excel
process cleanup, producing timeouts or misleading contradictory results. When
several cases need retesting, run one command, wait for it to finish and clean
up, then run the next.

Useful filters:

```powershell
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --case missing_trailing_required_argument
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --timeout 30 --json
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --timeout-retries 2
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --case string_scalar_member_access_compile --dialog-hold-seconds 20
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --strict
```

Promotion workflow for an observe-only case:

```powershell
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --case local_string_with_argument_statement --promote-observed
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
macros. If a case hangs, the coordinator times out the worker and attempts
to kill only the Excel process that the worker recorded. Timeout is not VBA
evidence: it is treated as oracle infrastructure health. The coordinator retries
the same case up to `--timeout-retries`; if every attempt times out, it aborts
the remaining run as `outcome: "oracle_failure"` and exits non-zero even without
`--strict`. Investigate the harness/Excel modal state before running additional
oracle cases. If an oracle failure follows accidental parallel oracle runs,
treat the parallelism as the first suspected infrastructure cause and rerun the
cases sequentially after checking for lingering Excel/VBE processes or modal
dialogs.

Only `accepted` and `rejected` are evidence outcomes. Any other worker outcome
(`timeout`, `worker_error`, setup failure, malformed output, and similar) is an
oracle infrastructure failure, aborts the remaining run, and exits non-zero.
If the compile command cannot be invoked, or any COM operation fails before a
VBE dialog is captured, the case is not treated as a VBA rejection.

For extension development, `ORACLE-FAIL` is a stop-the-line signal for the AI
agent. Do not continue adding or promoting oracle cases after it appears. First
perform a deep application inspection: check for lingering Excel/VBE processes,
modal dialogs, worker stage output, command-bar invocation, dialog detection and
dismissal, timeout/retry behavior, recent harness edits, and whether the failure
is reproducible with a single focused case. Resume oracle-backed development
only after the harness failure has a concrete explanation or fix.

By default the runner is observational and exits successfully even when a
fixture expectation differs from the observed result. The local Excel/VBE
automation path is operational; use `--strict` to enforce expectations.

## Timing, reliability, and lessons learned

Hard-won notes from getting the oracle reliable on a fresh machine. The failure
mode to fear most is a **silent false "accepted"**: the worker compiles a snippet
that VBE actually rejects, fails to capture the error dialog, and records the case
as accepted. That poisons the no-false-positive discipline (a rule shipped against
a bogus "accepted" control). Guard against it with the controls below.

1. **Excel must run visible.** The worker sets `Application.Visible = $true`. When
   the host is hidden, the VBE `Compile error:` / `Run-time error` modal is not
   surfaced as a detectable visible top-level window, so the Win32 dialog watcher
   never sees it and the rejection is misrecorded as `accepted`.

2. **Every case pays a cold-VBE first-compile cost.** Each case spawns its own
   Excel process, so the error dialog can take well over ten seconds to appear.
   The dialog-watch window must outlast that latency or rejects are missed — the
   original fixed 8-second cap was too short and produced systematic false
   "accepted" results.

3. **`--timeout` is the single knob.** In normal mode the watch window is
   `min(22, timeout - 22)` seconds; raising `--timeout` adds Excel-startup
   headroom. For a window longer than the 22s cap (a very slow VBE), use
   `--dialog-hold-seconds`, whose branch uses `timeout - hold - 5` uncapped.

4. **Accept cases bound the timeout.** A genuinely-accepted case shows no dialog,
   so the worker waits the *entire* watch window before concluding `accepted`,
   then quits Excel. `excel_start + watch + cleanup` must fit inside `--timeout`
   or the coordinator kills the worker as a false `timeout`/`oracle_failure`.
   Reject cases return as soon as the dialog is caught, so they are fast; it is
   the accept controls that set the timeout floor.

5. **The first run on a cold machine may time out entirely.** The first-ever COM
   Excel launch (Office first-run/activation) can exceed the timeout before
   compilation even starts. Treat a *single* initial timeout as cold-start
   warmup, re-run, and expect subsequent runs to be faster as the OS/COM caches
   warm. A *systematic* timeout or `oracle_failure` after warmup is a real harness
   problem — follow the `ORACLE-FAIL` deep-inspection steps below.

6. **Sanity-check the harness before trusting a batch.** Run a couple of
   known-reject fixtures and one known-accept fixture under `--strict`; if a
   known-reject reports `accepted`, the dialog capture is broken (see #1/#2) — do
   not promote anything until it is fixed:

   ```powershell
   py syntax_corpus/oracle/run_excel_vbe_oracle.mjs `
     --case missing_trailing_required_argument `
     --case bare_variable_statement `
     --case number_to_string_assignment --strict --json
   ```

7. **Tuning for efficiency.** The default `--timeout 45` suits a warm machine.
   Reject-heavy batches finish quickly (early return). To speed accept-heavy
   batches once warm, try `--timeout 40` and watch for accept-case timeouts;
   raise it for slow/cold machines. A future efficiency win is to warm the VBE
   compiler once per worker so a shorter watch window suffices.

## Automation Policy

The oracle must not use `SendKeys`, keyboard focus scripting, or
timing-dependent keystroke playback as evidence. Driving Excel/VBE should use
deterministic low-level hooks: COM object model calls for workbook/VBE actions,
VBE command-bar controls for compile invocation, Win32 window enumeration and
messages for dialog capture/dismissal, or UI Automation hooks where COM/Win32
cannot expose the required state. If a case can only be exercised through
foreground focus or keystroke replay, keep it unsupported or pending until a
low-level hook exists.

The harness must also account for blockers that could prevent completion. Known
prompts should be prevented or handled through low-level hooks. Unknown modals,
hangs, setup failures, and automation-busy states are oracle infrastructure
failures after timeout/retry cleanup, not accepted or rejected VBA evidence.

## Fixture Policy

Fixtures record empirical VBE behavior. They should be used to verify edge cases
and update deterministic analyzer rules, not to add guessed behavior.

The default fixture mode is `compile`. Compile fixtures insert the fixture
source, show the disposable module in VBE, then invoke the exact
`Compile <project-name>` VBE command bar control through COM. The worker does
not use keyboard accelerators for compile invocation. A `Compile error:` dialog
is treated as compile rejection. If no compile dialog appears before the popup
watch period ends, the case is recorded as compile-accepted.

By default a fixture supplies one standard-module `source` string. For
workbook-object behavior, a fixture may instead supply a `modules` array. Each
entry has `name`, `type` (`standard` or `class`), and `source`; mark the standard
module containing the `entryPoint` with `entry: true` when it is not the first
standard module. This lets the oracle verify class-module member behavior
without changing production code.

Use `mode: "compile_then_run"` for discovery cases where the failure phase is
not known yet. The coordinator invokes compile first. If compile rejects, the
case records compile rejection and does not run. If compile accepts, the
coordinator runs the declared `entryPoint` and records runtime acceptance or
runtime rejection. This is the preferred mode for new runtime-behavior
investigation; convert to a narrower `compile` or `run` fixture only when the
phase being asserted is already known.

`compile_then_run` is a discovery workflow, not the default long-term assertion
format. After a discovery case produces evidence, narrow the promoted fixture to
the phase actually being asserted:

- Compile rejects: convert to `mode: "compile"`, `expected: "rejected"`,
  `evidencePhase: "compile"`, `diagnosticMeaning: "compile-error"`.
- Compile accepts and runtime behavior matters: convert to `mode: "run"` with
  `expected: "accepted"` or `"rejected"` from the runtime result,
  `evidencePhase: "runtime"`, and the matching runtime diagnostic meaning.
- Compile accepts and runtime behavior does not matter: use `mode: "compile"`
  with `expected: "accepted"` and `diagnosticMeaning: "compile-valid"`.
- Keep `compile_then_run` only for intentional harness controls or unresolved
  discovery probes that need to keep proving the full two-phase path.

The worker keeps Excel visible but does not require foreground focus for normal
compile or run evidence. Dialog detection is owned by Win32 window enumeration
for the recorded Excel process; timeouts remain only a fallback and never
satisfy an expected or observe-only fixture result.

For harness debugging, `--dialog-hold-seconds N` keeps the first detected VBE
dialog visible for `N` seconds before dismissal so a developer can inspect or
screenshot it. This option is not evidence-changing; it only delays cleanup.
Use a `--timeout` comfortably larger than the hold window so the worker has time
to detect the dialog and clean up afterward.

`run` mode is for focused runtime-behavior probes only. Runtime fixtures must
name an `entryPoint`. The worker starts the same Win32 dialog watcher before
running the macro; a `Microsoft Visual Basic` dialog with `Run-time error` text
is recorded as `outcome: "rejected"` and dismissed with the dialog's `End`
button. The dismiss fallback also sends standard OK/Cancel/Abort dialog commands
because VBE runtime dialogs vary in their button command IDs. Use runtime
fixtures sparingly, only when a deterministic runtime error changes diagnostic
severity or corpus truth.

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

- `evidencePhase: "compile"` for VBE Compile command evidence.
- `evidencePhase: "runtime"` for `excel.Run` evidence.
- `mode: "compile_then_run"` may produce either compile or runtime evidence
  depending on where Excel/VBE rejects the snippet.
- `diagnosticMeaning: "compile-error"` when VBE Compile rejects the source.
- `diagnosticMeaning: "runtime-error"` when the source compiles but running the
  entry point raises a deterministic VBE runtime dialog.
- `diagnosticMeaning: "compile-valid"` or `"runtime-valid"` for controls that
  prove accepted behavior.
- `diagnosticMeaning: "observation"` for observe-only fixtures.

Oracle provenance is case-level. The file-level entry in
`syntax_corpus/corpus_provenance.json` only records that this fixture file is
owned by the oracle workflow.
