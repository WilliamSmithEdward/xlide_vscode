# VBE Oracle Harness

This folder contains an optional Windows-only harness that asks the VBE of a
real Office application, Excel, Word, PowerPoint or Access, whether small VBA
snippets are accepted or rejected. Its sibling for VB6, `twinbasic/`, asks
twinBASIC's compiler the same question and has its own README.

It is a developer audit tool only. XLIDE production code must stay deterministic
and must not depend on Excel, COM, Office, or VBE automation at runtime.

## Hosts

A fixture names the application it runs in with `host`: `excel`, `word`,
`powerpoint` or `access`. A fixture without one runs in Excel, which is where
most of the corpus was measured. The coordinator keeps its own copy of the
host table in `src/officeHostApps.ts` (ProgID and process image), and
`tests/oracle/vbeOracleWorker.test.ts` keeps the two in step. A change to the oracle
for one host is made for all four.

Each case gets a fresh file in its host: a workbook, a document, a
presentation, or an Access database written to the temp folder and deleted
afterwards. Where the hosts differ (measured on Office 16.0 build 20430):

- Word and PowerPoint can hand an automation request to an instance that is
  already running, and PowerPoint always does. The worker refuses to start
  either while one is open, with a `worker_error` that says so, rather than
  compile into someone's instance and then quit it.
- A new module in Word, PowerPoint or Access holds only the case's source:
  the worker deletes what the VBE put there, such as Access's own
  `Option Compare Database`. Excel keeps what it always kept.
- Access is never kept between cases: closing its database asks whether to
  save the case's modules, so each case quits it without saving.
- Run takes `'file'!Module.Proc` in Excel and PowerPoint, `Project.Module.Proc`
  in Word, and only the bare procedure name in Access.

`run_excel_vbe_oracle.mjs`, the coordinator's name when it drove only Excel,
still works, and so does `--cases-per-excel` for `--cases-per-instance`.

Fixtures containing procedure `Attribute ... .VB_...` lines are imported from
an exported component header, so default-member and enumeration attributes are
real VBE metadata rather than invalid editor statements. The worker removes
only its newly created component and deletes its temporary import file.

## Bitness

A fixture measured in 32-bit Office says `"bitness": 32`; one without a
`bitness`, or with `64`, was measured in 64-bit Office, where most of the
corpus was. The two compile different code: 32-bit Office has no LongLong
and its LongPtr is a Long, so a fixture's verdict holds for its bitness only.

- The coordinator measures each host's bitness from its registered exe (the
  PE header's machine field), in the probe that already finds the exe. A
  case for the other bitness is not run: it is reported `SKIP` with the
  reason (`stage: other_bitness`), like a known host crash. `--strict` counts
  it as skipped, not failed, so a run where every case was skipped exits 0;
  the summary says how many were for the other bitness and whether none ran.
  A run whose cases are all for the other bitness starts no Office and leaves
  the dialog sounds alone.
- A host whose bitness cannot be read runs nothing, so no compile is recorded
  against the wrong bitness.
- Each result records `hostBitness`, and the checkpoint hash covers it with
  the host build, which is the same number in both bitnesses.
- To measure 32-bit fixtures, run the coordinator on a machine with 32-bit
  Office.
- The corpus tests analyze a 32-bit fixture with `Win64` off, through
  `oracleConditionalCompilation()` in `tests/helpers/oracleModuleInput.ts`.
  A 32-bit fixture references no VBA project yet. The twinBASIC oracle leaves
  32-bit fixtures out of its parity report.

## Run

```powershell
npm run test:oracle:vbe
```

While it runs, the runner silences four Windows sounds for your user: Default
Beep, Asterisk, Exclamation and Critical Stop. Most rejected cases show a VBE
error box, which plays one of these sounds, and on some machines each sound
leaves thousands of kernel trace registrations that Windows frees only slowly,
so a long run could fill gigabytes of kernel memory. Windows plays the Default
Beep in place of a silenced dialog sound, so all four are needed; other sounds
are left alone. The runner first saves their values and registry types in
`~/.xlide-vbe-oracle-sounds.json` in your account profile, then puts them back
when it finishes or is interrupted, except any you set to another sound or
removed meanwhile. If it is killed before putting them back, the next run
does; a file that cannot be read is moved aside with a message instead. While
another run is going, a second run leaves the sounds to it. Recovery validates
the complete backup before any registry write: owner identity, unique allowed
events, supported string registry kinds and nonempty string sound paths.
Malformed backups are quarantined intact rather than restored or discarded.

Run oracle commands sequentially. Do not run multiple `test:oracle:vbe` or
`run_vbe_oracle.mjs` invocations in parallel, even for different cases.
The harness drives the VBE through a single-user COM/UI automation surface;
parallel runs can contend for VBE focus, command bars, modal dialogs, and host
process cleanup, producing timeouts or misleading contradictory results. When
several cases need retesting, run one command, wait for it to finish and clean
up, then run the next.

The runner enforces this per Windows user (issue #1438). Before it starts
any Office automation or touches the sounds, it takes a lock:
`oracleRunLock.mjs` listens on a named pipe named after your account. A
second run, from any directory, repository or terminal, exits with code 3
and names the run that holds it: its process, machine, start time,
directory and command. Wait for that run to finish, or stop it.

Windows frees the pipe the moment its owner exits, however it exits, so no
lock file can go stale and there is nothing to delete after a crash. If an
owner answers with something unreadable, or not in time, the run refuses
and starts nothing rather than guess. If a run is killed outright, for
example with Task Manager or `taskkill /F`, the next run can start at once,
and it first ends the Office instance the killed run left (below).

The runner ends what it started, and only that (issue #1439). The worker
records each Office instance it starts by image, process id and start time,
and the runner keeps that record until the instance is gone. Before it ends
one, the process at that id must still have that image and start time, so a
reused id or someone else's Office session is never touched.

- **Normal end, an exception, Ctrl-C, Ctrl-Break and closing the console**
  take one cleanup. It ends the worker and every recorded instance, removes
  the current case's temporary folder and puts the sounds back. Running it
  twice does nothing more. The lock is freed when the process exits.
- **Between cases**, any recorded instance but the one the worker still
  drives must be gone before the next case. One that cannot be ended stops
  the run with a `worker_error` at stage `host_cleanup` that names it.
- **Outright termination** runs no cleanup, so the records are also kept in
  `~/.xlide-vbe-oracle-hosts.json`. The next run reads it, ends those
  instances after the same identity check, and says which it ended. An
  instance it cannot end is named, and the run starts nothing.

Two limits remain. An instance that is still starting, between Office
handing it to the worker and the worker recording it, is not yet known.
If the run is killed in that moment, the instance stays open with its VBE
showing; end it in Task Manager. And Windows gives a closing console about
five seconds. The cleanup is synchronous and fits that for a run's one
instance, but an instance slower to die than that stays recorded, and the
next run ends it.

Every option that takes a value needs one, and the runner checks them all
before it starts anything (issue #1440). A value may not begin with `--`, so
`--checkpoint --strict` is refused rather than writing a file named
`--strict`; for a path that does begin so, write `./--name`. A fixture's
`host` is left out for Excel or is exactly `excel`, `word`, `powerpoint` or
`access`; anything else, `Excel` and `null` included, is refused. A
`--checkpoint` whose directory is missing or read-only is refused before
Office starts.

A case whose code names a type library early bound lists it in
`references`, each entry `{ "guid": "{...}", "major": 9, "minor": 6 }`.
The worker adds them by GUID and version before any module, as Tools >
References would, so the library must be registered on the machine. A
malformed entry is refused before Office starts.

Useful filters:

```powershell
node syntax_corpus/oracle/run_vbe_oracle.mjs --case missing_trailing_required_argument
node syntax_corpus/oracle/run_vbe_oracle.mjs --timeout 30 --json
node syntax_corpus/oracle/run_vbe_oracle.mjs --timeout-retries 2
node syntax_corpus/oracle/run_vbe_oracle.mjs --case string_scalar_member_access_compile --dialog-hold-seconds 20
node syntax_corpus/oracle/run_vbe_oracle.mjs --strict
node syntax_corpus/oracle/run_vbe_oracle.mjs --checkpoint oracle-results.jsonl
```

With `--checkpoint`, each case's result is appended to the file as soon as it
finishes. Running the same command again skips every case that already has an
`accepted` or `rejected` result there, so a run stopped by an `ORACLE-FAIL`
resumes where it stopped once the cause is fixed. Such results are marked
`resumed: true`. A case runs again if anything in its fixture has changed,
including its `expected` value or description, or if the worker or coordinator
has, or if the build of its host has.

That makes a checkpoint file safe to keep as a verdict cache across runs: an
Office update reruns every case. A host's build is the file version of the
executable its ProgID is registered to (`EXCEL.EXE`, `WINWORD.EXE`,
`POWERPNT.EXE`, `MSACCESS.EXE`), read once per run without starting it, for
each host the cases use; `--checkpoint` refuses to run if one cannot be read.
Each result records its `host` and the build of the running instance that
produced it as `hostBuild`, and a result whose build is missing or differs is
an infrastructure failure, never cached evidence. An Office update during a
run stops it at the next case that starts the new build; run it again, and
the old build's entries are not reused.

A branch that adds or edits fixtures can measure just those:

```powershell
node syntax_corpus/oracle/run_vbe_oracle.mjs --changed --checkpoint oracle-results.jsonl
node syntax_corpus/oracle/run_vbe_oracle.mjs --changed-from my-base-branch
```

`--changed` compares the cases file on disk, uncommitted edits included, with
the same file at the commit where the branch left `upstream/main`, or
`origin/main` when there is no `upstream`. `--changed-from <ref>` names
another base. A fixture runs when its id is new or any of its fields differs
from the base fixture of that id; a fixture the branch removed does not run,
and nor does one main changed after the branch left it. The run first prints
how many cases it selected for each host. With none, it exits 0 before it
takes the run lock or starts Office.

When the base cannot be worked out, `--changed` exits with code 2 and says
why, rather than running every fixture: git is missing, the file is not
tracked or not in a work tree, the clone is shallow, the base ref is missing,
it shares no history with `HEAD`, or the file does not exist or cannot be
read there. It cannot be combined with `--case`. It works with `--checkpoint`,
which still skips what an earlier run measured.

Promotion workflow for an observe-only case:

```powershell
node syntax_corpus/oracle/run_vbe_oracle.mjs --case local_string_with_argument_statement --promote-observed
```

Promotion is deliberately narrow. It requires at least one `--case`, refuses
fixtures whose `expected` value is already asserted, and writes only
`accepted`/`rejected` outcomes back to `vbe_oracle_cases.json` with
`provenance: "vbe-oracle-verified"`, `evidencePhase`, and
`diagnosticMeaning`. Timeouts and worker errors remain unpromotable.
It rewrites the whole file as `JSON.stringify(document, null, 2)` with DEL
and every non-ASCII character written as `\uXXXX`, so lay out hand-edited
cases the same way; `tests/oracle/vbeOracleKnownCrash.test.ts` checks it.

## Requirements

- Windows
- The Office applications the cases name installed (Excel for the bulk of
  the corpus)
- Trust access to the VBA project object model enabled in each of them
  (Access has no such setting)
- No unrelated modal VBE prompts blocking automation, and no Word or
  PowerPoint of your own open while cases for that host run

The runner starts an unsaved disposable file per case (Access's database is
on disk in the temp folder). Keeping the file in memory avoids local
macro-security policy blocking generated test macros. One worker serves the
cases in turn, and one host instance serves compile cases of that host,
which run no code: the worker closes each case's file, and starts a new
instance after every run case, after every Access case, after
`--cases-per-instance` cases (default 40), after any case that is not
`accepted` or `rejected`, after any case that leaves a file open, when the
next case names another host, and when a kept instance no longer takes the
worker's settings. `--cases-per-instance 1` gives every case its own
instance. If a case hangs, the coordinator times out the worker and kills it
and the host process that the worker recorded, by its pid and image name;
the next case starts a new worker. Timeout
is not VBA evidence: it is treated as oracle infrastructure health. The
coordinator retries the same case up to `--timeout-retries`, except one that
timed out behind a compile-error dialog, which is not retried; if every
attempt times out, it aborts the remaining run as `outcome: "oracle_failure"`
and exits non-zero even without `--strict`. Investigate the harness/host modal state before running additional
oracle cases. If an oracle failure follows accidental parallel oracle runs,
treat the parallelism as the first suspected infrastructure cause and rerun the
cases sequentially after checking for lingering Office/VBE processes or modal
dialogs.

Only `accepted` and `rejected` are evidence outcomes. Any other worker outcome
(`timeout`, `worker_error`, setup failure, malformed output, and similar) is an
oracle infrastructure failure, aborts the remaining run, and exits non-zero.
If the compile command cannot be invoked, or any COM operation fails before a
VBE dialog is captured, the case is not treated as a VBA rejection.

For extension development, `ORACLE-FAIL` is a stop-the-line signal for the AI
agent. Do not continue adding or promoting oracle cases after it appears. First
perform a deep application inspection: check for lingering Office/VBE processes,
modal dialogs, worker stage output, command-bar invocation, dialog detection and
dismissal, timeout/retry behavior, recent harness edits, and whether the failure
is reproducible with a single focused case. Resume oracle-backed development
only after the harness failure has a concrete explanation or fix.

By default the runner is observational and exits successfully even when a
fixture expectation differs from the observed result. The local VBE
automation path is operational; use `--strict` to enforce expectations.

## Timing, reliability, and lessons learned

Hard-won notes from getting the oracle reliable on a fresh machine. The failure
mode to fear most is a **silent false "accepted"**: the worker compiles a snippet
that VBE actually rejects, fails to capture the error dialog, and records the case
as accepted. That poisons the no-false-positive discipline (a rule shipped against
a bogus "accepted" control). Guard against it with the controls below.

1. **The host must run visible.** The worker sets each host's `Visible`. When
   the host is hidden, the VBE `Compile error:` / `Run-time error` modal is not
   surfaced as a detectable visible top-level window, so the Win32 dialog watcher
   never sees it and the rejection is misrecorded as `accepted`.

2. **Accept on a positive signal where there is one.** The original fixed
   8-second watch window produced systematic false "accepted" results, because
   rejects took longer than that to show their dialog (most of it was the
   Compile command's lookup, since replaced). Accepts no longer wait out the
   window: a compile case is accepted when the Compile command turns disabled,
   which it does only once the project has compiled, and is still disabled with
   no dialog seen 300 ms later. A run case is accepted when the host's `Run` returns
   and no dialog is seen within 300 ms, because `Run` returns only after its
   dialog is dismissed. Both also require the host to be alive (see #8).

3. **`--timeout` is the single knob.** In normal mode the watch window is
   `max(12, min(22, timeout - 22))` seconds; raising `--timeout` adds startup
   headroom. For a window longer than the 22s cap (a very slow VBE), use
   `--dialog-hold-seconds`, whose branch uses `timeout - hold - 5` uncapped.

4. **The watch window is an upper bound.** A compile case that neither shows a
   dialog nor disables the Compile command waits the *entire* window, then is
   recorded as `accepted` if the host is alive. `host_start + watch + cleanup`
   must fit inside `--timeout` or the coordinator kills the worker as a false
   `timeout`/`oracle_failure`.

5. **The first run on a cold machine may time out entirely.** The first-ever COM
   Office launch (Office first-run/activation) can exceed the timeout before
   compilation even starts. Treat a *single* initial timeout as cold-start
   warmup, re-run, and expect subsequent runs to be faster as the OS/COM caches
   warm. A *systematic* timeout or `oracle_failure` after warmup is a real harness
   problem — follow the `ORACLE-FAIL` deep-inspection steps below.

6. **Sanity-check the harness before trusting a batch.** Run a couple of
   known-reject fixtures and one known-accept fixture under `--strict`; if a
   known-reject reports `accepted`, the dialog capture is broken (see #1/#2) — do
   not promote anything until it is fixed:

   ```powershell
   node syntax_corpus/oracle/run_vbe_oracle.mjs `
     --case missing_trailing_required_argument `
     --case bare_variable_statement `
     --case number_to_string_assignment --strict --json
   ```

7. **Tuning.** The default `--timeout 45` suits a warm machine; raise it for
   slow or cold machines. Accepted and rejected cases both end on a signal, so
   a shorter timeout saves nothing on a healthy run.

8. **A crash is not an accept.** Compiling some snippets crashes Excel. The
   watcher then sees no dialog, which used to be recorded as `accepted`. Before
   accepting, the worker now checks that the host process it started is still
   running and answers a COM call; if not, the case is a `worker_error` and the
   run stops. A host that hangs instead, for example behind a crash report,
   blocks that call, and the case ends as a `timeout`.
   A fixture may name exact four-part builds in `crashesHostBuilds`, for
   example `["16.0.20430.20182"]`. On a listed registered build of its host,
   the coordinator does not run it: it reports `skipped` at `known_host_crash`
   and counts it separately in text and JSON. This is no evidence, expectation
   failure or infrastructure failure; `--strict` does not fail on a skip.
   Skips are never checkpointed and cannot be promoted with `--promote-observed`.
   Without `--checkpoint`, the build is read only for hosts whose selected
   fixtures carry the field. If that lookup fails, the fixture runs as usual.
   Other builds run as usual too, and an unlisted crash still stops the run.
   The two `#Const O = Nothing` condition fixtures crash all four hosts,
   measured in Excel, Word, PowerPoint and Access 16.0 (build 20430).

## Automation Policy

The oracle must not use `SendKeys`, keyboard focus scripting, or
timing-dependent keystroke playback as evidence. Driving the VBE should use
deterministic low-level hooks: COM object model calls for file/VBE actions,
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
is treated as compile rejection. When the Compile command turns disabled with
no dialog, or no compile dialog appears before the popup watch period ends, the
case is recorded as compile-accepted, but only if the host is still running and
answering COM calls. If the host crashed, the case is a
`worker_error`.

Some source never reaches the compiler: the VBE refuses to insert it, as its
editor would, for example with `Too many line continuations`. A refusal on a
short list of such messages is recorded as `outcome: "rejected"` with stage
`insert_refused` and `evidencePhase: "compile"`, in any mode. Any other error
while adding a module stays a `worker_error`; add a message to the worker's
list only once it has been seen.

By default a fixture supplies one standard-module `source` string. For
document-module behavior, a fixture may instead supply a `modules` array. Each
entry has `name`, `type` (`standard`, `class`, `userform` or `document`), and
`source`;
mark the standard module containing the `entryPoint` with `entry: true` when it
is not the first standard module. A `document` entry names one of the new
file's own document modules (`ThisWorkbook` or `Sheet1` in Excel,
`ThisDocument` in Word), whose code the worker replaces in place, since a
document module cannot be added. A new presentation has none. In Access a
`document` entry named `Form_<name>` or `Report_<name>` makes that form or
report, gives it a module and saves it, then takes the case's code. A
`userform` entry adds a UserForm, whose module holds the case's code. This lets the oracle verify class-module member behavior
without changing production code.

A module entry can also carry what its designer or attributes declare
(issue #1419). Each field is read back after setup, and a read-back that
fails ends the case as a `worker_error`, never as evidence:

- `designerControls` on a `userform` entry: `[{ "name", "progId" }]`. The
  worker adds each control to the form's designer, then finds it by name
  and checks its type. Only ProgIDs whose read-back type has been measured
  are taken: `Forms.CommandButton.1` and `Forms.ComboBox.1`, listed in the
  worker's `$DesignerControlInterfaces`. An empty array is a form with no
  controls, which tells the analyzer the control list is complete.
- `Attribute` lines in a module's source, such as
  `Attribute VB_PredeclaredId = True`. The module is imported rather than
  typed in, as procedure attributes already were. It is then exported again,
  and every attribute line of the case must be in the export.
- `codeName` on an Excel or Word `document` entry: the module's new code
  name, `ThisWorkbook` renamed `BookCode` say. The worker renames the
  module before it takes the case's code, and reads the name back. The
  analyzer gets the module under that name, and as the object the original
  name was: a workbook, a document or a worksheet. Access names a form's
  or report's module itself, so an Access case with one ends as a
  `worker_error`.
- `documentControls` on the case: `[{ "name", "progId", "layout" }]`, ActiveX
  controls put on the document before any code (issue #1481). In Word,
  `layout` is `inline` in the body (the default), `floating` on the page, or
  `header` in the first section's primary header. In Excel they go on the
  first worksheet, so `Sheet1` is the module that has them. In PowerPoint
  they go on the first slide, added blank when there is none, and the
  control brings the slide's module, `Slide1`, into the project. Each is
  read back by name and ProgID. Access has no such controls. The hosts
  refuse them while ActiveX is disabled in the Trust Center, which ends the
  case as a `worker_error`. The worker never renames a control, since
  setting its name through late binding is ambiguous, so name each one as
  the host does: `CommandButton1`, then `CommandButton2`.
- `projectReference` on the case: `{ "name", "modules" }`, another VBA
  project the case's project references (issue #1503). The worker makes it
  as a file of its own in the same host, names its project, adds `modules`
  as it adds a case's, saves it to the temp folder and references it with
  `References.AddFromFile`, then reads the reference's name back. The file is
  the kind each host can reference: a workbook in Excel, a template (.dotm)
  in Word, an add-in (.ppam) in PowerPoint, a database in Access. Word and
  PowerPoint refuse a reference to a document or presentation. Access holds
  one database at a time, so it closes the case's, makes the provider, and
  opens the case's again. The provider is deleted after the case.
  `tests/helpers/oracleModuleInput.ts` reads it for the analyzer as the
  engine reads a referenced project.
- `implicitMembers`, `designerClass` and `predeclaredId` are the analyzer's
  side of the same facts. `tests/helpers/oracleModuleInput.ts` hands them,
  and the controls `designerControls` names, to the analyzer, and every
  corpus sweep goes through it.

After a compile dialog, the worker records the code pane's selection as
`selectionAfterDialog`: the module, and the 1-based start and end lines and
columns. It is read after the dialog closes, so it is where the VBE left the
selection, not proof of what it highlighted while the dialog was up. A
selection that cannot be read is left out and fails nothing.

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

The worker keeps the host visible but does not require foreground focus for normal
compile or run evidence. Dialog detection is owned by Win32 window enumeration
for the recorded host process; timeouts remain only a fallback and never
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
because VBE runtime dialogs vary in their button command IDs. A `Compile error:`
dialog during a run is recorded as `outcome: "rejected"` with stage
`compile_dialog`. Once it has recorded either dialog in a run case, the watcher
ends that host instance: the dialog is the case's evidence, and the host is
replaced after every run case anyway. Waiting for `Run` instead can take long:
Excel's `Run` does not return after a compile error is dismissed, and in
Excel 16.0 (build 20430) a run-time error in a `Function` declared `As String`
crashes Excel once the dialog is dismissed, so `Run` returns only when Windows
Error Reporting has finished, 6 to 10 s later. Use runtime
fixtures sparingly, only when a deterministic runtime error changes diagnostic
severity or corpus truth.

An entry point may be a `Function`. If it returns a Date beyond what .NET can
hold, the macro has run but reading its return value fails with
`Not a legal OleAut date.`; the worker treats exactly that error as a normal
return, so a captured dialog still rejects the case and an accept still
requires the host to be alive. Any other error from `Run` is handled as before.

During compile fixtures, the worker starts a Win32 dialog watcher before invoking
VBE Compile. A dialog owned by the disposable host process with title
`Microsoft Visual Basic for Applications`, class `#32770`, and child text
containing `Compile error:` is recorded as `outcome: "rejected"`; the watcher
captures the dialog text, dismisses it with direct Win32 button/window commands,
verifies the dialog is no longer visible, and only then writes the oracle result
file. It does not reopen the Debug menu after dismissal; the case's file
is closed immediately afterward. Timeouts remain a fallback for hangs
or unrelated modal prompts, not the primary rejection signal. A timeout never
means accepted or rejected; after the retry budget is exhausted, it means the
oracle itself needs attention.

Use `expected: "observe"` when the case exists to collect behavior but the repo
does not yet assert a specific result.

Every oracle case must also declare `provenance`:

- `vbe-oracle-verified` for accepted/rejected expectations that are intended to
  assert Office VBE behavior.
- `observed-not-asserted` for `expected: "observe"` cases.

Every oracle case must declare the phase it proves and the meaning of that
observation:

- `evidencePhase: "compile"` for VBE Compile command evidence, a compile
  error raised while running, or the VBE refusing to insert the source.
- `evidencePhase: "runtime"` for evidence from the host's `Run`.
- `mode: "compile_then_run"` may produce either compile or runtime evidence
  depending on where the VBE rejects the snippet.
- `diagnosticMeaning: "compile-error"` when VBE Compile rejects the source.
- `diagnosticMeaning: "runtime-error"` when the source compiles but running the
  entry point raises a deterministic VBE runtime dialog.
- `diagnosticMeaning: "compile-valid"` or `"runtime-valid"` for controls that
  prove accepted behavior.
- `diagnosticMeaning: "observation"` for observe-only fixtures.

Oracle provenance is case-level. The file-level entry in
`syntax_corpus/corpus_provenance.json` only records that this fixture file is
owned by the oracle workflow.
