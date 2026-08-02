# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [3.1.3] - 2026-08-02

### Fixed

- **Re-running Analyze Workbook on an unchanged workbook is now instant, and
  runs no longer intermittently stall part-way.** Each completed analysis
  caused the explorer to fire tree updates a couple of seconds later
  (protection probes), the results panel treated those as a reason to
  re-analyze in the background, and the next explicit run then queued on the
  analysis worker behind that redundant work - stalling at the largest module,
  sometimes for several seconds. Two changes: analysis results are now cached
  per workbook against a content fingerprint (open-editor edits included) plus
  the severity settings, so an unchanged workbook answers in milliseconds with
  "Analysis up to date" instead of re-analyzing; and the results panel no
  longer re-analyzes on tree events at all, since protection badges do not
  change analysis inputs. Editing any module or changing severity settings
  still re-analyzes, incrementally.

## [3.1.2] - 2026-08-01

### Fixed

- **Analyze Workbook's progress notification counts up instead of appearing
  stuck.** Since 3.1.1 moved per-module analysis onto the worker thread, every
  per-module progress report fired within the first few milliseconds and the
  100ms display throttle dropped them all - so the toast sat on "Reading VBA
  modules..." for the whole run, which read as a hang even though the analysis
  itself is as fast as before (and roughly twice as fast on unchanged re-runs).
  Progress now reports each module's completion - "Analyzed <name> (n/N)" - so
  the counter climbs while a large module is still being analyzed.

## [3.1.1] - 2026-08-01

### Performance

- **Analyze Workbook no longer hitches the editor while it runs.** The command
  analyzed each module on the extension host - on a workbook containing a very
  large module, that meant sub-second UI stalls mid-command. Per-module
  analysis now runs on the analysis worker thread (the same path live
  diagnostics took in 3.1.0), with an identical in-host fallback if the worker
  is unavailable. Re-running the command on a workbook you have been editing is
  also faster now: the worker keeps per-module incremental state, so an
  unchanged module is not re-analyzed from scratch, and an unchanged workbook
  skips the module-source transfer entirely.

## [3.1.0] - 2026-08-01

### Performance

- **Typing in very large modules no longer lags.** The "local" diagnostics
  pass ran a full module analysis on the extension host - about 700-945 ms on
  a real 24,000-line class - 90 ms after every pause in typing, so keystrokes
  could jam behind it. Both diagnostic passes now run on the analysis worker
  thread when it is healthy: the host runs no analysis at all while you type,
  and the worker's incremental state means the follow-up pass re-analyzes only
  what changed. This also removes a similar freeze on first open of a large
  module. If the worker is unavailable, the in-host pass is paced well behind
  the typing burst instead of fired into every pause.

- **Repeated workbook reads are now nearly free.** XLIDE keeps one parsed
  workbook per file (validated against the file's timestamp and size on every
  call, so Excel, git, or another window writing the file is always seen).
  Expanding a 42-module workbook in the explorer went from paying a full parse
  on every one of its 44 internal reads to paying exactly one; warm reads
  dropped from ~4 ms to ~0.02 ms. Saves invalidate the cache atomically, and
  mutating operations never share the cached parse.

### Removed

- **Live Share integration.** Guest browsing was never functional: it required
  a shared-service API that Microsoft restricts to approved extensions, so the
  guest side always came up empty. The integration, its remote tree nodes and
  URI routing, the guest status bar item, and the `vsls` dependency are gone,
  along with the engine fallback path that existed only to serve it.

### Internal

- Retired the last bridge-era names and dead code: `BridgeError` is now
  `WorkbookEngineError`, the unused child-process watchdog option is gone, a
  full unused-export audit removed the final two dead functions, and the
  remaining Python-era test fixtures were updated. The import graph is fully
  connected and the engine's only Node dependencies are `fs`, `path`, and
  `zlib`.

## [3.0.0] - 2026-08-01

### Changed

- **XLIDE no longer needs Python. It reads and writes workbooks itself.**
  Every workbook operation used to run through a long-lived Python child
  process, which meant every install began with a setup gate: find an
  interpreter, install two libraries, and recover when either went missing.
  That whole layer is gone. XLIDE now implements the container formats
  directly - the [MS-CFB] compound file, [MS-OVBA] compression, the VBA
  project's dir/PROJECT/module streams, and the OOXML package and worksheet
  surface - in the extension itself.

  For you that means: nothing to install beyond the extension, no Setup
  section, no interpreter to configure, no backend to restart, and the
  workbook tree available the moment XLIDE loads.

  The engine was validated against the previous implementation before the
  swap: 77 modules across 4 real workbooks parse identically, compression is
  byte-identical for every module source, container round-trips are lossless
  and idempotent, and workbooks mutated through every write path open cleanly
  in real Excel with their VBA readable through the VBE.

### Performance

- **Reads are about twice as fast, and saves nearly twice as fast.** Opening a
  VBA project used to decompress every module even when the caller only wanted
  names, types, or one module out of forty - about three quarters of the cost
  of every read. Module source is now decompressed on first use, and module
  classification reads only the header. Saving re-deflates the rebuilt VBA
  project at a compression level tuned for the write path, which costs about
  half the time for a workbook well under a percent larger.

  On a 2.4 MB workbook with 42 modules: listing modules 8.1 ms -> 4.0 ms,
  reading a module 8.5 ms -> 4.0 ms, saving an edited module 34.5 ms -> 17.0 ms,
  creating one 61.5 ms -> 37.2 ms.

### Added

- **New analysis rule: Friend members reached through a late-bound receiver.**
  Friend members are not on a class's dispatch interface, so reaching one
  through a `Variant` or `Object` receiver - including a `Collection` element,
  since `Collection.Item` returns Variant - raises run-time error 438. VBA
  compiles it without complaint, so the failure only appears on the first
  execution that reaches the call, which is how this survives in shipped code.

  The rule reports only names that resolve exclusively to Friend members of
  project classes, and stays silent when the same name is Public anywhere, is
  part of the Excel object model, or belongs to a VBA runtime object - the
  runtime type inside a Variant is unknowable, so anything less certain would
  be a guess. Backed by three Excel/VBE oracle cases, including the control
  proving that the same read through a typed local runs clean.

  Scanning five real workbooks with it turned up three genuine instances of
  the bug in one function.

### Fixed

- **The VBA test host now proves it owns its Excel instance.** It quits that
  instance when a run ends and kills it on a hang, so attaching to an Excel you
  already had open would put your unsaved work at risk. It now snapshots the
  running Excel processes before creating its own and refuses the run rather
  than touching one of yours. The owned instance is also tied to the host
  process, so a crashed run cannot leave a hidden Excel holding your workbook
  open.

- **Fewer ways for a test run to wedge on a dialog.** Macro security is
  suppressed for the host's own instance (that prompt is Excel-owned and cannot
  be dismissed by anything), alert suppression is re-asserted before closing and
  quitting in case test code turned it back on, and the modal watcher now covers
  opening and teardown instead of only macro execution.

### Removed

- The `xlide.pythonPath` and `xlide.checkPythonLibraryUpdates` settings, the
  Setup section of the sidebar, the Python setup commands, and the
  `xlide_runOpenpyxl` agent tool. Sheet and cell operations are covered by
  `xlide_readCells`, `xlide_readFormulas`, and `xlide_writeCells`.

### Notes for AI agent users

- Tool descriptions now state plainly that VBA lives in the workbook, reachable
  either through the XLIDE tools or the `xlide-vba://` virtual file system where
  saving a module writes straight into the workbook - and that exported
  `.bas`/`.cls`/`.frm` files in a project are one-way artifacts that no workbook
  reads back, so editing one changes nothing. Agents should leave them alone
  unless you explicitly ask about the export artifacts.

## [2.6.1] - 2026-07-25

### Fixed

- **Workbook protection/signature detection works with pyOpenVBA 3.1.0.**
  pyOpenVBA 3.1.0 removed an import alias XLIDE relied on, so the signed/locked
  workbook badges and the `getProtectionInfo` agent tool failed for fresh
  installs (and for anyone using the new sidebar Update button). XLIDE now
  imports from the API's canonical home, verified against both pyOpenVBA 3.0.1
  and 3.1.0.

## [2.6.0] - 2026-07-25

### Performance

- **Very large modules are now fast - roughly 100x less editor-thread work.**
  A real-world 24,000-line class module used to cost ~10 seconds of analysis on
  the editor thread per pass, re-running after every change; typing, completion,
  and hover all contended with it. Three layers of work landed:
  - The analyzer itself is ~9x faster: several quadratic hot paths (conditional-
    compilation activity, doc-comment extraction, per-statement re-lexing,
    symbol and type-name lookups) were replaced with one-pass indexes and
    shared-token derivation - verified byte-identical diagnostics across three
    real-world codebases at every step.
  - Editing only a procedure body now re-analyzes just that procedure and
    splices the cached results for the rest of the module, with automatic
    full-pass fallback on any declaration/signature/directive change.
  - Full analysis passes run on a background worker thread, so even the first
    full pass of a huge module never blocks typing. If the worker cannot start,
    analysis transparently falls back in-process.

### Added

- **The sidebar offers a one-click Update when the required Python libraries
  are outdated.** Once per session (opt-out: `xlide.checkPythonLibraryUpdates`)
  XLIDE compares the installed pyOpenVBA/openpyxl versions against PyPI; when
  one is behind, the Required Python Libraries row shows a yellow Update
  Available state whose button upgrades both and restarts the backend.
  Outdated-but-installed libraries never block any feature.

### Fixed

- **Opening a large class module no longer flashes hundreds of bogus "'Me' is
  only valid in a class..." errors.** Passes that ran before the workbook
  reported the module's kind analyzed it as a standard module; they now wait
  and retry instead of publishing wrong-kind errors, and the fast per-keystroke
  pass reuses the last known module kind.

### Changed

- **Fewer popup notifications.** Success summaries whose outcome is already
  visible elsewhere (analysis results, validation passed, export/import
  success, "copied to clipboard") now show as a transient status-bar message
  instead of a bottom-right toast. Failures and actionable prompts are
  unchanged.

## [2.5.13] - 2026-07-09

### Fixed

- **The workbook tree recovers after a transient read failure instead of
  staying broken until a window reload.** If a tree expansion failed - most
  commonly because Excel briefly holds an exclusive lock on the workbook while
  opening or saving it - the workbook node was left permanently empty, and even
  closing Excel did not bring it back. A failed module or procedure listing now
  shows an in-tree "Load failed - click to retry" item whose tooltip carries the
  actual error (plus a hint about Excel locks); clicking it retries just that
  listing without collapsing the rest of the tree.

_Note: v2.5.12 was tagged but never published; its fixes ship with this release._

## [2.5.12] - 2026-07-08

### Fixed

- **Two analyzer false positives on legal VBA.**
  - A `Case Is > 5` comparison clause (MS-VBAL 5.4.2.10) is no longer reported
    as an invalid operator sequence: `Is` in a Case clause is grammar, not the
    object-identity operator. Operator runs inside a Case body still flag.
  - `s = 3000000000&"x"` is no longer reported as juxtaposed values: the VBE
    reads the glued `&` as concatenation because the digits overflow Long
    (oracle-verified accepted), so a `&`-suffixed integer literal never
    provably ends a value. The in-range form (`n = 5& 1`) is deliberately
    under-reported rather than risking the false positive; other suffixes
    (`n = 5% 1`) still flag.

### Added

- **A corpus-wide no-false-positive sweep over the accepted oracle cases.**
  Every VBE-verified accepted case in the corpus is analyzed with its full
  cross-module project context; any compile-error diagnostic (or runtime-error
  diagnostic on a runtime-verified case) firing on one now fails the test
  suite. This is the gate the juxtaposition false positive slipped past.

## [2.5.11] - 2026-07-08

### Fixed

- **Five analyzer false positives found by analyzing a real-world JSON library.**
  Valid VBA that the VBE compiles is no longer flagged as an error:
  - `Open path For Binary Access Read As #f` no longer reports the `Access`
    keyword as an undefined variable.
  - Assigning a Byte array to a String (`s = bytes`, or a `Function ... As
    String` returning its Byte-array buffer) - VBA's documented encoding
    conversion - is no longer reported as an array-to-scalar error.
  - `If n > 0 Then ReDim a(1 To n)` no longer reports the ReDim's own target as
    an unallocated-array access (single-line `If` Then/Else arms are now
    recognized as allocation sites).
  - `ReDim rows(1 To n) As Collection` no longer reports the type name in the
    `As` clause as an undefined variable.
  - A parenless call whose first argument is parenthesized - `AssertTrue
    (cond), "message"` - now counts all its arguments instead of reporting a
    wrong argument count.
  In every family the neighboring genuine errors still flag (e.g. `x = Access`,
  a Long array assigned to a String, real unallocated accesses, undeclared
  ReDim bounds, and truly missing arguments).

- **macOS Python setup now offers Download and detects an install
  automatically.** On a Mac without Python, macOS's built-in `python3` stub made
  the setup failure unrecognizable, so the sidebar offered only "Set Path" and
  never noticed a subsequent install. The stub is now classified correctly: the
  sidebar shows a Download button (python.org) with a plain-English explanation,
  re-checks every few seconds so a new install is picked up without a window
  reload, and well-known install locations (Homebrew, python.org) are probed
  directly - Homebrew's directory is often not on VS Code's PATH. An explicit
  `xlide.pythonPath` and a workspace `.venv` still take precedence; Windows and
  Linux behavior is unchanged.

## [2.5.10] - 2026-06-30

### Fixed

- **Indexed collection access no longer reports a false type mismatch.** Assigning
  a single element pulled from a collection - `Set ws = ThisWorkbook.Sheets("x")`,
  `Set c = ws.ChartObjects(1)`, `Set pf = pt.PivotFields(1)`, `Set s = ch.SeriesCollection(1)`,
  and the rest of the `Collection(Index)` family - resolved to the *collection*
  type instead of the element and was wrongly flagged `assignment-object-type-mismatch`.
  These now resolve to the element type (and member completion / chaining works
  through them, e.g. `ws.ChartObjects(1).Chart`), while a genuine mismatch such as
  assigning a `Range` to a `Worksheet` still fires.
- **Upgrading the extension no longer floods the Problems panel with settings
  errors.** New settings added in a release, an older workbook `.xlide_settings.json`,
  or an analysis rule code you set under a previous version that has since been
  renamed no longer blast a wall of "setting not set correctly" errors. Unset or
  defaulted values are never validated, stale workbook settings and rule codes are
  tolerated silently, and a value you genuinely set to something invalid is now a
  single gentle warning instead of a hard error on every open module.

## [2.5.9] - 2026-06-28

### Added

- **Juxtaposed expressions in an assignment are now flagged as syntax errors.**
  A statement like `n = 1 n 1` (two value expressions with nothing between them)
  no longer slips through silently - it is reported as "expected end of
  statement", matching what the VBA compiler rejects. The check is narrowly
  scoped to an assignment right-hand side, so it never fires on valid code such
  as a call written with a space (`Foo (x)`), jagged-array access (`arr(1)(2)`),
  or a type-declaration suffix (`Count&`).

### Changed

- **AI-assistant tool descriptions now keep agents on the live XLIDE module
  tree.** The VBA read/write/rename/delete language-model tools spell out that
  the workbook is a binary container, so VBA changes must go through XLIDE rather
  than editing the `.xlsm`/`.xlsb` file or its exported `.bas`/`.cls` artifacts
  directly.

### Fixed

- **Smart Tab no longer indents the whole line when you press Tab at the end of
  a line.** Tab now inserts a tab character when the cursor is at or past the
  line's content; it still indents the line when the cursor is in the leading
  whitespace, on a blank line, or when a multi-line selection is active.
- **The XLIDE Tests panel refreshes when you save a module.** Editing and saving
  a VBA module of the workbook now updates the discovered-tests view instead of
  waiting for a manual refresh.
- **A range of robustness fixes from two full-codebase adversarial review
  passes** - covering module export/sync edge cases, the Python bridge restart
  path, output-log path redaction, atomic-save file permissions, and several
  smaller correctness issues.

## [2.5.8] - 2026-06-28

### Changed

- **Live diagnostics hold "still-typing" syntax errors on the current line.** A
  syntax error on the line your cursor is on is suppressed until you leave the
  line, matching the VBE which validates a line only once you move off it - so a
  half-typed `If` no longer flashes a red squiggle while you are still writing
  it.

### Fixed

- **Member completion after a leading dot in expression position.** Inside a
  `With` block, typing a leading `.` where an expression is expected (e.g.
  `For Each wb In .`) now offers the With-target's members.
- **Resolved deferred findings from the adversarial code review.**

### Performance

- **Less redundant work on the per-keystroke analyzer and diagnostics hot
  paths**, reducing latency while typing in large modules.

## [2.5.7] - 2026-06-27

### Added

- **A setting to turn the explorer's auto expand/collapse on or off**
  (`xlide.explorer.autoExpandCollapse`, on by default). When on, the XLIDE
  explorer reveals the active module and collapses the others as you switch
  editor tabs. Turn it off to keep the tree exactly as you arrange it - switching
  tabs and expanding nodes never auto-collapses anything.

### Fixed

- **Opening a module could intermittently fail with "command not found"** (most
  often with several workbooks open) and stay broken until the window was
  reloaded. Command and virtual-filesystem registration are now resilient to a
  partial or stale re-activation, so one failed registration no longer strands
  the rest.
- **Switching between two workbooks' tabs now reliably collapses the workbook you
  left** - a strict one-workbook-at-a-time accordion. Previously the workbook you
  switched away from could stay expanded.

## [2.5.6] - 2026-06-27

### Added

- **Excel Integration settings (`xlide.excelIntegration.*`).** A new settings
  group controls what XLIDE does when a module save, add/rename/delete, F5, or
  Open Workbook needs to write a workbook that Excel holds open: a coordination
  mode (`block`, the default and safest; `closeTracked`; `closeForce`),
  `trackOpenedWorkbooks`, `reopenAfterClose`, `reopenMode`
  (`lastState`/`readOnly`/`readWrite`), and `reopenReadOnlyAfterSave`. XLIDE can
  now gracefully close and reopen a workbook in Excel to complete a write, and
  refresh Excel's stale read-only view after a save.
- **Comment continuation on Enter.** Pressing Enter on a VBA comment line starts
  the new line with an apostrophe to continue the comment. Controlled by
  `xlide.editor.continueCommentOnNewline` (default on) and
  `xlide.editor.mirrorCommentSpacing` (default on), which lines the continuation
  up with the previous comment's text.
- **New compile-error diagnostic `if-reserved-keyword-in-condition`**
  (VBE-oracle-verified): flags a reserved keyword in a block-If condition, such
  as `If If True Then` or `If True Then Then`.
- **New compile-error diagnostic for a standalone multi-argument parenthesized
  call** used as a statement (for example `mySub("a", "b")`).
- **Python availability pulse:** after you install Python, XLIDE detects it
  without a manual reload.

### Changed

- **F2 Rename now updates call sites across modules**, not only within the
  current module.
- **Inline (ghost-text) suggestions are turned off for VBA modules** so they no
  longer take priority over XLIDE's completion list.
- **The explorer tree only auto-collapses a workbook when focus moves to a tab in
  a different workbook**, not on transient focus loss.
- **A single, clear notification on a locked save or open**, instead of stacking
  XLIDE's own warning on top of VS Code's native one.

### Fixed

- **F5 on a procedure with required parameters now explains the problem** (it
  names the required parameters and how to run it) instead of failing with an
  opaque COM error.
- **F5 rides out a busy Excel** (`RPC_E_CALL_REJECTED`), such as a `MsgBox` left
  open from a previous run, and reports a clear "Excel is busy" message if it
  persists.
- **Saving a module while the workbook is open read-only in Excel** no longer
  surfaces a raw error, and F5 on a read-only workbook no longer races its own
  background refresh.
- **Transient file-lock failures while reading now retry** (Excel briefly locks
  the file while it saves) instead of a spurious "cannot read" error.
- **Race-condition hardening** across the Excel coordination paths, the explorer
  cache, dirty-module backups, and module export/import (per-folder
  serialization).

## [2.5.5] - 2026-06-20

### Changed

- **A procedure closed with the wrong `End` keyword is now a warning, not an
  error.** VBE accepts `End Sub`/`End Function`/`End Property` interchangeably as
  procedure closers (oracle-verified: `Property Get … End Function` compiles), so
  this is a style mismatch rather than a missing-closer compile error. The new
  `mismatched-end-keyword` warning fires on the opening line, the procedure is
  still treated as closed, and the existing quick-fix swaps the keyword to match
  (e.g. `End Function` → `End Property`). Genuinely unclosed procedures and stray
  closers remain errors (`missing-block-closer` / `unmatched-block-closer`).

## [2.5.4] - 2026-06-20

A conditional-compilation accuracy patch for the modern cross-compiler idiom that
guards twinBASIC-only intrinsics behind `#If TWINBASIC Then ...`. fastjson's
LibJSON drops from 2 reported errors to 0.

### Fixed

- **`TWINBASIC` now defaults to `False`.** It is a compiler auto-constant defined
  only by the twinBASIC compiler; in Excel VBA it is undefined (False), so its
  `#If TWINBASIC Then` branches are inactive. Those branches use twinBASIC-only
  intrinsics (e.g. `PutMemPtr`/`GetMemPtr`) that are genuinely undefined in VBA,
  so analyzing the dead branch produced spurious `unknown-call` errors. Unlike a
  genuine unprovable host flag (which stays `unknown`), `TWINBASIC`'s value in
  VBA is known.
- **A boolean `#Const` now compares equal to its VBA numeric value** (`False = 0`,
  `True = -1`). Conditions such as `#Const Windows = (Mac = 0)` and
  `#If Windows And (TWINBASIC = 0)` previously mis-evaluated (a boolean was never
  equal to `0`), which could deactivate a live branch and hide its declarations —
  surfacing as a false `unknown-call` for a procedure defined inside it. Both are
  VBE-oracle verified.

## [2.5.3] - 2026-06-20

A diagnostics accuracy patch that clears a large class of false positives on
pointer/memory-heavy VBA (e.g. fastjson's LibJSON dropped from 61 reported
errors to 2), all VBE-oracle verified.

### Fixed

- **Hidden VBA built-ins are no longer reported as undeclared.** The runtime
  catalog gained the pointer functions `VarPtr`/`StrPtr`/`ObjPtr`, the
  byte-string family `LenB`/`LeftB`/`RightB`/`MidB`/`InStrB`/`AscB`/`ChrB` (and
  their `$` variants), and the `vbLongLong` `VbVarType` constant — all real,
  always-available built-ins that VBE compiles under `Option Explicit`.
- **`scalar-redim` no longer flags a member-array ReDim.** `ReDim x.arr(...)`
  (or `x!arr(...)`) resizes the dynamic-array *member*, not the container `x`;
  the rule mistook the qualifier for the array being resized. Qualified ReDim
  targets are now skipped.

### Internal

- Added four VBE-oracle fixtures (the pointer functions, the byte-string family,
  the `vbLongLong` constant, and the UDT member-array ReDim) and wired them into
  the `undeclared-variable` and `scalar-redim` audit entries.

## [2.5.2] - 2026-06-20

A diagnostics accuracy patch driven by three false-positive-shaped findings on
real stdVBA modules, each adjudicated against the Excel/VBE oracle.

### Fixed

- **`required-param-after-optional` no longer flags a Property Let/Set value
  parameter.** The mandatory value parameter of a `Property Let`/`Property Set`
  may legally follow an `Optional` index parameter (e.g.
  `Property Let X(Optional ByVal i As Long, ByVal v As Long)`); only non-value
  parameters obey the required-after-optional ordering. An interior required
  parameter after an `Optional` one is still flagged. (VBE-oracle verified.)
- **`exit-wrong-proc` no longer flags `Exit Function`/`Exit Sub` inside a
  `Property Get`.** A `Property Get` is value-returning, and VBE accepts every
  `Exit` kind there; `Property Let`/`Set`, `Sub`, and `Function` still require
  the matching keyword. (VBE-oracle verified across the full truth table.)

### Changed

- **`set-required` is reclassified from a compile error to a deterministic
  runtime error (Run-time error 91).** A bare assignment to an object target
  (`obj = …` without `Set`) compiles cleanly and fails only when it executes —
  VBE-oracle verified across `= Null`, `= New`, and Function-return-name
  assignment. The finding itself is unchanged (it is still a real bug, e.g.
  `protGetNextDescendent = Null` should be `Set … = Nothing`); only its evidence
  label is corrected, matching `object-variable-not-set`. The rule may now be
  downgraded to a warning via settings.

### Internal

- Added eight VBE-oracle fixtures covering the property value parameter, the
  `Exit`-in-Property truth table, and the missing-`Set` runtime-91 cases, and
  promoted the three rules to `vbe-oracle-verified` in the diagnostic influence
  audit.

## [2.5.1] - 2026-06-20

A correctness patch for the diagnostics engine.

### Fixed

- **`module-declaration-in-procedure` no longer suppresses a module's entire
  diagnostic output.** When a comment-only line was the first statement after a
  conditional-compilation directive (`#If` / `#Else` / `#End If`) inside a
  procedure body — a common 32/64-bit pattern — the rule's alternative-header
  probe tokenized the line to an empty list and threw a `TypeError`. Because
  `analyzeModule` converts any rule exception into an empty result, that single
  throw silently discarded **every** diagnostic for the affected module, not just
  the offending line. `tokenName` is now guarded against an empty token list, and
  its two near-identical copies are consolidated onto one implementation so the
  guard cannot drift again. (Behaviorally, the unguarded copy also mis-stripped an
  unterminated `[name` bracketed identifier; the unified version is correct.)

## [2.5.0] - 2026-06-17

Version 2.5.0 builds on the v2.4.0 static-analysis baseline with two goals, both
reaching their definition of done: completing the MS-VBAL §5.6 expression binder
and cashing it in for binder-dependent diagnostics, and finishing the
syntax-corpus mining so no source file is left un-dispositioned. The same
no-false-positive discipline applies — every shipped red has positive, negative,
and no-diagnostic controls plus a named evidence source (MS-VBAL, the Excel/VBE
oracle, or deterministic XLIDE metadata), and anything not provable stays quiet
and is deferred with a documented reason. The TypeScript suite grew to 2,071
tests; the Excel/VBE oracle now backs the diagnostics with 397 verified cases.
See `docs/static_analysis_completeness_2.5.0.md` for the auditable record.

### Added

- **`argument-shape-mismatch`** (compile-error): a bare array variable or
  same-module user-defined `Type` value passed where a parameter is a scalar — or
  a scalar (including `Variant`) passed where a parameter is declared an array —
  is a VBE compile error. The rule decides on declared shape only, never
  element-type coercion; it is oracle-verified across 9 cases and is disjoint from
  `byref-argument-type-mismatch`.
- **Operator-shape diagnostics**: `non-scalar-binary-operand` (an array or
  same-module `Type` used as the operand of a scalar-requiring operator),
  `is-operator-non-object` (`Is` on a provably scalar operand), and
  `typeof-is-always-false` (a `TypeOf x Is Y` that can never hold), all off the
  §5.6 expression AST.
- **Expression-AST structuring** for named arguments (`name:=expr`), omitted
  arguments, and bang (`!`) member access.

### Changed

- **Flow precision**: the shared dataflow now merges `If`/`ElseIf`/`Else` branch
  arms, so `object-variable-not-set` (Run-time error 91) and
  `unallocated-dynamic-array-access` (Run-time error 9) check accesses inside
  balanced `If` arms — conservatively falling back on any label, `GoTo`,
  `On Error`, or loop. Default-member (`VB_UserMemId`) matching is now a
  deterministic numeric parse against a named DISPID constant.
- **Syntax corpus fully mined**: every remaining `mining` source file is
  dispositioned and promoted to `reference`; zero files remain in `mining`.

### Performance

- A behavior-preserving pass on the per-edit analysis hot path: the operand rules
  share **one** expression-tree traversal per procedure body via an
  expression-visitor registry (rather than one walk each);
  `procedureHasUnstructuredFlow`, `moduleNonCallableSymbols`, and the same-module
  `Type`-name set are memoized per parse; `me-outside-object-module` rides the
  shared statement walk; and the branch-merge join drops a per-name allocation.
  The full 2,071-test suite stays byte-identical.

### Removed

- The one-time popup recommending users disable AI inline (ghost-text) completions
  for XLIDE VBA modules. XLIDE now coexists with inline suggestions — the `smartTab`
  keybinding yields to a visible suggestion so `Tab` accepts the ghost text — so
  the recommendation was obsolete.

### Deferred (documented)

- The comparison / Boolean / string-concatenation scalar-coercion matrix, Date
  coercion, and default-member-aware diagnostics — VBA coerces these at runtime,
  so no no-false-positive compile red is provable. Numeric/host boundary overflow
  and flow phase 2 (definite assignment) remain oracle- / binder-gated. See the
  completeness report.

## [2.4.0] - 2026-06-14

Version 2.4.0 is the static-analysis completeness release. It closes the
evidence-led completeness sprint: every shipped diagnostic now has positive,
negative, and no-diagnostic (no-false-positive) controls plus a named evidence
source — MS-VBAL, the Excel/VBE oracle, or deterministic XLIDE metadata — and
everything that cannot be proven without the expression binder or further oracle
mapping is explicitly deferred with a documented reason. The release ships with
an auditable completeness record (`docs/static_analysis_completeness_2.4.0.md`).
The TypeScript test suite grew to 1,954 tests; 342 Excel/VBE oracle cases now
back the diagnostics.

### Added

- **Numeric overflow diagnostics**, oracle-verified against the live Excel VBE:
  out-of-range `Long` and `Currency` literal assignments/arguments (Run-time
  error 6), and a new compile-time `suffixed-literal-overflow` for over-range
  `%` (Integer) type-suffixed literals. The `&` (Long) suffix is deliberately
  excluded because it is ambiguous with the concatenation operator.
- **Declaration and identifier diagnostics**: invalid identifier start/character,
  empty `Type`, duplicate `Type` fields, too-many-parameters, identifier-too-long,
  `Optional`/`ByVal` UDT-parameter constraints, and non-constant `Const`/`Enum`/
  parameter-default values.
- **Control-flow and statement diagnostics**: stray `Else`/`ElseIf` outside an
  `If`, duplicate `Case Else`, `Me` outside an object module, invalid assignment
  targets, `Open` missing `For`, `TypeOf` missing its operand, and impossible /
  too-many array-declaration bounds.

### Fixed

- The structural block-balance pass no longer emits phantom
  `unmatched-block-closer` / `missing-block-closer` errors when a `: Rem ...`
  comment trails a statement (`stripVba` now blanks `Rem` at any statement start).

### Changed

- Per-rule evidence audit across all 112 diagnostic codes; the syntax corpus and
  diagnostic-influence audit are now provenance-tracked and test-enforced.
- Minor live-diagnostics performance: two whole-source rules now use the shared
  cached tokenizer, removing redundant per-keystroke lexing.

## [2.3.0] - 2026-06-11

Version 2.3.0 is the audit-remediation release: roughly 205 commits of
performance, correctness, and packaging hardening across the analyzer, editor
providers, commands, webviews, Python bridge, and VSIX layout. The headline is
the VBA analysis engine, whose quadratic hot paths were rebuilt to scale
near-linearly with module size.

Internally, the release also restructured the code it hardened: the
diagnostics engine moved to a rule registry with per-family rule modules,
commands split into per-domain modules, webview HTML/CSS/JS moved to template
assets under `assets/webview/`, workbook module operations and the
per-workbook project index became shared services, and the C# and PowerShell
test-host sources were externalized to `assets/testhost/`. The TypeScript
test suite grew from 1,604 to 1,735 tests over the effort, joined by new
real-workbook Python backend tests.

### Changed

- **Analysis engine performance** - each analysis pass now lexes and parses a
  module once and shares statement tokens, signature tables, type
  environments, and member resolution across all rules, instead of re-lexing
  per statement and re-parsing the module for every dotted reference.
  Measured on a ~4,000-line module, full analysis dropped from ~115 s to
  ~0.95 s (~122x faster), and warm member completion resolves ~350x faster.
- **One shared project index per workbook** - diagnostics, completion, hover,
  signature help, navigation, and semantic tokens now read a single
  incremental per-workbook project index that folds in module changes,
  replacing four divergent caches that each rebuilt project context while
  typing.
- **Faster startup** - the Python backend now starts lazily on first use
  (expanding a workbook, an XLIDE command, or the sidebar becoming visible)
  instead of spawning in every window at activation, the Excel host object
  model builds lazily, and the extension bundle is smaller; measured window
  startup improved by about 12%. Workspaces without Excel workbooks never
  start Python.
- Rebuilt Smart Enter on the analyzer lexer's stripped-line substrate, so
  statements hidden in trailing `: Rem ...` comments no longer fool block
  auto-close, `With` continuation, or loop-iterator sync.
- Shrank the VSIX download from about 2.1 MB to about 536 KB: the marketplace
  icon went from 1.69 MB to a 28 KB 256x256 PNG, dev-only commands are hidden
  from the Command Palette outside development mode, and local development
  configuration and coverage output no longer leak into the package.

### Fixed

- Fixed silent data loss when saving a module that had also been edited in
  the Excel VBE: module file stats now derive from the real workbook file
  mtime, so VS Code's save-conflict detection prompts instead of silently
  overwriting the newer workbook copy.
- Fixed `getWorkbookInfo` and `listSheets` crashing on every workbook with
  openpyxl 3.1 and newer, where `ReadOnlyWorksheet` lost its `dimensions`
  property and `defined_names` became a dict.
- Fixed the lexer reading a file-number `#` (for example `Write #1, x`) as a
  date-literal opener, which swallowed the rest of the statement's tokens.
- Space-triggered completion no longer fires the full resolver cascade for
  spaces typed in ordinary code; `As `-type completion, `End`/`Exit`
  keywords, event stubs, and labels still work.
- Module-name input boxes now share one validator, closing the creation path
  that accepted digit-leading module names VBA rejects.
- The `xlide_createWorkbook` agent tool now refuses to overwrite an existing
  workbook instead of silently replacing it.
- A crashed Python backend now offers a Restart Backend action instead of
  asking for a whole window reload.
- Live Share guest writes are now recorded in the host's write audit and
  refresh the host's open editors instead of clobbering the guest's change on
  the next save.
- Stale dirty-module backups are now pruned on activation instead of
  accumulating in extension global storage.

## [2.1.2] - 2026-06-09

### Changed

- Refined the marketplace README positioning for new VBA programmers, students,
  experienced VBA developers, and agentic AI workflows over real workbook
  context.
- Clarified the XLIDE workflow around auto-detected workbooks, tree navigation,
  direct workbook read/write editing, local disk push/pull sync, detailed diffs,
  and committing exported modules to version control.
- Documented that import/export is workbook-scoped when multiple workbooks are
  present, including selected-workbook previews, sidecar settings, and import
  apply behavior.

## [2.1.1] - 2026-06-09

### Fixed

- Fixed module export/import preview coloring so create rows read as additions,
  delete rows read as removals, and overwrite/update rows keep modified
  treatment.
- Improved sync preview clarity with operation-specific diff titles, semantic
  status badges, a `Select Pending` action, and clearer copy-button tooltips for
  missing workbook or repo sides.
- Rewrote the README for a broader Excel-user audience, putting the main value
  proposition, getting-started path, and full cross-platform links first.

## [2.1.0] - 2026-06-09

### Added

- Added module-qualified procedure, function, constant, enum, and exported
  global resolution across diagnostics, completion, hover, signature help,
  rename, and Option Explicit checks.
- Added module-level XML documentation comment support so module summaries and
  member comments flow through IntelliSense surfaces.
- Added `XLIDE: Copy Performance Snapshot` and optional
  `xlide.performance.trace` output logging so editor, analysis, backend,
  virtual filesystem, sidebar, sync, documentation metadata, and VBA test
  latency can be diagnosed from one recent timing buffer.
- Expanded performance trace coverage across Python bridge startup/RPC calls,
  `xlide-vba://` reads/writes, workbook analysis stages, sidebar workbook
  discovery/rendering, module sync/export previews, analysis/test result
  webviews, documentation metadata reloads, and owned Excel test execution.
- Promoted a large wave of Excel host metadata for completion, hover, signature
  help, and receiver-chain inference, including WorksheetFunction, Pivot
  objects, QueryTables, chart internals, ShapeRange, comments, sort/filter
  helpers, form controls/OLEObjects, workbook connections, slicers/timelines,
  shape and chart formatting internals, conditional formatting subtypes,
  legacy drawing objects, sparklines, XML maps, publish objects, and web options.

### Changed

- Refocused and closed the Version 2.1.0 roadmap as the completed
  red-squiggle completeness sprint, with parser, binder, diagnostics, host
  metadata, rename, IntelliSense, and performance work prioritized by
  developer-experience impact.
- Moved all remaining open 2.1.0 backlog into `docs/roadmap_version_2.2.0.md`,
  with object member/event authoring continuation as the new 2.2 Priority 1.
- Aligned normal module rename behavior with the shared module-rename strategy
  so module-qualified references update consistently outside the class-module
  tree path.
- Moved user-facing documentation comment guidance into the user guides area and
  recorded internal oracle discipline for sequential, non-parallel runs.
- Improved workbook context and analysis responsiveness with bounded parallel
  module analysis, more explicit trace stages, cancellation/supersession
  handling, and cache pruning for editor project context.
- Changed XLIDE VBA editor defaults to use 4-space indentation and a quieter
  minimap/overview-ruler profile for large generated or workbook-backed modules.

### Fixed

- Hardened shared expression binding and name resolution for bare and
  module-qualified identifiers, procedure return variables, exported globals,
  constants, enum members, arrays, `With` receivers, source-shadowed runtime
  names, and host globals.
- Reduced false positives from live workbook testing around Option Explicit,
  module-qualified calls/reads, Attribute metadata placement, function return
  checks, array-return assignments, inactive `#If` branches, `VBA7`/platform
  defaults, parser-recovered procedure headers, and late-bound object/Variant
  receiver behavior.
- Expanded high-confidence red diagnostics for declaration order, parameter and
  property shapes, duplicate labels, enum ambiguity, conditional branch order,
  `For`/`Next` variable mismatches, `For Each` control/source typing,
  `ReDim`/`Erase`/`LBound`/`UBound` array misuse, deterministic runtime
  conversions, `Null`, `CVErr`, scalar/object `Set` usage, ByRef binding, and
  straight-line object-variable-not-set cases.
- Fixed source-shadowed runtime names, intrinsic/host-global resolution pressure,
  `With` receiver lookup, enum member ambiguity, and module-qualified exported
  member lookup in red-squiggle diagnostics and IntelliSense.

## [2.0.2] - 2026-06-04

### Fixed

- Fixed workbook analysis false positives in common line-numbered VBA, including
  `On Error GoTo 0`, `On Error GoTo -1`, `Erl`, and line-numbered
  `Select Case`/`Case` blocks.
- Hardened diagnostics and call analysis around numeric line labels for
  assignments, `Set` assignments, `Const` writes, procedure calls, argument
  counts, and mismatched `Exit` statements.
- Improved workbook analysis result navigation by queueing rapid row clicks and
  centering the selected finding in the opened module.

## [2.0.1] - 2026-06-04

### Changed

- Removed contribution links from the in-extension XLIDE sidebar so the product
  shell ends with Support actions.
- Moved open-source support links to the bottom of the repository README.
- Tightened VSIX packaging excludes so workbook files and test workbook folders
  stay out of Marketplace packages.

## [2.0.0] - 2026-06-04

Version 2 turns XLIDE from a workbook/module bridge into a fuller VBA
development environment for VS Code: project-aware editing, deterministic
analysis, previewable sync, workbook tests, safer workbook mutation, and
agent-verifiable workflows.

### Added

- **XLIDE Activity Bar/sidebar** - dedicated setup health, selected-workbook
  actions, global settings, support commands, and a stable workbook-action
  surface while keeping workbook/module navigation in the VS Code Explorer.
- **Workbook-wide VBA analysis** - `XLIDE: Analyze Workbook` opens a dedicated
  results UI with module grouping, severity filters, counts, copy/export
  actions, workbook/global rule settings, tracking controls, and click-through
  navigation.
- **Agent analysis verification** - `xlide_analyzeWorkbook` returns the same
  workbook analysis shape used by the UI:
  `{filePath, moduleCount, errorCount, warningCount, problems: [...]}`.
- **Deterministic diagnostic metadata** - rules now carry category,
  VBE-compile-equivalence, diagnostic kind, stable codes, and source labels so
  red diagnostics stay reserved for proven compile/runtime errors.
- **Expanded VBA diagnostics** - v2 adds and hardens high-confidence checks for
  block balance, unterminated strings, duplicate declarations, `Const`
  assignment, invalid procedure headers, unbalanced parentheses, declaration and
  call hygiene, argument counts and type mismatches, unknown calls, invalid
  declaration names, scalar member access, object/scalar `Set` usage, return
  assignment warnings, test marker syntax, and source-backed `member-not-found`.
- **Project-aware language service** - completions, hover, signature help,
  go-to-definition, find references, rename, semantic type coloring, and code
  actions now use the shared analyzer/project model where XLIDE can bind the
  target deterministically.
- **Syntax and editor hardening** - Smart Enter, block snippets, close-keyword
  completions, keyword casing, comment continuation, paired `For`/`Next`
  iterator edits, parser recovery, and MS-VBAL-backed keyword/token handling
  were expanded for common VBA editing flows.
- **Source-backed object/member understanding** - project classes, public
  fields, properties, return-name assignments, UDT fields, document/UserForm
  code names, known runtime signatures, and the first generated Excel host
  surfaces participate in completion, hover, navigation, diagnostics, and
  assignment validation where the surface is proven.
- **Documentation metadata** - inline `'''` XML doc comments and external
  `.vbref.xml` metadata enrich hover, completion, and signature help for
  source-backed and externally documented symbols.
- **Generated host/reference metadata slice** - generated Excel reference
  metadata is now used for the proven Excel `Workbook` surface, with coverage
  and provenance tracked separately from runtime extension code.
- **Workbook VBA test runner** - marked `@xlide-test` procedures run through an
  XLIDE-owned read-only Excel host, with `XlideAssert.bas` support, skip/xfail
  metadata, tag filters, current-module/current-test commands, rerun flows,
  output capture, artifacts, and `status_for_ci.json`.
- **Agent test execution** - `xlide_runVbaTests` runs the same workbook tests
  headlessly and returns `{ok, summary, artifacts, report}` for AI-agent and CI
  verification.
- **Previewable import/export sync** - bulk export/import and current-module
  export use diff previews, explicit apply steps, changed/skipped/removed/failed
  summaries, safe true-up behavior, and one workbook settings owner.
- **Workbook and global settings surfaces** - `xlide.openGlobalSettings`,
  workbook-facing analysis/sync settings, guarded severity overrides,
  untracked-rule controls, strict sidecar parsing, and provenance labels make
  settings explicit without weakening diagnostic determinism.
- **Code actions** - deterministic quick fixes/source actions cover the shipped
  safe edit set, including adding `Option Explicit`, fixing known call syntax,
  adding/removing `Set`, inserting block closers, moving misplaced `Option`
  statements, splitting local `Dim` initializers, adding suppression comments,
  creating safe private stubs, analyzing the current module, and exporting the
  current module.
- **Safety, trust, and recovery** - dirty `xlide-vba` backup restoration,
  explicit mutation/run commands, write audit summaries, workbook lock/open-state
  checks, COM timeout/cleanup handling, support bundles, copy diagnostics, and
  redacted support output make workbook operations easier to trust.
- **Performance budgets** - v2 records launch-facing latency budgets for
  keystroke diagnostics, module analysis, workbook analysis, project index
  rebuilds, and sidebar health refreshes; deeper performance hardening moved to
  the v2.1.0 roadmap.
- **User-facing documentation** - README and user guides now cover setup,
  workbook workflows, analysis/ignores, import/export sync, testing,
  automation/CI, safety/support, and the v2 feature surface.

### Changed

- Replaced the older dedicated `xlide.diagnostics.optionExplicit` model with
  guarded global/workbook analysis settings, including `visibleSeverities`,
  `untrackedRules`, and `ruleSeverityOverrides`.
- Consolidated workbook-specific configuration under
  `<workbook>.xlide_settings.json`; older workbook sidecar names are not part of
  the supported v2 settings contract.
- Moved remaining binder, designer/UserForm metadata, external metadata,
  host-metadata completeness, workbook-to-workbook transfer, and performance
  scale work to `docs/roadmap_version_2.1.0.md`.
- Kept hard object-member absence diagnostics limited to exhaustive receiver
  surfaces; incomplete host/designer/external surfaces can power completion and
  hover without inventing red `member-not-found` errors.
- Updated macro/test workflows so execution remains explicit, Windows Excel COM
  remains execution-only, and normal read/edit/analyze/sync workflows continue
  to use the Python backend.

### Fixed

- Malformed workbook settings sidecars now report explicit settings errors
  instead of being silently treated as empty defaults.
- Tightened Live Share guest messaging around the current supported behavior:
  guests can edit host-opened VBA buffers, but only the host can browse/open new
  workbook modules through XLIDE.
- Refined release docs and roadmap references so v2 is closed and forward
  scope points to the v2.1.0 roadmap and related sub-roadmaps.

## [1.0.9] - 2026-05-26

### Added

- **`xlide_listWorkbooks`** agent tool - discovers all `.xlsm`, `.xlsb`, and
  `.xlam` files in the workspace so an agent can find the target workbook.
- **`xlide_getWorkbookInfo`** agent tool - returns sheets, VBA modules, and
  named ranges in one round trip.
- **`xlide_listSheets`** agent tool - lists worksheet names and used dimensions
  for cell-range discovery.
- **`xlide_readFormulas`** agent tool - reads raw formula strings instead of
  computed values.
- **`xlide_runOpenpyxl`** agent tool - executes arbitrary openpyxl code against
  a workbook for worksheet-level automation.
- **`xlide_renameModule`** and **`xlide_deleteModule`** agent tools - expose
  existing module mutation support to AI agents.
- **`.github/copilot-instructions.md`** - canonical XLIDE agent workflow loaded
  automatically by Copilot in the repository.

### Fixed

- Clarified the `xlide_writeModule` tool description: writing a missing module
  name creates the module.
- Updated `xlide_readCells` and `xlide_writeCells` descriptions to direct
  agents through `xlide_listSheets` when the sheet name is unknown.

## [1.0.8] - 2026-05-26

### Fixed

- Refreshed the affected module's procedure list after Rename Symbol or manual
  source edits so the XLIDE Explorer no longer shows stale procedure names until
  a full refresh.

## [1.0.7] - 2026-05-26

### Changed

- Unified the XLIDE Explorer welcome message for host and Live Share guest
  sessions, including workbook discovery, Refresh, and Live Share notes.

## Earlier Development Notes

These notes predate the current v2 changelog structure and are kept for
historical context.

### Added

- Module tree accordion behavior and debounce.
- Add Class Module command and context-menu entry.
- Improved module type detection for UserForms and document modules.
- COM window restore/focus after opening a workbook or running a macro.
- Outline/breadcrumb symbol-kind polish for VBA declarations.
- Per-workbook module-list caching, filesystem watcher debounce, and RPC
  cancellation token support.
- Smoke test command, TypeScript/Python unit tests, and CI workflow.
- Protected-workbook editing, signature-invalidation notices, protection and
  signature badges, Validate VBA Project, and New Macro-Enabled Workbook.
- Early VBA snippets, enter-time block closing, status bar items, workbook-locked
  error UX, and marketplace display-name polish.

### Changed

- Reorganized context menus into create, edit, workbook, transfer, and settings
  groups.
- Split class-module creation into `xlide.newClassModule` instead of routing it
  through the generic new-module path.

---

*XLIDE follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) conventions.*
