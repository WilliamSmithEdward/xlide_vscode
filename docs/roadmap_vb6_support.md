# XLIDE Roadmap: VB6 Support

Classic Visual Basic 6 projects become first-class in XLIDE: a `.vbp` shows in
the tree like any other project, its `.bas`/`.cls`/`.frm` files get the same language
services VBA has (completion, hover, signature help, navigation, diagnostics),
its forms open in a designer, and a developer-only twinBASIC oracle keeps the
analyzer honest where no VB6 install exists to ask. This roadmap names the
complete system, the vertical slices that build it, the evidence each slice
must produce, and the risks. It follows the owner's decisions of 2026-09-02:

- Strict VB6. twinBASIC's language extensions are out of scope.
- No `VB6.OLB` is available. The `VB` library model comes from documents and
  from twinBASIC, carried as reported evidence with its gaps preserved.
- A `.vbp` project shows in the XLIDE tree like the Office files do; its files are
  edited as the plain text they already are.
- Forms get a designer, after they are first readable.
- twinBASIC is a developer oracle only: never bundled, never in a runtime path.
- First increment covers `.bas`, `.cls`, `.frm`; other module kinds are
  recognized in the manifest and left opaque.
- Reference material stays in this repository.
- Fixtures come from open-source VB6 codebases with clear licenses.

The evidence discipline is the repository's own (`xlide_development_principles.md`):
deterministic logic, no guessed diagnostics, a hard diagnostic only when proven,
and provenance recorded for every host fact. VB6 raises the bar on provenance,
because the two authorities that exist for it, Microsoft's archived
documentation and twinBASIC, are both documents about the runtime rather than
the runtime itself.

## The intended complete system

| Surface | VBA today | VB6 at the end of this roadmap |
| --- | --- | --- |
| Container | Office file with an MS-OVBA project inside | `.vbp` manifest over loose text files |
| Tree | file, modules, procedures | project, modules (kind from the manifest), procedures |
| Editing | `xlide://` virtual documents written back to the container | the native `.bas`/`.cls`/`.frm` files |
| Host model | Excel, Word, PowerPoint, Access, MSForms | `VBA` + `VBRUN` from the runtime's own type libraries, `VB` from documents and twinBASIC |
| Forms | MS-OFORMS engine, designer over the binary storage | VB6 forms engine over text `.frm` plus `.frx`, designer as a custom editor over the real file |
| Oracle | Excel/VBE harness (`syntax_corpus/oracle`) | twinBASIC harness beside it, verdicts typed `twinbasic-oracle-verified` |
| Run | F5 into Excel | deferred: opt-in build through a configured twinBASIC once its CLI is finished |

## What is already in place (measured 2026-09-02)

- MS-VBAL covers the VB6 grammar including `Load`, `Unload`, `Implements`,
  `Event`, `RaiseEvent`, `Declare`. A VB6 form's code-behind with control arrays
  (`Command1(1)`, `Load Command1(1)`), `App`, `Screen`, `Printer`, `Clipboard`,
  `WithEvents` and `Timer1_Timer` analyzes with zero syntax rejections when the
  controls are supplied as implicit members.
- `splitFrmSource` accepts a real VB6 `.frm` (`VERSION 5.00`, `Begin VB.Form`)
  and returns the code-behind cleanly.
- The host registry treats a named host with no model as "assert nothing", so a
  `vb6` token is safe before its model exists.
- `msvbvm60.dll` (shipped with Windows, present in `SysWOW64`) carries the
  `VBA` type library (resource 1, 29 types) and the `VBRUN` type library
  (resource 3, 96 types); both load with the pythoncom extractor already used for
  Excel. Neither contains `App`, `Screen`, `Printer`, `Clipboard`, `Forms` or the
  intrinsic controls: those are `VB6.OLB`, which is not on this machine.
- The `TWINBASIC` conditional constant is already modeled.
- Import/export already treats a folder of `.bas`/`.cls`/`.frm` as a first-class
  counterpart.

## Measured gaps the slices must close

- `hasAuthoritativeDesignerHeader` and `parseUserFormControls` fail on a real
  VB6 form: the fixture's `BeginProperty Font ... EndProperty` blocks are neither
  `Begin` nor `End` nor `name = value`, so the gate answers "not authoritative"
  and the control list comes back empty. The synthetic probe passed only because
  it had no property blocks.
- `.frx` references take the form `Caption = $"Form1.frx":0000` (a `$` prefix on
  string-valued blobs) and the blob itself is length-prefixed
  (`8f 00 00 00` then 143 bytes of text in the fixture). Pictures, icons, list
  contents and long strings all live there. There is no Microsoft specification;
  the layout is measured from fixtures, the way MS-OFORMS gaps were.
- `controlTypeFor` maps only `Forms.*` prog ids to `MSForms.*`; `VB.*` passes
  through untyped, so nothing is known about a VB6 control's members.
- twinBASIC's `VB` package documents 6 objects and 29 controls. Four of the
  controls are twinBASIC's own (`CheckMark`, `MultiFrame`, `QRCode`, `Report`) and
  must be excluded from a strict-VB6 model; `OLE` is marked a "compatibility
  stub, mostly unimplemented"; `App` marks `LogEvent`, `StartLogging`,
  `StartMode`, `TaskVisible`, `UnattendedApp` and others as "reserved for
  compatibility with VB6; not currently implemented".

## Slice 1: a VB6 project in the tree

The `.vbp` is the container. It is a text manifest; keys observed in the first
fixture: `Type`, `Form`, `Module`, `Reference`, `ResFile32`, `IconForm`,
`Startup`, `Title`, `ExeName32`, `Name`, `Description`, `CompatibleMode`,
`MajorVer`/`MinorVer`/`RevisionVer`, and the compiler switches. `Class`,
`UserControl`, `PropertyPage`, `Designer`, `RelatedDoc`, `Object` and the
`[MS Transaction Server]` section are documented shapes to be confirmed against
the wider fixture set before they are relied on.

- [x] `src/vba/vb6/vbpProject.ts`: parse the manifest into a project model
      (kind of project, module list with kind and file path, references,
      objects, startup). Every key not understood is preserved verbatim; the
      parser never invents a module the manifest does not name.
- [x] Project discovery beside the macro-container glob, a tree node per `.vbp`, module
      nodes from the manifest (`standard`, `class`, `userform`; the opaque kinds
      shown with their manifest name and no children), procedure nodes from the
      existing source scan. Clicking opens the native file.
- [x] `hostTokenForFileName` answers `vb6` for `.vbp`; a `.bas`/`.cls`/`.frm` that
      belongs to a discovered project reaches the same answer through
      `vbaDocumentLocation.ts` and the locator; `VbaHostToken` gains `vb6`,
      registered with the empty model until Slice 3.
- [x] Fixtures under `tests/fixtures/vb6/<project>/`, each with its upstream
      `LICENSE` and a `NOTICE.md` naming the source commit; first set from the
      MIT-licensed candidates already inventoried (`fafalone/RunAsTrustedInstaller`,
      `Gagniuc/Diabetes-prediction-1.0`, `RZulu54/ChessBrainVB`,
      `opensoldat/polyworks`, `Sibra-Soft/audiostation`). The license text is
      read, not inferred from GitHub's detection, before a file is vendored.
- [x] `docs/architecture.md`: container table row, tree row, and a "Files to
      keep up to date" row for VB6 projects.

Definition of done: every vendored `.vbp` loads; the tree shows its modules
with the right kinds; a manifest key the parser does not know survives a
parse-and-print round trip; tests cover a malformed manifest, a missing file,
and the opaque kinds. Met on 2026-09-02 (commits `6fbb1fb` and the slice 1b
commit after it). Learned on the way: a VB6 module's source is served as
file-aligned text (designer block blanked, attributes kept) because a file has
no virtual view to hide its header behind, and the same blanking runs on the
editor buffer before analysis.

## Slice 2: VB6 forms readable

- [x] A VB6 `.frm` header parser (`src/vba/vb6/frmHeader.ts`) that understands
      `Begin <ProgId> <Name>` nesting, `BeginProperty`/`EndProperty` blocks,
      trailing comments after values (`0   'False`), `Index` on control-array
      members, and both `"file.frx":offset` and `$"file.frx":offset` references.
      It replaces the `Forms.*`-only path in `vbaUserFormControls.ts` for VB6
      files rather than sitting beside it.
- [x] A printer that reproduces the header byte for byte from the parsed tree,
      pinned on every fixture form (the same byte-identity oracle the OFORMS
      engine uses).
- [x] An `.frx` reader for referenced blobs, measured on fixtures: string blobs,
      pictures, list contents. Unreferenced bytes are preserved untouched.
- [x] The implicit-member path types `VB.*` controls by prog id (the type the
      `vb6` model will key on); a form module's `Me` stays untyped until that
      model exists (Slice 3), where `combined:<form>|VB.Form` lands.

Definition of done: every fixture `.frm` parses and prints back identical; the
analyzer reports no `undeclared-variable` on any fixture code-behind with the
header-supplied controls; a mutation that drops the nested controls fails the
tests. Met on 2026-09-02. Measured on the way, over 34 forms from five
projects: the header grammar (OCX `Object =` lines before the form block,
`BeginProperty` groups with optional class ids nested inside each other, keys
with parentheses and dots, `'comment` glosses after numeric values, both
`"file.frx":HEX` and `$"file.frx":HEX` references), the layout the designer
writes (three spaces per level, keys padded to 16 columns, extender keys such
as `Object.Width` to 23, a trailing space after every `Begin` and
`BeginProperty` name, two spaces before a comment), and the sidecar records
(contiguous; 32-bit-length pictures and `$` strings; 8-bit-length short
strings; 16-bit-counted `List`/`ItemData` rows). The formatter regenerates
all 34 headers byte for byte from the model.

## Slice 3: the `vb6` host model

- [x] `VBA` and `VBRUN` generated from `msvbvm60.dll` resources 1 and 3 through
      the existing extractor path (`scripts/generate-host-object-model.mjs` and
      the pythoncom typelib dumper), landing as `src/analyzer/host/vb6ObjectModelData.ts`
      with `source: "typelib"`.
- [x] The `VB` library transcribed from twinBASIC's package documentation
      (`docs.twinbasic.com/tB/Packages/VB/`), filtered to VB6's real surface,
      with each member carrying `source: "twinbasic-docs"` and any "reserved"
      or "unimplemented" flag preserved; cross-read against the archived
      Microsoft pages (App, Screen, Printer, Clipboard, the control pages) with
      disagreements recorded rather than resolved by preference.
- [x] Globals: `App`, `Screen`, `Printer`, `Printers`, `Clipboard`, `Forms`, and
      the `Global` members (`Load`, `Unload`, `LoadPicture`, `SavePicture`,
      `LoadResString`, `LoadResPicture`, `LoadResData`).
- [x] Analyzer semantics for control arrays: `Ctl(i)` yields the control's type,
      event handlers carry `Index As Integer`, `Load`/`Unload` accept an
      indexed control.
- [x] Diagnostic policy: the `VB` model is reported evidence, so it drives
      completion, hover and signature help, and it never produces a red on its
      own. `member-not-found` on a `VB.*` type stays quiet until Slice 4 has
      confirmed the member's absence.

Definition of done: completion on `App.`, `Screen.`, `Printer.`, `Me.` and on
every intrinsic control type in the fixtures; hover shows provenance; the
fixture code-behind produces no diagnostics that the fixtures' own authors would
call wrong (each fixture is a negative control). Met on 2026-09-02, with two
departures from the plan recorded rather than papered over. VBA6 (resource 1)
is dumped as evidence but not modelled: the analyzer's VBA runtime already
answers those names for every host, and a second copy would collide with it.
The cross-read against Microsoft's archive is name-level: the archive keeps
the reference pages but not the "Applies To" object lists behind them, so
per-object membership rests on twinBASIC's pages; it settled every property
and method (Opacity, TransparencyKey, Anchors, Dock, App.IsInIDE and the rest
left out on the record, `DataMemberChanged` kept by name because the archive
lost its page), and it could not settle events, whose pages the archive is
also missing (Unload, the OLE drag-and-drop events, Scroll). The model as
generated: 51 types (31 `VB` classes plus `Forms` transcribed, 19 `VBRUN`
types read), 1869 members, 591 constants in 61 enumerations. A VB6 form's
event-handler stubs (`Form_Load`, `Command1_Click(Index As Integer)`) come
from the model's events, and an MDI form's from `MDIForm_`.

## Slice 4: the twinBASIC oracle

twinBASIC is a superset with published incompatibilities, so a verdict from it
is typed evidence about twinBASIC and inferred evidence about VB6. The harness
records that, never "VB6-accepted".

- [x] `syntax_corpus/oracle/twinbasic/`: a runner shaped like
      `run_excel_vbe_oracle.mjs` and a worker that stages a case as a VB6
      project in a scratch folder, imports it (`bin\twinBASIC_win64.exe import
      "<out.twinproj>" "<folder>" --overwrite`), builds it
      (`twinBASIC.exe "<proj>" --buildAndExit64`), and reads the outcome from the
      compiler's own markers (`* BUILD SUCCESSFUL *`, `*** BUILD FAILURE ***`)
      and its diagnostics. The twinBASIC location comes from an environment
      variable; nothing is bundled. A watchdog bounds every run and kills only
      the processes the worker spawned. (Measured departure: the markers and
      diagnostics never leave the IDE's panels, so the verdict is read from
      what the IDE does - see the README there.)
- [x] Case provenance `twinbasic-oracle-verified`, evidence phase `compile`, and
      the same `accepted`/`rejected` vocabulary as the VBE corpus.
- [x] A parity report: the existing 418 VBE-verified cases run through
      twinBASIC, giving a measured agreement matrix. That number is the fidelity
      the `VB` model and every VB6 diagnostic inherit.
- [x] Surface extraction: a language-server client (`--lspPort`) asking hover
      and completion over a project that references `VB`, diffed against the
      Slice 3 transcription. Members the server does not know are flagged in the
      model; members it knows and the docs do not are recorded, not added.
      (Measured departure: the IDE ships the VB package as MIT-licensed
      twinBASIC source, and its export is the exact surface, `[Unimplemented]`
      and `[Hidden]` included; `scripts/twinbasic-vb-surface.mjs` reads that
      instead of sampling a server.)

Definition of done: the harness runs end to end on the fixture projects; the
parity matrix is in `docs/`; the `VB` model carries a per-member flag from the
diff; oracle runs stay sequential and out of `npm test`. Met on 2026-09-02,
with the measured departures recorded above and in
`syntax_corpus/oracle/twinbasic/README.md`. The parity matrix
(`docs/vb6_twinbasic_parity.md`, twinBASIC BETA 983, Excel and Office
referenced): 356 of 418 VBE-verified cases agree at compile time (85.2%),
with no infrastructure failure. The disagreement runs almost entirely in the
superset direction: twinBASIC accepts 59 constructs the VBE rejects
(`Attribute` lines inside procedures, `DoEvents` as a `Call` target, a
type-declaration character beside an `As` clause, fixed-length strings of
length 0 or above 65,526, Public `Type`/`Declare`/arrays in a class module,
`#ElseIf` after `#Else`, duplicate `Case Else`, `Friend`/`Implements`/
`WithEvents` in a standard module, the VBE's implementation limits on
arguments, dimensions, identifier and line length, `Exit Sub` inside a
`Property Let`) and rejects 3 the VBE accepts (a numeric literal default on
an Optional Long, an over-range `&` literal glued to a string, a `ByVal
ParamArray`). So a twinBASIC rejection is strong evidence of a VB6
rejection (95 of 98) and a twinBASIC acceptance is weaker evidence of
acceptance (261 of 320): the oracle can confirm that a diagnostic is right
to fire far better than it can confirm that silence is right. Runs are
sequential per instance and out of `npm test`; the harness runs three
instances staggered, which is what made the full corpus a ten-minute run.
Open residual: the compiler's diagnostics never leave the IDE's panels, so
the verdict has no text or position; reaching them means the compiler's
websocket protocol, or a twinBASIC change to `--buildAndExit`.

## Slice 5: the VB6 forms designer

The document is the `.frm` itself. The designer is a custom editor over the
real file, the pattern the OFORMS designer already uses over its document:
gestures rewrite the header block, the code-behind is untouched, the tab carries
the dirty dot, Ctrl+Z is text undo, and save writes the file.

- [x] Canvas rendering for the intrinsic control set in twips, from the Slice 2
      tree and the Slice 3 property surface; menus (`Begin VB.Menu`) rendered as
      a menu bar; MDI forms recognized and shown as a plain form until MDI has
      its own treatment.
- [x] Designer ops on the text tree: add, move, resize, remove, reparent,
      z-order (block order), tab order, control arrays (`Index`), properties
      pane driven by the `VB` model. (Measured departure: the pane's vocabulary
      is the design-time set measured per kind on the fixtures, because the
      model's property list is the runtime surface with no design-time flag;
      see below.)
- [x] `.frx` writes for the string kinds the reader measured (short and long
      strings), taken by the sidecar when the document saves; pictures and
      lists are refused with the reason, never guessed.
- [x] Byte identity on save when nothing changed, and diffs confined to the
      header when something did, pinned on every fixture.

Definition of done: every fixture form opens, round-trips identically, survives
each gesture with a header-only diff, and reopens in twinBASIC's own IDE with
the same control tree (the only external designer available to check against).
Met on 2026-09-02, with these measured departures and limits. The properties
pane is driven by what the header states plus the design-time vocabulary
measured per control kind on the fixture forms (`VB6_DESIGN_PROPERTIES` in
`src/vba/vb6/frmScene.ts`, every key cross-read against the `VB` model), not by
the model's own property list: that list is the runtime surface (`hWnd`,
`Parent`, `SelText` beside `Caption`) and carries no design-time flag to
filter on, so a model-driven pane would offer properties the designer never
writes. The model does drive the pane's editors: a property declared as an
enum whose constants the model holds gets a dropdown of those constants, a
Boolean gets True/False, and the value the header takes back carries the gloss
the fixtures measured for it (`3  'Fixed Dialog`) or none. Sidecar writes are one
measured kind, a string with line breaks, in the short or long record layout
the reader measured; the records a designer places wait with the document and
reach the sidecar when it saves, so a document closed unsaved leaves the
sidecar as it was. Pictures, lists and every other record are read, never
written, and a gesture that would set one is refused with the reason. The twinBASIC reopen check is a developer check, not a test:
Slice 4 measured that the IDE imports a `.vbp` only through its own dialogs,
so the check is opening a fixture project there by hand after a gesture, and it
was not run in this slice.

## Slice 6: build and run (deferred)

twinBASIC's command-line build exists (`--buildAndExit32`/`--buildAndExit64`
on `twinBASIC.exe` in beta 983) but the finished form is still open
(twinbasic/twinbasic#508). When it lands, an opt-in "Build with twinBASIC"
command through a configured path, the way Excel is invoked for tests. Not
before.

## Risks

- **Fidelity without VB6.** No VB6 exists here to ask. Every `VB` fact is
  reported, and the oracle is a superset. Mitigation: provenance on every
  member, the parity matrix as a published number, and no red diagnostic from
  the `VB` model alone.
- **`.frx` is undocumented.** Mitigation: fixtures first, byte identity on
  untouched bytes, refusal on unmeasured blob kinds.
- **License diligence.** Mitigation: the license text read before vendoring, an
  attribution file per fixture, and no fixture from a repository without one.
- **Scope creep into twinBASIC's dialect.** Mitigation: the four twinBASIC-only
  controls and every extension are excluded by name; a case that only twinBASIC
  accepts is recorded as such and does not move the VB6 model.
- **The beta compiler in the repository.** It is git-ignored (`/twinBASIC*/`)
  and excluded from the package (`twinBASIC*/**`); the harness finds it through
  an environment variable so a clone without it still tests green.

## References

- [MS-VBAL] (`docs/[MS-VBAL].pdf`, 2025-05-18 build; also online under
  learn.microsoft.com/openspecs): the grammar, including `Load`, `Unload`,
  `Implements`, `Event`, `RaiseEvent`, `Declare`.
- Microsoft's archived Visual Basic 6.0 documentation
  (`learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/`):
  the Documentation Map `aa232759`, the Reference (one alphabetical tree, e.g.
  App Object `aa267182`, Clipboard Object `aa267187`, CommandButton Control
  `aa267189`, Controls Collection `aa445317`), and the appendix "Visual Basic
  Specifications, Limitations, and File Formats" `aa733725` with Project File
  Formats `aa241721` and Form Structures `aa241723`. Pages are `NOINDEX` but
  fetch normally; the source archive is not on GitHub, so transcription is a
  per-page crawl driven by the archive's `toc.json`.
- twinBASIC documentation (`docs.twinbasic.com`): FAQ (compatibility scope and
  known gaps), the `VB`, `VBA` and `VBRUN` package pages with per-member
  signatures and "reserved for compatibility" flags.
- `msvbvm60.dll` type libraries `VBA` and `VBRUN` (Windows-shipped).
- Rubberduck's `VBAParser.g4` (ANTLR, "based on MS VBAL", GPLv3): a grammar to
  read against, not to copy from.
- Community `.frx` notes (vb-decompiler.org and forum threads): orientation
  only; nothing in them is treated as verified.

## Files to keep up to date

| Change | Files to touch |
| --- | --- |
| VB6 project container | `src/vba/vb6/vbpProject.ts`, `src/macroContainerUi.ts` (discovery, context values), `src/analyzer/host/hostRegistry.ts` (`vb6` token), `tests/fixtures/vb6/**` with `LICENSE` and `NOTICE.md`, `tests/vb6Project.test.ts`, `docs/architecture.md` |
| VB6 form header / `.frx` | `src/vba/vb6/frmHeader.ts`, `src/vba/vb6/frx.ts`, `src/vbaUserFormControls.ts`, `tests/vb6Forms.test.ts`, `docs/architecture.md` |
| `vb6` host model | `src/analyzer/host/vb6ObjectModel.ts` and generated `vb6ObjectModelData.ts`, the generator under `scripts/`, `tests/vbaHostModels.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum), `docs/architecture.md` |
| twinBASIC oracle | `syntax_corpus/oracle/twinbasic/**`, `package.json` (`test:oracle:twinbasic`), `docs/` parity report, `xlide_development_principles.md` (oracle usage) |
| VB6 forms designer | `src/vb6FormDesigner.ts` and the canvas beside `src/vba/oforms/preview.ts`, `package.json` (custom editor), `tests/vb6Designer.test.ts`, `docs/architecture.md` |

## Releases

Slices 1 to 3 ship as 5.1.0: projects in the tree, forms readable, language
services with the `vb6` model. Slice 4 ships when its parity report exists.
Slice 5 ships as its own minor release.
