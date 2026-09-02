# VB6 reference data

The `vb6` host object model (`src/analyzer/host/vb6ObjectModel.ts`, data in
`src/analyzer/host/vb6ObjectModelData.ts`) is generated from dumps under
`reference/vb6/json`. `reference/` is not tracked - the dumps are regenerated
locally, like the Office hosts' - so this page records where they come from,
how to regenerate them, and what the filtering decided;
`docs/vb6_reference_exclusions.md` is the transcriber's own record of every
member left out or carried reserved, and `THIRD_PARTY_NOTICES.md` carries the
license notice for the transcribed text. Three libraries, two sources, every
dump naming its own.

## VBRUN and VBA: the runtime's type libraries

`msvbvm60.dll` carries two type libraries as resources. Resource 3 is
`VBRUN`, the runtime's own objects (DataObject, AmbientProperties,
PropertyBag, Hyperlink ...) and every intrinsic constant (vbKeyReturn,
vbNormal, vbModal ...). Resource 1 is `VBA`, the VBA6 language runtime
(Strings, Math, Collection, ErrObject ...).

    python scripts/dump-vb6-typelib.py

reads both with pythoncom (Windows only) and writes one JSON per type in the
shape pyVBAReference uses for the Office libraries, plus `libraryId` and
`source`. The current model came from `msvbvm60.dll` file version 6.0.98.48
(the copy Windows ships in SysWOW64). Hidden and restricted types and members
are left out, as the Object Browser leaves them out; `_index.json` lists what
was skipped.

Only VBRUN is modelled. The VBA dumps are evidence: the analyzer's VBA runtime
already answers those names for every host, and a second copy would collide
with it. `libraryId: "VBA"` dumps are the ones the generator reports as
"evidence-only".

## VB: transcribed from twinBASIC's documentation

`VB.OLB` declares App, Screen, Printer, Clipboard, the Global object, Form,
MDIForm and the intrinsic controls. It ships only with the Visual Basic 6.0
IDE and is not available here, so the `VB` library is transcribed from the
package documentation of twinBASIC, a VB6-compatible compiler whose pages
state VB6 compatibility member by member:

    git clone --depth 1 --filter=blob:none --sparse https://github.com/twinbasic/documentation.git
    (cd documentation && git sparse-checkout set docs/Reference/Default)
    node scripts/transcribe-vb6-docs.mjs <path to that checkout>

The pages are `docs/Reference/Default/VB/<Class>/index.md` in that
repository (MIT; the notice is in `THIRD_PARTY_NOTICES.md`); the current
model came from commit `5799ad8e236a77d2b9431f56ab3346fabcbefc12`
(2026-07-17). Filtered to VB6's real surface:

- twinBASIC's own controls (CheckMark, MultiFrame, QRCode, Report) are not
  transcribed;
- a member the page marks as a twinBASIC addition ("New in twinBASIC.",
  "twinBASIC-specific", "VB6 had no equivalent") is left out and listed under
  the dump's `excludedMembers`; a parenthetical "(new in twinBASIC)" beside
  one constant value marks that value, not the member;
- a member the page marks "Reserved for compatibility with VB6; not currently
  implemented in twinBASIC" is real VB6 that the oracle does not run: it is
  kept, flagged `reserved`, and its note travels into the model as the
  member's remarks;
- `Forms` is carried as a class of its own (Count, Item) from the Global
  page's section, so `Forms` chains to a collection rather than to one Form.

### Cross-read against Microsoft's archive

    node scripts/fetch-vb6-reference-names.mjs

fetches the table of contents of the archived Visual Basic 6.0 documentation
(`learn.microsoft.com/previous-versions/visualstudio/visual-basic-6/toc.json`)
and records every property, method, event, function, statement and object
name it documents in `reference/vb6/vb6-reference-names.json`. The
transcriber reads that file: a property or method whose name has no page
anywhere in the VB6 reference is not VB6, whatever the twinBASIC page says,
and is left out with the reason "no page in the VB6 Language Reference"
(Opacity, TransparencyKey, Anchors, Dock, VisualStyles, App.IsInIDE and the
rest, all in `docs/vb6_reference_exclusions.md`).

This is a name-level check, and it has limits the record states rather than
hides:

- the archive keeps the reference pages but not the "Applies To" object
  lists behind them, so which objects a documented name belongs to cannot be
  read from Microsoft; per-object membership rests on twinBASIC's pages;
- the archive is missing event pages VB6 certainly had (Unload, the OLE
  drag-and-drop events, a scroll bar's Scroll), so events are not cross-read;
- one real VB6 member, `DataMemberChanged`, has no page of its kind in the
  archive and is kept by name (`KNOWN_VB6_MISSING_FROM_ARCHIVE`).

### Cross-read against twinBASIC's own source

The IDE ships the VB compatibility package as a `.twinproj` of MIT-licensed
twinBASIC source (`packages\{F50B82D0-DCAB-43FE-9631-11959D4A4728}_VB\package.twinproj`).
Exported with the compiler's CLI, it is the oracle's exact statement of what
it implements, member by member:

    <ide>\bin\twinBASIC_win64.exe export "<ide>\packages\{F50B82D0-DCAB-43FE-9631-11959D4A4728}_VB\package.twinproj" "<folder>\" --overwrite
    node scripts/twinbasic-vb-surface.mjs <folder> --ide "twinBASIC v0.15.983 (BETA 983)"

`scripts/twinbasic-vb-surface.mjs` reads the exported `.twin` classes (bases
by `Inherits`, coclasses by their default interface, `[Unimplemented]` and
`[Hidden]` attributes, `#If FEATURE_...` guards recorded and treated as on)
into `reference/vb6/twinbasic-vb-surface.json`; its reader is pinned by
`tests/twinbasicVbSurface.test.ts`. The transcriber then marks every `VB`
member with `oracle`: `implemented`, `unimplemented` (declared for VB6
compatibility, not run) or `absent` (unknown to the package, so no oracle
verdict can vouch for it), and the generator carries the flag into the
model as `HostMember.oracle`. `docs/vb6_reference_surface.md` is the
record: per-class counts, the absent members, where the pages and the
package disagree on implementation, and the members the package has that
the pages do not (recorded, never added). This replaces the LSP probing the
roadmap first proposed: the source says exactly what a hover could only
sample.

## Generating the model

    node scripts/generate-host-object-model.mjs vb6

reads `reference/vb6/json` and writes `src/analyzer/host/vb6ObjectModelData.ts`:
VB types under `VB.`, VBRUN types under `VBRUN.`, constants and enumerations,
each type carrying its dump's `source` as provenance. Events are carried
(kind `event`) because a VB6 form's handler stubs come from them; they never
appear as object members. Nothing is exhaustive: the model offers and
describes, and never proves a member absent (`docs/roadmap_vb6_support.md`,
Slice 3).
