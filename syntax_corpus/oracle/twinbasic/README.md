# twinBASIC Oracle Harness

An optional, Windows-only developer audit tool that asks twinBASIC's compiler
whether a VBA/VB6 snippet builds. It is the oracle for the VB6 work
(`docs/roadmap_vb6_support.md`, Slice 4): XLIDE's production code never
depends on it, it is never part of `npm test`, and nothing from twinBASIC is
bundled.

twinBASIC is a superset of VB6 with published incompatibilities. Its verdict
is typed evidence about twinBASIC and inferred evidence about VB6; the harness
records it that way and never as "VB6-accepted".

## Run

```powershell
$env:XLIDE_TWINBASIC_DIR = 'C:\path\to\twinBASIC_IDE_BETA_983'
npm run test:oracle:twinbasic
```

or `node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --twinbasic <dir>`.
The folder is the IDE's: `twinBASIC.exe` and `bin\twinBASIC_win64.exe` must
be in it. Run one invocation at a time. Every run opens the IDE once per case
(the IDE is the only build driver twinBASIC ships), so expect windows to
appear and disappear on the desktop.

Useful forms:

```powershell
node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --case missing_trailing_required_argument
node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --limit 20 --json
node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --report syntax_corpus/oracle/twinbasic/parity_results.json --parity-doc docs/vb6_twinbasic_parity.md
node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --case some_case --keep --no-controls
node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs --limit 40 --append --report syntax_corpus/oracle/twinbasic/parity_results.json
```

`--append` keeps the report's earlier evidence and skips the cases it already
holds, so a long batch can be run in chunks; every chunk still runs its own
controls.

Speed and scale, all measured on BETA 983:

- a rejection is read 4 s after the empty `Build\` folder appears
  (`--stall`), an acceptance in about 3 s;
- `--parallel 3` runs three IDE instances at once, started 1.5 s apart
  (`--stagger`); each case has its own folder and process tree. Four started
  in the same instant left one stuck before it built, which is what the
  stagger and the startup ceiling (`--startup`, one retry) are for. The full
  VBE corpus (418 cases) took ten minutes this way, with no infrastructure
  failure;
- `--excel` also references the Excel and Office type libraries registered
  on the machine, read from the registry, because the VBE corpus was
  verified inside Excel and 27 of its cases name Excel objects. Off by
  default: language cases must not need Excel.

By default the cases are the VBE corpus (`../vbe_oracle_cases.json`), which
is how the parity report in `docs/vb6_twinbasic_parity.md` is made.

## How a case is asked

Measured on twinBASIC BETA 983; each step below was chosen because the
alternative was tried and failed.

1. The case is staged as a twinBASIC project folder: `Settings` from
   `Settings.template.json` (a Standard EXE that references the VB
   compatibility package; `project.optionExplicit` is off so a case's own
   `Option Explicit` decides, as in the VBE), one `.twin` file per module
   wrapping the VB6 source in `Module Name` / `Class Name`, and a startup
   module whose `Main` calls the entry point, takes `AddressOf` every
   procedure of every standard module, and instantiates every class module.
   The reach matters: twinBASIC's build only fails on code it links, and a
   syntax error in a procedure nothing reaches builds clean (measured). A
   `.bas` dropped into `Sources` is stored but never compiled, so modules are
   always `.twin`. A `.vbp` cannot be built headlessly either: the IDE
   imports it and then asks to save.
2. `bin\twinBASIC_win64.exe import "<case.twinproj>" "<folder>\" --overwrite`
   turns the folder into a project, headlessly.
3. `twinBASIC.exe "<case.twinproj>" --buildAndExit64` builds it. The IDE
   prints nothing outside its own panels and always exits 0; the verdict is
   read from what it does:
   - it exits and `Build\` holds the EXE: **accepted**;
   - it stays up past the stall window with `Build\` created and empty:
     **rejected**. This is the IDE's documented behaviour on compile errors
     (it stalls on "Please wait..."), and the empty `Build\` folder is the
     compiler having started the build and refused to write the output;
   - it exits without an EXE, or stays up with no `Build\` folder: an
     infrastructure outcome (`build_incomplete`, `timeout`), never evidence.
4. The IDE process tree is killed by PID, and only that tree.

## Controls

Every batch runs `control_accept` (a valid module) and `control_reject` (a
trailing comma omitting a required argument) before and after the cases. A
control that fails before the batch aborts the run as `ORACLE-FAIL`; a
control that fails after it demotes every `rejected` in the batch to
`unverified`, because a rejection is read from a stall and only the controls
prove the pipeline could have built. `--no-controls` exists for single-case
debugging; its results are observations.

## What it cannot say

- No error text or position: the compiler's diagnostics live in the IDE's
  panels, and reaching them means speaking the compiler's own websocket
  protocol. That is recorded as the open residual in the roadmap.
- Only the compile phase. Runtime cases from the VBE corpus are compared at
  compile time (they compiled before they ran), never executed.
- Members of a class module that a syntax error leaves unreachable, and
  `Property` procedures in standard modules (`AddressOf` cannot name them),
  are compiled only if the case's own code reaches them; the run records
  what it reached under `reach`.
- A case that declares `Sub Main` is skipped (`unsupported`): the harness
  needs the startup module's `Main`.
- The IDE's own text says twinBASIC Personal Edition refuses command-line
  builds; the machine this was measured on was allowed. A `timeout` at the
  startup stage with no `Build\` folder is the signature to check first.

## Provenance

Outcomes the harness writes are `accepted` and `rejected`, evidence phase
`compile`, and any corpus case promoted from a twinBASIC verdict must carry
`provenance: "twinbasic-oracle-verified"` (allowed in
`syntax_corpus/corpus_provenance.json`). None has been promoted yet: the
harness's first job is the parity report, which measures how far twinBASIC
agrees with the VBE on the 418 VBE-verified cases, and that number is the
fidelity every VB6 diagnostic and the `VB` model inherit.
