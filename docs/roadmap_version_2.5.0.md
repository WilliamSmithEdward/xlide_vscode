# XLIDE Roadmap Version 2.5.0

Version 2.5.0 builds on the v2.4.0 static-analysis completeness baseline. It has
two goals:

1. **Expression binder (MS-VBAL §5.6) completion** — the keystone capability that
   unlocks the diagnostics v2.4.0 deferred.
2. **Syntax-corpus completeness** — finish turning the corpus from a partly-mined
   archive into a fully-mined, promoted, provenance-tracked evidence system.

The same evidence-led, no-false-positive discipline as v2.4.0 applies: a hard
diagnostic ships only when it is provably correct (MS-VBAL, the Excel/VBE oracle,
or deterministic XLIDE-owned metadata with tests), and anything unknown,
ambiguous, or `Variant` stays quiet. Each shipped behavior needs a positive, a
negative, and a no-diagnostic control plus a named evidence source.

## Goal 1: Expression binder (MS-VBAL §5.6)

Slices 1–3 landed in v2.4.0 (the `ExprNode` AST, statement wiring into
`Assignment`/`Call` nodes, and `If`/`ElseIf`/`Else` branch modeling). v2.5.0
completes the binder and cashes it in for the deferred type families:

- [ ] Named / omitted arguments and bang (`!`) member access in statement
  structuring (currently kept raw).
- [ ] Flow-sensitive identifier binding and definite assignment — with care
  around loops, `GoTo`, and error handlers (the main false-positive risk).
- [ ] Binder-dependent diagnostics, each gated on no-FP evidence:
  - Comparisons (`=`/`<>`/`<`/`Like`/`Is`) numeric/string/Date/Object matrix.
  - Date coercion (string/numeric→Date, Date parameters; locale-sensitive cases
    stay quiet until the oracle proves them).
  - Broad array element typing and array-parameter / `ParamArray`-element
    compatibility.
  - Default members (`VB_UserMemId = 0` implicit `.Value`/`.Item` resolution at
    assignment and call sites; quiet when the default member is unknown).
  - Boolean operators (`And`/`Or`/`Not`, numeric→Boolean).
  - Non-scalar ByRef (object refs, arrays, `Variant`, named arguments).
  - String-concatenation operand typing; positional-after-named arguments
    (PCEC_008 — oracle-confirmed, binder-gated).

Definition of done: the binder is complete enough that each family above either
ships with the three controls + a named source, or is explicitly deferred with a
documented reason.

## Goal 2: Syntax-corpus completeness

The corpus is already provenance-tracked and test-enforced; the compile-error
vein is fully mined. v2.5.0 finishes the rest. The measurable burn-down — the
per-source-file mining checklist, the per-surface disposition (including the two
surfaces added this cycle, `error-handling-flow` and `xlide-directives`), and the
named remaining veins — lives in the **Completeness Checklist** section of
`syntax_corpus/managed_backlog.md`; the items below are its roadmap summary:

- [ ] Mine the remaining corpus veins with the Excel/VBE oracle —
  runtime-resolution, class/UserForm canary matrix, host-behavior — promoting
  no-FP candidates as `spec-derived` / `vbe-oracle-verified` and refuting/recording
  the rest.
- [ ] Oracle-map the deferred numeric/host boundaries: `Single`/`Double` and
  hex/octal width overflow, and the `&`-suffix overflow boundary (needs next-token
  concat-vs-suffix mapping).
- [ ] Close the realtime incomplete-expression corpus tail (UserForm/class
  partial-state recovery controls beyond the v2.4.0 slice).
- [ ] Promote or retire every remaining pending Markdown case; move fully-mined
  source files to `reference`/`Archive` status and keep `corpus_provenance.json`,
  `diagnostic_influence_audit.json`, and the manifest test-enforced.

Definition of done: the corpus is fully mined — every high-value case is promoted
or recorded as a tracked deferral, the realtime tail is closed, and no source
file is left in an un-dispositioned `mining` state.

## Carried forward (beyond the two pillars)

Tracked, but not the v2.5.0 focus — these are product/object-member features that
the binder may partly enable but that stand on their own:

- Object-model member binding: declared `Event` / `WithEvents` / `RaiseEvent`
  binding beyond module-kind validation, and `Implements` interface-member
  completeness.
- External `.vbref.xml` object/member metadata as a diagnostic source.

## Won't implement

Recorded in the MS-VBAL verification-map "Won't Implement" section — no oracle
path and near-zero payoff:

- Date-literal inner-grammar validation (locale-sensitive → false-positive-prone).
- Exact legacy-codepage non-Latin identifier ranges (old-VBE detail).
- UserForm/chart designer-backed (`.frm`/`.frx`) member parsing.

## Relationship to v2.4.0

v2.4.0 closed the provable static-analysis surface and recorded everything it
left quiet as a documented deferral (`docs/static_analysis_completeness_2.4.0.md`,
the roadmap "Deferred" section, and the type-coverage Readiness Classification).
v2.5.0 is where the keystone deferral — the expression binder — and the corpus
tail get built, so those quiet areas can become provable diagnostics.
