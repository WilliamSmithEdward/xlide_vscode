# XLIDE Roadmap Version 2.5.0

Version 2.5.0 builds on the v2.4.0 static-analysis completeness baseline. It has
two goals:

1. **Expression binder (MS-VBAL §5.6) completion** - the keystone capability that
   unlocks the diagnostics v2.4.0 deferred.
2. **Syntax-corpus completeness** - finish turning the corpus from a partly-mined
   archive into a fully-mined, promoted, provenance-tracked evidence system.

The same evidence-led, no-false-positive discipline as v2.4.0 applies: a hard
diagnostic ships only when it is provably correct (MS-VBAL, the Excel/VBE oracle,
or deterministic XLIDE-owned metadata with tests), and anything unknown,
ambiguous, or `Variant` stays quiet. Each shipped behavior needs a positive, a
negative, and a no-diagnostic control plus a named evidence source.

## Goal 1: Expression binder (MS-VBAL §5.6)

Slices 1-3 landed in v2.4.0 (the `ExprNode` AST, statement wiring into
`Assignment`/`Call` nodes, and `If`/`ElseIf`/`Else` branch modeling). v2.5.0
completes the binder and cashes it in for the deferred type families:

- [x] Named / omitted arguments and bang (`!`) member access in statement
  structuring - **shipped** in commit `b64b520` (the `Argument` AST carries
  `name` / `nameSpan` and omitted slots, and `MemberAccessExpr.accessKind`
  distinguishes `'dot'` from `'bang'`); produced by the §5.6 parser
  (`parseExpression.ts`) and consumed by the structuring rules `parseModule.ts`
  (parenless-call args via `parseParenlessArguments`) and `typeOfIs.ts` (walks
  `Argument.value` with an omitted-slot guard).
- [x] Flow-sensitive identifier binding and definite assignment - the branch-merge
  precision shipped (below); the *new* definite-assignment red is **moved to v2.6.0**
  (`docs/roadmap_version_2.6.0.md`, Goal 1), where the loop / `GoTo` / error-handler
  flow joins it needs are scoped.
  - ✅ Structural `If`/`ElseIf`/`Else` branch-merge - **shipped.** The shared
    straight-line dataflow now intersects branch arms (`walkBranchMergedBody`
    in `diagnostics/dataflow.ts`, gated by `procedureHasUnstructuredFlow` in
    `flow/procedureUnstructured.ts`), extending the two existing runtime-error
    flow reds - `object-variable-not-set` (RTE 91) and
    `unallocated-dynamic-array-access` (RTE 9) - to (a) check accesses *inside*
    balanced `If` arms from the correct entry state and (b) keep precise
    post-block state when every arm agrees. A tracked local advances to its
    good state only when it reaches that state on **every** arm **and** a
    syntactic `else` is present; otherwise it follows the existing conservative
    demotion. Any label, `GoTo`/`GoSub`, `On..GoTo/GoSub`, `On Error`, or
    `Resume` makes the whole procedure fall back to the straight-line walk, and
    `For`/`Do`/`While`/`With`/`Select` stay conservative (never entered).
    Controls: positive (Nothing / unallocated access inside a balanced arm
    flags), negative (set / allocated-before-`If` quiet; `On Error` / `GoTo` /
    loop → conservative `0`), no-diagnostic (balanced-arm allocation stays
    quiet); evidence is MS-VBAL §5.4.2 structured-`If` semantics plus the
    RTE 91/9 reds these rules already pin. No new diagnostic code; FP-safe by
    construction (fires only on provably-bad state).
  - **Moved to v2.6.0:** the *new* definite-assignment red (use-before-assignment)
    needs loop / `GoTo` / error-handler join modeling and oracle-pinned
    use-before-def cases to stay no-FP - scoped as Goal 1 of
    `docs/roadmap_version_2.6.0.md`.
- [x] Binder-dependent diagnostics - every family dispositioned, either shipped
  with the three controls + a named source, or deferred with a documented reason:
  - ✅ **Non-scalar operands** (a bare array or same-module user-defined `Type`)
    of scalar-requiring binary operators - **shipped** as `non-scalar-binary-operand`
    (oracle-verified across concatenation, arithmetic, comparison, and Boolean
    operator classes). `Is` on a scalar operand ships as `is-operator-non-object`;
    a non-numeric string literal in arithmetic (RTE 13) ships as
    `string-arithmetic-coercion`.
  - ✅ **Argument shape** (array / `Type` → scalar parameter, or scalar /
    `Variant` → array parameter - i.e. the provable non-scalar-ByRef cases) -
    **shipped** as `argument-shape-mismatch` (9 promoted oracle cases: 5 reject
    shapes, 4 accept controls; disjoint from `byref-argument-type-mismatch`).
  - ✅ Positional-after-named arguments (PCEC_008) - **shipped** under the
    `argument-count` rule via the token-level call extractor (oracle-verified
    `positional_after_named_argument_compile` plus paren-call and ParamArray
    forms, with legal positional-then-named and all-named orderings accepted).
  - ⊘ **Comparison / Boolean / string-concat scalar-coercion matrix** - *deferred
    with reason.* Scalar-vs-scalar mismatches in `=`/`<>`/`<`/`>`, `And`/`Or`/`Not`,
    and `&` coerce at runtime or are valid - never compile errors (oracle controls
    `string_concat_nonnumeric_string_runtime` and `numeric_plus_numeric_string_runtime`
    accept); no no-FP compile red is provable. The non-scalar-shape and Is-scalar
    halves shipped above.
  - ⊘ **Date coercion (broad family)** - *deferred with reason.* Implicit
    string→Date / numeric→Date assignment and `Date` parameters are runtime
    conversions, not compile errors (oracle `date_assignment_statement_compile`
    accepts `Dim d As Date: d = "1/2/2020"`); convert-vs-RTE-13 is value- and
    locale-dependent. The deterministic sub-slice (a plainly non-date string into
    `CDate`) ships as `runtime-conversion-value`.
  - ⊘ **Default members** (`VB_UserMemId = 0` implicit `.Value`/`.Item`) -
    *deferred with reason.* Implicit default-member resolution is a runtime member
    lookup (oracle `nonscalar_range_concat_control` proves `s = "x" & r`, with
    `Dim r As Range`, compiles because `Range` has a coercing default member);
    neither "has a coercing default member" nor "has none" is provable without a
    host default-member resolver carrying the post-coercion scalar type. The
    `VB_UserMemId` metadata infra is built and test-enforced; the *diagnostic*
    defers.
  - ⊘ **Array-element typing / `ParamArray`-element / object-or-`Variant` ByRef**
    - *deferred with reason.* `ParamArray` is `Variant` and accepts any element
    (oracle control accepts); broad element typing and both-arrays-different-
    element-type are tracked narrow follow-ons.

Definition of done: the binder is complete enough that each family above either
ships with the three controls + a named source, or is explicitly deferred with a
documented reason.

## Goal 2: Syntax-corpus completeness

The corpus is already provenance-tracked and test-enforced; the compile-error
vein is fully mined. v2.5.0 finishes the rest. The measurable burn-down - the
per-source-file mining checklist, the per-surface disposition (including the two
surfaces added this cycle, `error-handling-flow` and `xlide-directives`), and the
named remaining veins - lives in the **Completeness Checklist** section of
`syntax_corpus/managed_backlog.md`; the items below are its roadmap summary:

- [x] Mine the remaining corpus veins with the Excel/VBE oracle -
  runtime-resolution, class/UserForm canary matrix, host-behavior - promoting
  no-FP candidates and refuting/recording the rest. **Done:** every vein in the
  six previously-`mining` files is dispositioned (shipped / already-covered /
  control-or-refute-as-valid / defer-with-reason) per `managed_backlog.md`.
- [x] Oracle-map the deferred numeric/host boundaries - **dispositioned as
  deferred-with-reason:** `Single`/`Double` overflow is not a clean compile/RTE-6
  literal boundary, hex/octal width overflow needs the lexer to expose declared
  literal width plus an oracle map, and the `&`/`^`/`!`/`#`/`@` suffix boundary
  has a *proven* false positive (`s = 3000000000&"x"` compiles as concatenation -
  the lexer glues the token), so `parseVbaIntegerLiteral` is kept out of the
  literal-typing path.
- [x] Close the realtime incomplete-expression corpus tail - **done** (the
  pure-test realtime-recovery and casing veins are covered by deterministic unit
  tests; the addendum's residual veins are dispositioned).
- [x] Promote or retire every remaining pending Markdown case; move fully-mined
  source files to `reference` status and keep `corpus_provenance.json`,
  `diagnostic_influence_audit.json`, and the manifest test-enforced. **Done:** all
  six remaining `mining` files flipped to `reference`; zero files left in
  `mining`; the provenance, audit, and manifest tests stay green.

Definition of done: the corpus is fully mined - every high-value case is promoted
or recorded as a tracked deferral, the realtime tail is closed, and no source
file is left in an un-dispositioned `mining` state.

## Carried forward (beyond the two pillars)

Tracked, but not the v2.5.0 focus - these are product/object-member features that
the binder may partly enable but that stand on their own. **Now scheduled as Goal 4
of the v2.6.0 roadmap** (`docs/roadmap_version_2.6.0.md`):

- Object-model member binding: declared `Event` / `WithEvents` / `RaiseEvent`
  binding beyond module-kind validation, and `Implements` interface-member
  completeness.
- External `.vbref.xml` object/member metadata as a diagnostic source.

## Won't implement

Recorded in the MS-VBAL verification-map "Won't Implement" section - no oracle
path and near-zero payoff:

- Date-literal inner-grammar validation (locale-sensitive → false-positive-prone).
- Exact legacy-codepage non-Latin identifier ranges (old-VBE detail).
- UserForm/chart designer-backed (`.frm`/`.frx`) member parsing.

## Relationship to v2.4.0

v2.4.0 closed the provable static-analysis surface and recorded everything it
left quiet as a documented deferral (`docs/static_analysis_completeness_2.4.0.md`,
the roadmap "Deferred" section, and the type-coverage Readiness Classification).
v2.5.0 is where the keystone deferral - the expression binder - and the corpus
tail get built, so those quiet areas can become provable diagnostics.
