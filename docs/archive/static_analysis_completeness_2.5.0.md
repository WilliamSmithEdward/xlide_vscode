# Static-Analysis Completeness Report - Version 2.5.0

This is the release-gate artifact for XLIDE v2.5.0. It records what is complete,
what is intentionally quiet, and what moves to a later release - backed by named
evidence rather than a green test run alone. See `docs/roadmap_version_2.5.0.md`
for the two goals and their definitions of done, and
`docs/static_analysis_completeness_2.4.0.md` for the prior baseline.

## Bottom line

v2.5.0 is **closeable**. Both goals reach their definition of done:

1. **Expression binder (MS-VBAL §5.6) completion** - the binder is complete
   enough that every binder-dependent diagnostic family is dispositioned: shipped
   with positive + negative + no-diagnostic controls and a named source, or
   explicitly deferred with a documented reason.
2. **Syntax-corpus completeness** - no source file remains in `mining`; every
   high-value vein is shipped, refuted, already-covered, or recorded as a tracked
   deferral.

The project's #1 rule holds throughout - **no false positives**: a hard/red
diagnostic fires only when the construct is provably wrong (MS-VBAL, the Excel/VBE
oracle, or deterministic XLIDE metadata); anything unknown, ambiguous, `Variant`,
or runtime-convertible stays quiet. The one new red shipped this cycle survived an
8-angle adversarial false-positive hunt with zero findings.

## Evidence base

Counted from the repository (not estimated):

| Metric | Value | Source |
| --- | --- | --- |
| Diagnostic codes (audited) | **117** (was 112) | `syntax_corpus/diagnostic_influence_audit.json` |
| - VBE-oracle-verified | 71 | audit `status` |
| - spec-derived (MS-VBAL) | 46 | audit `status` |
| Diagnostic kind | 100 compile-error · 11 deterministic-runtime · 2 runtime-risk · 4 style-policy | `ruleMetadata.ts` / audit |
| Excel/VBE oracle cases | **397** (was 342), 100% `vbe-oracle-verified` | `syntax_corpus/oracle/vbe_oracle_cases.json` |
| - rejected probes / accepted controls | 241 / 156 | oracle `expected` |
| MS-VBAL verification-map rows | **195 Verified** (was 182); remaining Partial rows all deferred | `docs/spec/MS-VBAL.verification-map.md` |
| TypeScript tests | **2,071** (was 1,954) | `npm test` |
| Corpus files left in `mining` | **0** | `syntax_corpus/corpus_provenance.json` |

## What shipped this cycle

**Goal 1 - binder + binder-dependent reds:**

- **Expression AST (§5.6) completion** - named arguments (`name:=expr`), omitted
  arguments, and bang (`!`) member access are now structured (the `Argument` AST
  carries `name`/`nameSpan` and omitted slots; `MemberAccessExpr.accessKind`
  distinguishes `'dot'` from `'bang'`).
- **`argument-shape-mismatch`** (new compile-error red) - an array or same-module
  `Type` argument passed into a scalar parameter, or a scalar/`Variant` into an
  array parameter, is rejected. Decides on declared shape only, never element-type
  coercion; 9 promoted oracle cases (5 reject shapes, 4 accept controls); disjoint
  from `byref-argument-type-mismatch`; zero false positives under an 8-angle
  adversarial hunt.
- **`non-scalar-binary-operand`**, **`is-operator-non-object`**,
  **`string-arithmetic-coercion`**, **`typeof-is-always-false`** - the
  provable-shape and provable-scalar subsets of the comparison / Boolean /
  concatenation / `Is` / `TypeOf...Is` families.
- **Structural `If`/`ElseIf`/`Else` branch-merge** - the shared dataflow now
  intersects branch arms, extending `object-variable-not-set` (RTE 91) and
  `unallocated-dynamic-array-access` (RTE 9) to check accesses *inside* balanced
  `If` arms; FP-safe and conservative on any label / `GoTo` / `On Error` / loop.
- **Default-member metadata hardening** - `VB_UserMemId` matching parses the value
  numerically against a named DISPID constant.

**Goal 2 - corpus completeness:** all six remaining `mining` files dispositioned
and flipped to `reference`; the realtime-recovery tail closed; the provenance,
`diagnostic_influence_audit`, and manifest tests stay green.

## Definition-of-done attestation

**Goal 1** - each binder-dependent family is in a terminal state:

| Family | Disposition |
| --- | --- |
| Non-scalar operands (array/UDT in scalar operators) | **Shipped** - `non-scalar-binary-operand` |
| `Is` on a scalar operand | **Shipped** - `is-operator-non-object` |
| Argument shape (array/UDT ↔ scalar / array param; non-scalar ByRef) | **Shipped** - `argument-shape-mismatch` |
| Positional-after-named arguments | **Shipped** - `argument-count` (PCEC_008) |
| Non-numeric string in arithmetic (RTE 13) | **Shipped** - `string-arithmetic-coercion` |
| Comparison / Boolean / concat scalar-coercion matrix | **Deferred** - runtime-coerced, never a compile error (oracle controls accept) |
| Date coercion (broad) | **Deferred** - runtime/locale-dependent; the `CDate` sub-slice ships as `runtime-conversion-value` |
| Default members (`VB_UserMemId = 0`) | **Deferred** - runtime member lookup; metadata infra built, diagnostic defers |
| Array-element typing / `ParamArray`-element / object-`Variant` ByRef | **Deferred** - narrow follow-ons / `Variant`-accepting |
| Definite assignment (use-before-assignment) | **Deferred** - Flow phase 2 (needs loop/`GoTo`/handler joins) |

**Goal 2** - zero files in `mining`; every vein shipped / refuted / already-covered
/ deferred-with-reason; provenance + audit + manifest tests enforced.

## Deferrals (documented, carried)

Each is recorded in `docs/roadmap_version_2.5.0.md` and
`syntax_corpus/managed_backlog.md` with its reason:

- Scalar-coercion comparison/Boolean/concat matrix; Date coercion (broad);
  default-member-aware diagnostics; host-event wrong-signature binding (compiles +
  silently no-ops, so the oracle accepts it and no reject-red is promotable);
  `Null`/`Empty` `Variant` operand coercion.
- Numeric/host boundaries: `Single`/`Double`/`Decimal`, hex/octal width, and the
  `&`/`^`/`!`/`#`/`@` suffix overflow boundary (a proven false positive keeps
  `parseVbaIntegerLiteral` out of the literal-typing path).
- Flow phase 2 definite assignment; `WithEvents As Object` source-type
  restriction; continuation-count / line-length limits.

## Relationship to v2.4.0

v2.4.0 closed the provable static-analysis surface and deferred the expression
binder and the corpus tail. v2.5.0 builds the keystone binder and closes the
corpus - turning several of those quiet areas into provable diagnostics while
holding the no-false-positive line on the rest.
