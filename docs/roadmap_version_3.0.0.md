# XLIDE Roadmap Version 3.0.0

Version 3.0.0 is the deferred product/backlog roadmap. It holds the work
intentionally excluded from the v2.4.0 static-analysis completeness sprint so that
the v2.4.0 analyzer surface could close with clean, provable edges. Nothing here
is a regression or an oversight — each item is a documented deferral carried
forward from `docs/roadmap_version_2.4.0.md` (see its "Deferred" and "Out Of
Scope" sections and `docs/static_analysis_completeness_2.4.0.md`).

The same evidence-led, no-false-positive discipline applies: a hard diagnostic
ships only when it is provably correct (MS-VBAL, Microsoft VBA/Office docs, the
Excel/VBE oracle, or deterministic XLIDE-owned metadata with tests), and anything
unknown, ambiguous, or `Variant` stays quiet.

## 1. Expression binder (MS-VBAL §5.6) and the binder-dependent type families

The §5.6 expression binder is the keystone that unblocks the deferred type
families. Slices 1–3 (expression AST, statement wiring, branch modeling) landed in
v2.4.0; the remaining work:

- Named / omitted arguments and bang (`!`) member access in statement structuring.
- Flow-sensitive identifier binding and definite assignment (FP-prone around
  loops, `GoTo`, and error handlers — needs care).
- Binder-dependent type rules, each gated on no-FP evidence: broad comparisons
  (`=`/`<`/`Like`/`Is` matrix), Date coercion, broad array element/parameter
  compatibility, default members (`VB_UserMemId = 0` implicit calls), Boolean
  operators, non-scalar ByRef, string-concatenation operand typing, and
  positional-after-named arguments (PCEC_008).

## 2. Oracle-gated numeric and host families

- `Single` / `Double` / `Decimal`, hex/octal width, and `&`/`^`/`!`/`#`/`@`-suffix
  numeric overflow — each needs VBE-oracle boundary mapping (the `&`-suffix case
  specifically needs next-token concat-vs-suffix mapping).
- Broader Excel host object surfaces and host-object RHS compatibility, promoted to
  exhaustive only after generated-dump provenance or oracle-backed controls.

## 3. Object-model and metadata binding (the original 3.0.0 backlog)

- Declared `Event` member binding beyond module-kind validation, `WithEvents`
  event-source type compatibility, and `RaiseEvent` signature/arity validation.
- `Implements` interface-member completeness.
- External `.vbref.xml` object/member metadata as a diagnostic source.

## Won't implement (recorded, not scheduled)

These have no oracle path and near-zero payoff; they are documented in the
MS-VBAL verification-map "Won't Implement" section and are not planned for any
release:

- Date-literal inner-grammar validation (locale-sensitive → false-positive-prone).
- Exact legacy-codepage non-Latin identifier ranges (old-VBE implementation
  detail).
- UserForm/chart designer-backed (`.frm`/`.frx`) member parsing.

## Relationship to v2.4.0

v2.4.0 closes the provable static-analysis surface. This roadmap is the agreed
home for everything that surface deliberately left quiet. When the expression
binder advances, items move from here into a concrete release with the usual
valid/invalid/unknown controls and a named evidence source.
