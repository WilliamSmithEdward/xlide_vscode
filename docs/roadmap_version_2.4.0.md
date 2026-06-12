# XLIDE Roadmap Version 2.4.0

Version 2.4.0 is the static-analysis completeness sprint. Its job is to close
the remaining evidence and coverage gaps around the VBA analyzer before the
broader product/backlog work resumes in Version 3.0.0.

This roadmap is deliberately evidence-led. A diagnostic, parser rule, binder
rule, completion inference, or corpus case is complete only when it has the
right backing source: MS-VBAL for core language grammar/semantics, Microsoft
VBA or Office documentation for documented runtime/host behavior, the Excel/VBE
oracle for compiler/runtime behavior, or deterministic XLIDE-owned metadata
with tests.

Version 2.4.0 also adopts `docs/development_strategy_guide.md` as an operating
guide. Static-analysis completeness work should look for opportunities to
centralize shared rules, simplify ownership, remove dead or duplicate paths,
split monolithic files when a clearer boundary exists, and keep each refactor
tied to a concrete behavior or verification improvement.

## Release Relationship

This roadmap is the thematic owner of static-analysis completeness; it is not a
single release. The structural consolidation in Priority 5 - shared per-workbook
index, module splits, and dead-code removal - shipped in the v2.3.0
audit-remediation release; its ongoing-discipline items, along with Priorities
1-4 and 6, remain open and may land across one or more subsequent releases. Keep
this file as the completeness source of truth even as individual priorities
ship, and move work out only through the Out Of Scope boundary below.

## What "Complete" Means Here

Completeness is evidence-led, not feature-count-led. The shared bar for every
shipped analyzer behavior is:

- A positive control (valid code stays quiet), a negative control (invalid code
  is flagged), and a no-diagnostic control (unknown, ambiguous, or incomplete
  code stays quiet).
- A named evidence source: MS-VBAL for core grammar/semantics, Microsoft
  VBA/Office documentation for documented runtime/host behavior, the Excel/VBE
  oracle for compiler/runtime verdicts, or deterministic XLIDE-owned metadata
  with tests.
- A documented suppression boundary for everything the rule cannot prove.

Static analysis is "complete" when Priority 6 can publish that audit, not when a
test run is green. Anything XLIDE cannot prove stays quiet and is recorded as an
intentional deferral.

## Critical Path: Finish The Expression Binder First

Most remaining completeness work depends on one capability gap. The parser
models declarations, procedures, parameters, and block balance fully, but
procedure-body statements are still captured as raw `Statement` text and only
member-access chains are parsed. `docs/spec/MS-VBAL.verification-map.md` records
the `Expression AST (calls/member/ops)` row (MS-VBAL section 5.6) as `Partial`.

Until the full expression AST lands, these stay deferred and must remain quiet:

- Flow-sensitive identifier binding and full shadowing.
- Arbitrary-expression argument and assignment typing beyond literals and single
  declared variables.
- `If` / `ElseIf` / `Else` branch modeling, currently flattened into one block.
- Comparisons, array element typing, bang (`!`) member access, `[A1]` evaluate
  shorthand, and `TypeOf ... Is`.

Land the section 5.6 expression layer before the Priority 4 type families that
depend on it. See Suggested Sequencing.

## Priority 1: Static Analysis Completeness

Purpose: finish the high-confidence analyzer surface without weakening the
project's no-false-positive rule.

Developer-experience impact:

- Moves workbook analysis closer to a complete first pass over real VBA
  projects.
- Makes live diagnostics and workbook reports feel consistent across syntax,
  declarations, binding, calls, assignments, object access, and runtime facts.
- Keeps hard red diagnostics reserved for behavior XLIDE can prove.
- Turns remaining "known incomplete" analyzer areas into explicit tracked
  closure items instead of ambient backlog.

Scope:

- [ ] Audit every active diagnostic rule for evidence source, false-positive
  policy, test coverage, and current gaps.
- [ ] Close remaining parser and binder gaps that block complete static
  analysis: expression binding, name-resolution edge cases, shadowing, call
  targets, assignment targets, and member receivers.
- [ ] Expand workbook fixture coverage for multi-module projects, module-kind
  behavior, duplicated names, ambiguous visible symbols, external-reference
  suppression, and same-workbook live overlays.
- [ ] Unify object/member rule policy across source-backed classes, UDTs, host
  metadata, runtime metadata, and future external metadata.
- [ ] Finish the deterministic runtime-analysis slice for constant folding,
  argument bounds, conversions, array allocation, object initialization, and
  runtime error diagnostics where oracle evidence proves behavior.
- [ ] Keep incomplete, late-bound, external-reference, `Variant`, ambiguous,
  and host-incomplete shapes quiet unless a rule has exhaustive provenance.
- [ ] Record every intentionally deferred false-positive risk in the relevant
  verification map, corpus matrix, or roadmap note.

Definition of done:

- Every shipped analyzer diagnostic has positive, negative, and
  no-diagnostic controls, a named evidence source, and a documented suppression
  boundary for unknown or ambiguous cases.

## Priority 2: MS-VBAL Completeness

Purpose: finish the MS-VBAL-backed language surface so lexer, parser, keyword,
declaration, statement, expression, and diagnostic behavior are traceable to
the specification wherever the spec governs the answer.

Developer-experience impact:

- Makes XLIDE's static analysis defensible as a VBA language implementation,
  not a collection of ad hoc editor checks.
- Gives future analyzer changes a stable spec map and fixture standard.
- Reduces churn from unclear grammar assumptions, especially around expression
  syntax and declaration forms.

Scope:

- [ ] Audit `docs/spec/MS-VBAL.verification-map.md` for every `Partial` and
  implicit gap in lexer, keyword, parser, symbol, expression, and diagnostic
  rows.
- [ ] Promote remaining MS-VBAL-backed syntax and declaration behavior to
  `Verified` when implementation, fixtures, and section references line up.
- [ ] Finish or explicitly defer exact legacy-codepage identifier behavior,
  date literal grammar, implementation-reserved names, expression grammar, and
  conditional-compilation edge cases.
- [ ] Add missing section citations to analyzer rules that currently rely on
  broad or inherited spec references.
- [ ] Separate core-language MS-VBAL facts from host, runtime-library, and
  Excel/VBE oracle facts in the verification map.
- [ ] Ensure new parser/diagnostic work updates the verification map in the
  same change as implementation and tests.

Progress:

- [x] Closed the `reserved-for-implementation-use` verification gap: the shared
  keyword table now treats `Attribute`, `VB_Name`, and related names as
  reserved for declaration validation while keeping them out of keyword casing
  so exported Attribute-line metadata remains raw and round-trippable.
- [x] Added focused non-Latin identifier lexer fixtures for the current
  Unicode-letter approximation and kept exact legacy-codepage support explicitly
  deferred in the MS-VBAL verification map.

Definition of done:

- The verification map gives a current, useful answer for each implemented
  MS-VBAL-governed feature: verified, intentionally partial, or deferred with a
  clear reason.

## Priority 3: Syntax Corpus Closure

Purpose: finish turning the syntax corpus from a broad discovery archive into a
managed, provenance-tracked evidence system for realtime and settled analyzer
behavior.

Developer-experience impact:

- Prevents stale Markdown examples from accidentally becoming diagnostic truth.
- Gives realtime editing regressions concrete fixtures with expected ranges and
  no-diagnostic controls.
- Makes corpus work measurable: pending cases get promoted, revised, or retired.

Scope:

- [ ] Audit `syntax_corpus/managed_backlog.md` and
  `syntax_corpus/corpus_provenance.json` for pending syntax, realtime recovery,
  limits, runtime-resolution, project-binding, and legacy-edge cases.
- [ ] Promote high-value Markdown cases into executable fixtures with
  `spec-derived` or `vbe-oracle-verified` provenance before they drive hard
  diagnostics.
- [ ] Retire, rewrite, or mark observational cases that conflict with MS-VBAL,
  Microsoft documentation, or asserted VBE oracle results.
- [ ] Add range-sensitive expectations for diagnostics whose usefulness depends
  on precise underline placement.
- [ ] Close the realtime incomplete-expression corpus around active edit spans,
  partial statements, partial calls, incomplete strings, and settled-state
  controls.
- [ ] Keep `diagnostic_influence_audit.json` synchronized with active rule
  evidence.

Definition of done:

- The syntax corpus can be used as a reliable backlog and regression source
  because each diagnostic-driving case has machine-readable provenance and a
  clear expected behavior.

## Priority 4: Type Corpus Closure

Purpose: finish the type-analysis coverage matrix enough that new binder and
diagnostic work can be selected from verified gaps instead of broad guesses.

Developer-experience impact:

- Makes type-analysis progress visible across scalars, objects, arrays, enums,
  UDTs, classes, document modules, host objects, runtime metadata, and workbook
  projects.
- Reduces false positives by requiring explicit valid, invalid, and unknown
  cases before new type rules ship.
- Gives the next major object/member and metadata work a clean foundation.

Scope:

- [ ] Audit `docs/type_analysis_corpus_coverage.md` and convert `Missing`,
  `Pending`, and high-value `Partial` rows into actionable fixture/oracle tasks.
- [ ] Finish matrices for ByRef compatibility, Date coercion, object
  assignment, fixed-length strings, operators/comparisons, arrays, enums, UDTs,
  classes/document modules, host receiver chains, runtime metadata, and
  cross-module binding.
- [ ] Require each new type-rule family to include one valid case, one invalid
  case, one unknown/no-diagnostic case, and an explicit verification path.
- [ ] Expand workbook project fixtures for shadowing, duplicate non-type
  symbols, external-reference shapes, overlay scenarios, and project-scale
  type/member binding.
- [ ] Keep generated host/reference coverage aligned with
  `docs/excel_reference_coverage.md`, including exhaustive-vs-completion-only
  policy for hard member diagnostics.
- [ ] Record which type corpus rows are ready for implementation, which need
  oracle work, and which must wait for broader expression binding.

Progress:

- [x] Retired the stale `ParamArray typing` Pending row in
  `docs/type_analysis_corpus_coverage.md`: declaration syntax is now recorded
  as verified for the active `paramarray-not-last`,
  `paramarray-with-optional`, and `paramarray-non-variant` diagnostic slice,
  while call element inference remains explicitly partial.

Type-family closure status (live detail in
`docs/type_analysis_corpus_coverage.md`):

| Type family | Status | Remaining for closure |
| --- | --- | --- |
| ByRef compatibility | Partial (scalar exactness) | object references, arrays, Variant behavior, named arguments, runtime mutation |
| Date coercion | Missing | string/numeric-to-Date, Date parameters; locale-sensitive strings stay no-diagnostic |
| Comparisons | Missing | numeric/string/Date/Object, `Like`, `Is` operator matrix |
| Arrays | Pending | dynamic/fixed, element type, array-parameter compatibility, `ParamArray` elements |
| Numeric family / overflow | Partial (Byte/Integer decimal) | Long/LongLong/LongPtr, Single, Decimal, Currency, suffixes, hex/octal bounds |
| Boolean | Partial | parameters, numeric-to-Boolean, `And`/`Or`/`Not` operands |
| Optional / default parameters | Partial (scalar defaults) | Date defaults, object/array invalid defaults, missing and named optional slots |
| Fixed-length strings | Partial (declaration metadata) | assignment/truncation, nonliteral length expressions |
| Enums | Partial | enum-to-integer compatibility, unknown and ambiguous enum names |
| UDTs | Partial | nested fields, arrays in UDTs, object members in UDTs |
| Return assignment | Partial | Property Let/Set consistency, object returns through default members |

Each row ships only with the valid, invalid, and unknown/no-diagnostic controls
and the verification path required below.

Definition of done:

- Type-analysis coverage is complete enough to drive future binder slices from
  a named matrix, and every implemented type rule has valid, invalid, and
  unknown/no-diagnostic controls.

## Priority 5: Architecture Hygiene and Simplification

Purpose: use the static-analysis completeness sprint to make the analyzer and
language-service codebase easier to understand, test, and extend.

Developer-experience impact:

- Makes future diagnostics, binder work, corpus promotion, and metadata
  integration cheaper because shared rules live in obvious places.
- Reduces drift between live diagnostics, workbook analysis, completion, hover,
  signature help, navigation, quick fixes, and test/oracle workflows.
- Improves maintainer confidence by retiring duplicated pipelines and dead code
  as rule coverage becomes more complete.

Scope:

- [ ] Treat `docs/development_strategy_guide.md` as the working checklist for
  meaningful 2.4.0 analyzer and language-service changes.
- [ ] During each static-analysis slice, identify the broader rule behind the
  case and route all relevant surfaces through the same helper or owner.
- [x] Look for duplicated validation, parsing, binding, metadata-precedence,
  diagnostic, completion, navigation, formatting, or fixture-building paths and
  consolidate them when the change is connected to the active behavior.
  (Completed by the v2.3.0 audit-remediation release: shared per-workbook
  project index, workbook module operations, sidecar codec, test-run
  pipeline, identity/text helpers, and consolidated webview scaffolding.)
- [x] Split large analyzer or provider modules into focused files when a
  stable ownership boundary emerges, especially for expression binding,
  diagnostic families, metadata resolution, corpus/oracle plumbing, or
  provider adapters.
  (Completed by the v2.3.0 audit-remediation release: diagnostics rule
  registry with per-family rule modules, per-domain command modules,
  provider subsystem modules, and externalized webview/test-host assets.)
- [x] Remove dead code, stale helpers, redundant fallbacks, and secondary
  pipelines once the shared rule and regression tests cover the old behavior.
  (Completed by the v2.3.0 audit-remediation release's dead-code sweep across
  the analyzer barrel, symbol index, settings mutators, and test plumbing.)
- [ ] Keep confidence levels explicit: known facts can drive hard diagnostics,
  inferred facts can guide completion or non-red help, and unknown facts must
  stay quiet.
- [ ] Add tests for the general rule, not only the reported example, and verify
  affected surfaces that consume the same shared helper.
- [ ] Avoid broad cleanup that is unrelated to the active 2.4.0 completeness
  work; refactors should leave a clearer rule, simpler ownership, or fewer
  valid implementation paths.

Definition of done:

- Major 2.4.0 analyzer changes leave behind fewer duplicated paths, clearer
  module ownership, and tests that prove shared-rule behavior across affected
  surfaces.

## Priority 6: Completeness Reporting and Release Gate

Purpose: make the 2.4.0 completeness claim auditable before closing the release.

Developer-experience impact:

- Gives maintainers a single checklist for deciding whether static analysis is
  ready for the next product/backlog phase.
- Keeps evidence artifacts, tests, and roadmap claims aligned.

Scope:

- [ ] Add or update a short static-analysis completeness summary that links the
  MS-VBAL map, syntax corpus, type corpus, diagnostic influence audit, workbook
  fixtures, and remaining intentional deferrals.
- [ ] Run the TypeScript analyzer tests and any targeted oracle checks needed
  to verify newly promoted cases.
- [ ] Ensure README, architecture, roadmap, and corpus docs point to Version
  2.4.0 as the active completeness sprint and Version 3.0.0 as the deferred
  product/backlog roadmap.
- [ ] Close 2.4.0 only with a clear list of what is complete, what is
  intentionally quiet, and what moved to Version 3.0.0 or later.

Definition of done:

- The release has an auditable static-analysis completeness record, not just a
  set of passing tests.

## Out Of Scope: Deferred To Version 3.0.0

These are intentionally excluded from static-analysis completeness so the
analyzer surface can close with clean edges. They are tracked in
`docs/roadmap_version_3.0.0.md`:

- UserForm designer-backed members and form/control event-handler authoring.
- Declared `Event` member binding beyond module-kind validation, `WithEvents`
  event-source type compatibility, and `RaiseEvent` signature/arity validation.
- `Implements` interface-member completeness.
- External `.vbref.xml` object/member metadata as a diagnostic source.
- Default-member expression semantics such as `textValue = p`.

Keeping these out is what lets Priorities 1-4 reach a provable, auditable close.

## Suggested Sequencing

1. Expression binder (MS-VBAL section 5.6): the critical-path keystone that
   unblocks flow-sensitive binding, branch modeling, and arbitrary-expression
   typing.
2. Type families (Priority 4): ByRef, Date, comparisons, and arrays first; most
   fall out of step 1 once expression types are available.
3. Deterministic runtime slice (Priority 1): division-by-zero reachability and
   `On Error Resume Next` policy, broader runtime-argument bounds, `IIf` eager
   branch faults, and `Null`/`Empty` coercion controls, now feasible on the
   binder.
4. MS-VBAL Partial decisions (Priority 2): finish or explicitly defer the
   remaining `Partial`/`Pending` rows - date-literal inner grammar, non-Latin
   codepage ranges, malformed directive blocks, and pointer-sized API checks -
   each with a documented reason.
5. Corpus promotion, per-rule audit, and the completeness report (Priorities 3
   and 6): the release gate that lets this roadmap close and hands the remainder
   to Version 3.0.0.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/roadmap_version_2.1.0.md`
- `docs/roadmap_version_2.4.0.md`
- `docs/roadmap_version_3.0.0.md`
- `docs/development_strategy_guide.md`
- `docs/spec/MS-VBAL.verification-map.md`
- `docs/spec/MS-VBAL.version.md`
- `docs/xlide_vba_language_service_roadmap.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/type_analysis_corpus_coverage.md`
- `docs/xlide_vba_analysis_test_strategy.md`
- `docs/excel_reference_coverage.md`
- `syntax_corpus/README.md`
- `syntax_corpus/managed_backlog.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
- `tests/fixtures/vbaProjects/`
- `tests/vbaDiagnostics.test.ts`
- `tests/vbaParser.test.ts`
- `tests/vbaSymbolGraph.test.ts`
- `tests/vbaRuntime.test.ts`
