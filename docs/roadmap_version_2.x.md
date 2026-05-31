# XLIDE Roadmap Version 2.x

Unified development roadmap for the current XLIDE workstream: realtime VBA
linting, language service hardening, type analysis, corpus coverage, and
Excel/VBE compatibility validation.

## North Star

XLIDE should make VS Code feel like a serious VBA IDE for real Excel projects:
fast while typing, deterministic in its diagnostics, honest about uncertainty,
and close to VBE behavior where VBE behavior is the relevant authority.

The goal is not to make VBA stricter than VBA. The goal is to catch real compile
errors early, offer useful guidance where appropriate, and avoid noisy or
heuristic diagnostics.

## Operating Principles

- Deterministic logic only. No heuristic analyzer behavior unless the operator
  explicitly approves it before implementation.
- Prefer no diagnostic over a guessed diagnostic.
- Red/error means a deterministic compile blocker or a construct XLIDE can prove
  invalid.
- Yellow/warning means XLIDE guidance, maintainability advice, or a soft risk
  that may still compile.
- Diagnostic language must match severity. Red/error diagnostics use
  authoritative wording such as "will fail", "will raise", or "is invalid" when
  the behavior is proven. Yellow/warning diagnostics use non-authoritative
  wording such as "may", "can", "risk", or "consider" because they represent
  guidance or uncertainty.
- For complicated work, create or update a roadmap before implementation.
- The Excel/VBE oracle is a discovery, debugging, and corpus-coverage tool, not
  a routine per-change test.
- The syntax corpus is evidence, not authority. Corpus cases may be incomplete
  or wrong and should be verified against the oracle when they drive analyzer
  behavior.
- Deterministic runtime-error rules may be red even when
  `vbeCompileEquivalent` is false, but only with focused runtime oracle evidence
  or an equally deterministic local proof.

## Current Baseline

### Landed

- Realtime diagnostics are split between structural linting and analyzer rules.
- Active diagnostics are catalogued in `src/analyzer/diagnostics/ruleMetadata.ts`.
- Rule metadata includes category, VBE compile-equivalence, and diagnostic-kind
  fields so compile-time and deterministic-runtime red squiggles stay separate.
- Same-module callable signatures are parsed for parameters and return types.
- High-confidence argument count and type mismatch diagnostics exist.
- Non-callable call statements are flagged, including bare variable statements
  that VBE Compile rejects.
- Runtime metadata is curated explicitly; parameter types are not inferred from
  parameter names.
- Inline documentation comments support descriptive metadata plus optional
  `type`, `unit`, and `value` hints.
- Excel/VBE oracle harness exists under `syntax_corpus/oracle/`.
- The wider unverified type-analysis backlog is tracked in
  `docs/type_analysis_corpus_coverage.md`.
- Oracle cleanup is coordinator-owned: the recorded disposable Excel PID is
  killed after each oracle case.

### Current Validation Layers

- Fast routine checks:
  - `npm.cmd test`
  - `npm.cmd run check-types`
  - `npm.cmd run compile`
- Optional oracle checks:
  - `npm.cmd run test:oracle:vbe`
  - Use only for VBE behavior discovery, corpus validation, or oracle harness
    changes.

## Workstream A: Diagnostic Policy and Metadata

Purpose: make every diagnostic self-describing and severity-safe.

- [x] Add `category` to every rule.
- [x] Add `vbeCompileEquivalent` to every rule.
- [x] Add `diagnosticKind` to every rule.
- [x] Document red/yellow/no-diagnostic severity policy.
- [x] Document diagnostic language policy: red is authoritative, yellow is
  advisory.
- [x] Add tests that assert every rule declares category and VBE equivalence.
- [x] Add a short diagnostic policy table to architecture docs.
- [ ] Use category/equivalence fields when reporting workbook lint summaries.

Definition of done:

- Every emitted diagnostic can be classified as VBE compile-equivalent,
  deterministic runtime error, runtime-risk, XLIDE guidance, or
  project/model-specific.
- New diagnostics cannot be added without category, VBE equivalence, and
  diagnostic-kind metadata.

## Workstream B: Syntax Corpus and Oracle Coverage

Purpose: keep corpus fixtures useful without letting stale assumptions become
truth.

- [x] Add Excel/VBE oracle harness.
- [x] Capture VBE compile popups as the primary oracle signal.
- [x] Keep oracle out of routine per-change verification.
- [x] Kill the recorded disposable Excel PID after each oracle case.
- [x] Mark corpus fixtures with provenance:
  - spec-derived
  - VBE-oracle-verified
  - observed but not asserted
  - pending verification
- [x] Add an oracle-result update workflow for promoted corpus cases.
- [x] Review current syntax corpus for cases that predate the oracle and mark
  them pending until verified.
- [x] Add focused oracle fixtures for newly debated syntax edges only.
- [x] Add a type-analysis corpus coverage matrix for pending, partial, missing,
  and verified type-analysis areas.
- [ ] Complete full `syntax_corpus` digestion into managed backlog categories
  (syntax, realtime recovery, type analysis, host behavior, completion context,
  UserForm/designer symbols, limits, and legacy edges). Treat every Markdown
  case as raw material until promoted through spec, oracle, or deterministic
  XLIDE-owned evidence.

Definition of done:

- Corpus-driven analyzer behavior has traceable provenance.
- If a corpus case and the oracle disagree, the case is treated as suspect until
  resolved.

## Workstream C: Phase 1 Type System Cleanup

Purpose: finish the conservative first slice before broadening inference.

- [x] Same-module procedure parameter and return metadata.
- [x] Literal and declared-variable argument type inference.
- [x] Named argument mapping before validation.
- [x] Runtime function participation from curated metadata.
- [x] High-confidence nonnumeric string literal mismatch diagnostics.
- [x] Add fixture coverage for every current compatibility edge:
  - [x] numeric literal to numeric parameter
  - [x] numeric string to numeric parameter
  - [x] nonnumeric string to numeric parameter
  - [x] Variant suppressing hard mismatch
  - [x] named argument order
  - [x] omitted required argument
  - [x] omitted optional argument behavior
- [x] Separate compile-equivalent argument errors from deterministic runtime
  errors in tests and metadata.
- [x] Split object argument mismatches into a compile-equivalent red diagnostic
  after focused oracle verification.
- [x] Add red deterministic-runtime-error diagnostic for nonnumeric string
  literals in numeric arithmetic expressions after focused oracle verification
  showed VBE Compile accepts the representative case but execution raises
  runtime error 13.
- [x] Promote deterministic nonnumeric string-to-number/Boolean assignment and
  argument coercion failures to red runtime-error diagnostics after focused
  runtime oracle verification.
- [x] Record the runtime type mismatch experiment matrix in
  `docs/vba_runtime_type_mismatch_oracle_matrix.md`.
- [x] Add compile-equivalent diagnostics for runtime functions used as `As` type
  names and `Set` assignments to known intrinsic scalar variables.

Definition of done:

- Phase 1 diagnostics have valid/invalid/unknown fixtures.
- Every hard error is deterministic and either VBE-equivalent or explicitly
  justified as XLIDE-invalid.

## Workstream D: Phase 2 Expression Return Types

Purpose: infer simple expression result types without guessing.

First vertical slice:

- [x] Same-module function call return type.
- [x] Nested same-module call return type.
- [x] Curated runtime conversion function return types:
  - [x] `CStr`
  - [x] `CDbl`
  - [x] `CCur`
  - [x] `CLng`
  - [x] `CBool`
- [x] Arithmetic result family inference for obvious numeric expressions.
- [x] String concatenation result inference for `&`.
- [x] Unknown or `Variant` expression operands suppress hard diagnostics.
- [x] Assignment type mismatch diagnostics after deterministic expression proof
  for scalar assignments.

Out of scope until proven:

- Locale-sensitive string coercion.
- Full operator overload/coercion lattice.
- Object member receiver chains.
- Flow-sensitive variable narrowing.

Definition of done:

- Expression inference only returns a concrete type when the source provides
  deterministic proof.
- Inference feeds argument validation and assignment validation without adding
  noisy false positives.

## Workstream E: Project-Wide Binder

Purpose: move from same-module checks to workbook-aware analysis.

- [x] Resolve unique exported standard-module `Sub`/`Function` signatures for
  cross-module argument count and type diagnostics.
- [ ] Resolve public procedures across standard modules for module-qualified
  calls and ambiguous duplicate behavior.
- [ ] Model module-level visibility and shadowing.
- [ ] Resolve `As` type names against project classes, UDTs, enums, and host
  object types before flagging broad unknown type names.
- [ ] Resolve enums and enum members across modules.
- [ ] Resolve UDT names across modules.
- [ ] Add workbook-level fixture builder for project analysis tests.
- [ ] Keep project diagnostics stable under module order changes.

Definition of done:

- Cross-module procedure calls are validated deterministically.
- Shadowing behavior is explicit and tested.

## Workstream F: Object and Member Types

Purpose: validate Excel/VBA object use where receiver type is known.

- [ ] Track `Set` assignments to known object types.
- [ ] Resolve class/document module member calls.
- [ ] Resolve curated Excel object model receiver chains.
- [ ] Add `member-not-found` only when receiver type is known.
- [ ] Add `set-required` and `set-forbidden` only where deterministic.

Definition of done:

- Object diagnostics do not guess from names.
- Host metadata has auditable provenance.

## Workstream G: Realtime Experience

Purpose: keep the live editor useful while the user is mid-keystroke.

- [ ] Suppress hard errors for incomplete expressions where VBE behavior is not
  yet deterministically knowable.
- [ ] Make diagnostic ranges precise and stable.
- [ ] Use metadata categories to tune Problems output and future filters.
- [ ] Keep signature help, hover, completion, and diagnostics sharing the same
  symbol/type model.

Definition of done:

- Typing incomplete code does not create avoidable noise.
- Completed invalid code produces deterministic, useful diagnostics.

## Immediate Next Steps

1. Digest the remaining pending Markdown corpus files into managed backlog
   categories, including
   `syntax_corpus/xlide_vba_realtime_linting_final_corpus_addendum.md`.
2. Use `docs/type_analysis_corpus_coverage.md` to choose the next verified
   corpus additions for the project-wide binder.
3. Start the project-wide binder vertical slice for public procedures across
   standard modules.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/type_analysis_corpus_coverage.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/xlide_vba_linting_test_strategy.md`
- `docs/xlide_development_principles.md`
- `src/analyzer/diagnostics/ruleMetadata.ts`
- `syntax_corpus/README.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
- `syntax_corpus/oracle/README.md`
- `syntax_corpus/oracle/run_excel_vbe_oracle.py`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
- `tests/corpusProvenance.test.ts`
