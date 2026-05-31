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
- For complicated work, create or update a roadmap before implementation.
- The Excel/VBE oracle is a discovery, debugging, and corpus-coverage tool, not
  a routine per-change test.
- The syntax corpus is evidence, not authority. Corpus cases may be incomplete
  or wrong and should be verified against the oracle when they drive analyzer
  behavior.

## Current Baseline

### Landed

- Realtime diagnostics are split between structural linting and analyzer rules.
- Active diagnostics are catalogued in `src/analyzer/diagnostics/ruleMetadata.ts`.
- Rule metadata includes category and VBE compile-equivalence fields.
- Same-module callable signatures are parsed for parameters and return types.
- High-confidence argument count and type mismatch diagnostics exist.
- Non-callable call statements are flagged, including bare variable statements
  that VBE Compile rejects.
- Runtime metadata is curated explicitly; parameter types are not inferred from
  parameter names.
- Inline documentation comments support descriptive metadata plus optional
  `type`, `unit`, and `value` hints.
- Excel/VBE oracle harness exists under `syntax_corpus/oracle/`.
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
- [x] Document red/yellow/no-diagnostic severity policy.
- [x] Add tests that assert every rule declares category and VBE equivalence.
- [x] Add a short diagnostic policy table to architecture docs.
- [ ] Use category/equivalence fields when reporting workbook lint summaries.

Definition of done:

- Every emitted diagnostic can be classified as VBE compile-equivalent,
  runtime-risk, XLIDE guidance, or project/model-specific.
- New diagnostics cannot be added without category and VBE equivalence metadata.

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
- [ ] Add an oracle-result update workflow for promoted corpus cases.
- [ ] Review current syntax corpus for cases that predate the oracle and mark
  them pending until verified.
- [ ] Add focused oracle fixtures for newly debated syntax edges only.

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
- [ ] Add fixture coverage for every current compatibility edge:
  - numeric literal to numeric parameter
  - numeric string to numeric parameter
  - nonnumeric string to numeric parameter
  - Variant suppressing hard mismatch
  - named argument order
  - omitted required argument
  - omitted optional argument behavior
- [ ] Separate compile-equivalent argument errors from runtime-risk type
  warnings in tests and metadata.

Definition of done:

- Phase 1 diagnostics have valid/invalid/unknown fixtures.
- Every hard error is deterministic and either VBE-equivalent or explicitly
  justified as XLIDE-invalid.

## Workstream D: Phase 2 Expression Return Types

Purpose: infer simple expression result types without guessing.

First vertical slice:

- [ ] Same-module function call return type.
- [ ] Nested same-module call return type.
- [ ] Curated runtime conversion function return types:
  - `CStr`
  - `CDbl`
  - `CCur`
  - `CLng`
  - `CBool`
- [ ] Arithmetic result family inference for obvious numeric expressions.
- [ ] String concatenation result inference for `&`.
- [ ] Unknown or `Variant` expression operands suppress hard diagnostics.

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

- [ ] Resolve public procedures across standard modules.
- [ ] Model module-level visibility and shadowing.
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

1. Add an oracle-result update workflow for promoted corpus cases.
2. Audit existing corpus cases that influenced current diagnostics and mark
   whether they are oracle-verified.
3. Start Phase 2 with same-module return type inference for nested calls.
4. Add assignment type mismatch as warning/error only after the expression
   return type model can prove the source and target types.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/xlide_vba_linting_test_strategy.md`
- `docs/xlide_development_principles.md`
- `src/analyzer/diagnostics/ruleMetadata.ts`
- `syntax_corpus/README.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/oracle/README.md`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
- `tests/corpusProvenance.test.ts`
