# Static-Analysis Completeness Report - Version 2.4.0

This is the Priority 6 release-gate artifact for the v2.4.0 static-analysis
completeness sprint. It records what is complete, what is intentionally quiet,
and what moves to a later release - backed by named evidence rather than a green
test run. See `docs/roadmap_version_2.4.0.md` for the priority definitions.

## Bottom line

v2.4.0 is **closeable**. Every diagnostic the analyzer emits is backed by a named
evidence source and gated by the project's #1 rule - **no false positives**: a
hard/red diagnostic fires only when the construct is provably wrong; anything
unknown, ambiguous, `Variant`, or runtime-convertible stays quiet. Everything that
cannot meet that bar today is **explicitly deferred with a documented reason**,
which the roadmap's definition of done treats as a valid form of closure.

## Evidence base

Counted from the repository (not estimated):

| Metric | Value | Source |
| --- | --- | --- |
| Diagnostic codes (audited) | **112** | `syntax_corpus/diagnostic_influence_audit.json` |
| - VBE-oracle-verified | 66 | audit `status` |
| - spec-derived (MS-VBAL) | 46 | audit `status` |
| Diagnostic kind | 96 compile-error · 10 deterministic-runtime · 2 runtime-risk · 4 style | `ruleMetadata.ts` |
| Excel/VBE oracle cases | **342**, 100% `vbe-oracle-verified` | `syntax_corpus/oracle/vbe_oracle_cases.json` |
| - rejected probes / accepted controls | 209 / 133 | oracle `expected` |
| MS-VBAL verification-map rows | **182 Verified**; remaining Partial rows all deferred | `docs/spec/MS-VBAL.verification-map.md` |
| Per-rule audit (controls) | 94/112 met the full bar on first pass; **0 missing positive/negative controls** | rule-evidence audit |

Evidence artifacts, each kept in sync with the live analyzer:

- **MS-VBAL grammar/semantics:** `docs/spec/MS-VBAL.verification-map.md` (per-code
  rows; a `corpusProvenance` test asserts the audited-code set equals the live
  registry + structural code set and that `diagnosticKind` matches per rule).
- **Provenance audit:** `syntax_corpus/diagnostic_influence_audit.json` (every
  code → evidence source, asserted oracle cases, documented suppression boundary).
- **Compiler/runtime ground truth:** `syntax_corpus/oracle/` (the Excel/VBE oracle
  harness and its 342 verified cases).
- **Type coverage + readiness:** `docs/type_analysis_corpus_coverage.md` (the
  ready / needs-oracle / awaits-binder classification).
- **Corpus management:** `syntax_corpus/managed_backlog.md`,
  `corpus_provenance.json`.
- **Workbook fixtures:** `tests/fixtures/vbaProjects/` (multi-module, module-kind,
  duplicate-name, ambiguous-symbol, external-reference, live-overlay canaries).

## Priority status

- **Priority 1 - Static-analysis completeness:** the per-rule evidence audit is
  complete (10-auditor pass over all 112 codes; punch-list cleared - one real
  false positive fixed, the two structural codes brought under provenance, the
  no-diagnostic boundary controls backfilled). Remaining bullets are binder/flow
  work, deferred with reasons (see Deferrals). **Closeable.**
- **Priority 2 - MS-VBAL completeness:** 182 verified rows; reserved-name and
  non-Latin-identifier gaps closed; all 112 codes carry a `specReference`;
  core-language vs host/runtime facts are separated. The named "remaining Partial
  decisions" (`PtrSafe`/`Long`-where-`LongPtr`, malformed directive blocks) are
  **decided**: the deterministic slices ship and are Verified; the heuristic
  remainders are deferred as FP-risky. **Closeable.**
- **Priority 3 - Syntax corpus closure:** the corpus is a provenance-tracked,
  test-enforced system; the PCEC and batch-2/3 oracle veins are fully reconciled;
  the influence audit is synced. Range-sensitive span controls and the
  realtime-recovery controls are added (Priority 3 closure tests); residual
  binder-dependent corpus cases are tracked deferrals. **Closeable.**
- **Priority 4 - Type corpus closure:** every shipped type-rule family has
  valid + invalid + unknown controls and a named source; the coverage matrix is
  current (Arrays/UDTs rows reconciled) and the **readiness classification** is
  recorded. The binder-dependent frontier is deferred. **Closeable.**
- **Priority 5 - Architecture hygiene:** shipped in v2.3.0 (shared per-workbook
  index, module splits, dead-code removal). **Done.**
- **Priority 6 - Completeness reporting & release gate:** this document. **Done.**

## No-false-positive discipline

What makes the deferrals legitimate is that shipping FP-prone rules is forbidden,
so deferral is the *correct* outcome, not a shortcut:

- 96 of 112 codes are VBE-compile-equivalent (provably wrong at compile time);
  the rest are deterministic runtime errors VBE raises, or downgradeable
  style/risk advisories.
- New rules are oracle-probed first and adversarially FP-hunted before shipping.
  This sprint's hunts caught and prevented real false positives (e.g. the `&`
  type-suffix overflow - `3000000000&"x"` is VBE-accepted as concatenation - was
  dropped; the `: Rem` structural leak was fixed).
- A `corpusProvenance` test makes the evidence chain unbreakable: no code can
  ship without metadata, an audit entry, and (for oracle-verified codes) at least
  one asserted oracle case.

## Intentional deferral inventory

These are deliberate, documented deferrals - the empty roadmap checkboxes reflect
sequencing, not omissions. Full reasons live in the roadmap "Deferred" and "Out
Of Scope" sections, the verification-map "Won't Implement" section, and the type
coverage readiness classification.

**Awaits the expression binder (MS-VBAL §5.6, the named critical-path keystone):**
broad comparisons (`=`/`<`/`Like`/`Is` matrix), Date parameters, broad array
element/parameter compatibility, default members (`VB_UserMemId = 0` implicit
calls), Boolean operators, non-scalar ByRef, flow-sensitive binding / full
shadowing, string-concatenation operand typing, positional-after-named arguments
(PCEC_008), and named/omitted-argument statement structuring. Each is FP-prone
without the binder and stays quiet.

**Needs VBE-oracle boundary mapping:** `Single`/`Double`/`Decimal`, hex/octal
width, and `&`/`^`/`!`/`#`/`@`-suffix numeric overflow (the `&`-suffix case needs
next-token concat-vs-suffix mapping); broader host object surfaces and host-object
RHS compatibility; Date coercion. (`Byte`/`Integer`/`Long`/`Currency`/`%`-suffix
overflow already ship.)

**Deferred as FP-risky heuristics (deterministic slices already ship):**
`Long`-where-`LongPtr` pointer-sizing (deterministic `declare-missing-ptrsafe`
ships, Win64-gated); broader malformed conditional-directive blocks
(`else-branch-order` + structural balance ship); `IIf` eager-branch faults
(literal-fatal branches already caught); `On Error Resume Next` reachability
gating (would only quiet existing diagnostics); broader `Null`/`Empty` coercion
(the deterministic `Null`-to-scalar case ships).

**Won't implement (no oracle path, near-zero payoff):** date-literal inner-grammar
validation (locale-sensitive → FP-prone); exact legacy-codepage non-Latin
identifier ranges (old-VBE detail). Recorded in the verification-map Won't
Implement section.

**Out of scope → Version 2.5.0:** `Event`/`WithEvents`/`RaiseEvent` member
binding beyond module-kind validation; `Implements` interface-member completeness;
external `.vbref.xml` metadata as a diagnostic source.

## Release-gate checklist

- [x] Every shipped diagnostic has positive, negative, and no-diagnostic controls
  (per-rule audit; backfilled gaps).
- [x] Every diagnostic has a named evidence source and a documented suppression
  boundary (provenance audit + verification map; structural codes now included).
- [x] Evidence artifacts are synchronized and test-enforced (`corpusProvenance`,
  `diagnosticRegistryCoverage`).
- [x] Every intentional deferral is recorded with a reason (this report + roadmap
  Deferred / Out-Of-Scope + verification-map Won't Implement + coverage readiness).
- [x] The full TypeScript test suite passes.

On these grounds, v2.4.0's static-analysis surface can close, handing the
binder-dependent and out-of-scope work to Version 2.5.0.
