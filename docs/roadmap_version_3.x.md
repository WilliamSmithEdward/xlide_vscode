# XLIDE Roadmap Version 3.x

Forward backlog for work intentionally moved out of Version 2.x closeout.
Version 2.x is closed around its launch-hardening scope; Version 3.x is for
deeper language-modeling, object-metadata, performance hardening, and
cross-workbook workflows that should not block v2 release readiness.

## North Star

V3 should extend XLIDE's deterministic VBA understanding without weakening the
core rule: no hard diagnostics from guesses. Every new binder, metadata, object,
or workflow surface must be backed by source facts, generated metadata with
provenance, the Microsoft specification, or focused VBE/Excel oracle evidence.

## Related Open Roadmaps and Evidence

V3 workstreams should point to the more specific sub-roadmaps and evidence files
that already own detailed backlog, provenance, and verification decisions. Use
this roadmap for product-level scope, and use the files below as the active
detail layer.

| Area | Source | Use in v3 |
| --- | --- | --- |
| Syntax corpus and pending edge cases | `syntax_corpus/managed_backlog.md`, `syntax_corpus/README.md` | Promote cases only through provenance, spec, oracle, or deterministic XLIDE-owned evidence; do not treat pending Markdown cases as authority. |
| VBE/Excel oracle workflow | `syntax_corpus/oracle/README.md`, `syntax_corpus/oracle/vbe_oracle_cases.json` | Add focused canaries and verification cases for disputed VBA behavior before adding hard diagnostics. |
| MS-VBAL alignment | `docs/spec/MS-VBAL.version.md`, `docs/spec/MS-VBAL.verification-map.md`, `docs/xlide_vba_language_service_roadmap.md` | Keep grammar, parser recovery, and diagnostic behavior tied to the recorded spec version and verification map. |
| Type-analysis coverage | `docs/type_analysis_corpus_coverage.md`, `docs/xlide_vba_type_system_roadmap.md` | Choose binder, type, and object slices from pending or partial coverage instead of broad guesses. |
| Diagnostic evidence and test strategy | `docs/xlide_vba_analysis_test_strategy.md`, `syntax_corpus/diagnostic_influence_audit.json`, `syntax_corpus/corpus_provenance.json` | Preserve rule provenance and evidence before adding or changing hard diagnostics. |
| External metadata | `docs/xlide_external_member_metadata.md` | Keep external object/member schema, authoring, reload, validation, and troubleshooting details in one source of truth. |
| Host/reference metadata | `docs/excel_reference_coverage.md`, `scripts/generate-excel-reference-metadata.mjs`, `scripts/generate-office-reference-metadata.mjs` | Track generated metadata provenance and coverage before promoting host surfaces into hard diagnostics. |
| COM/test runner | `docs/xlide_vba_com_test_runner.md` | Preserve owned-host, read-only, timeout, cleanup, and no-SendKeys safety contracts while expanding test surfaces. |
| Performance | `docs/xlide_performance_budgets.md` | Use measured budgets and correctness-preserving safety rules before optimizing analyzer or workbook paths. |
| Development principles | `docs/xlide_development_principles.md` | Keep v3 changes deterministic, provenance-backed, scoped, and aligned with XLIDE's source-of-truth rules. |

When a v3 workstream changes one of these areas, update both this roadmap and
the linked sub-roadmap or evidence file in the same change. If this roadmap and
a more specific evidence file disagree, the more specific file wins until this
roadmap is corrected.

## Workstream A: Binder Shadowing and Expression Binding

Purpose: refine the broad v2 backlog item into implementable binder slices.

- [ ] Define the exact bare-identifier precedence ladder for expression
  contexts:
  - procedure locals and parameters
  - procedure labels where syntactically relevant
  - procedure return variables
  - same-module private/public declarations
  - exported standard-module members
  - enum members and UDT/type names where expression use is valid
  - class/document/UserForm code names
  - runtime, host, and external reference symbols
- [ ] Model shadowing separately for declaration contexts, call contexts,
  assignment targets, member receivers, type positions, and `New` positions.
- [ ] Preserve no-diagnostic behavior for ambiguous exported names and unknown
  external references until the resolver can prove the target.
- [ ] Add external-reference constants and globals as an explicit metadata-backed
  identifier source, with provenance and ambiguity behavior.
- [ ] Add deterministic arbitrary-expression binding only for proven expression
  shapes:
  - parenthesized expressions
  - known function/property return values
  - known member chains
  - collection/default-member calls where metadata proves the returned type
  - `With` receivers and nested `With` receiver stacks
- [ ] Add workbook fixture scenarios covering shadowing, duplicates,
  multi-workbook isolation, live-source overlays, and unresolved external names.

Definition of done:

- The binder can explain why a name resolves, why it is ambiguous, or why XLIDE
  intentionally stays silent.

## Workstream B: Object Member Model

Purpose: broaden object-member authoring and diagnostics beyond the v2 shipped
source-backed and generated Excel slices.

- [ ] Add designer-backed UserForm metadata ingestion sufficient to prove form
  controls, control types, and control event surfaces.
- [ ] Offer UserForm form/control event handlers only when designer-backed
  metadata proves the control/event surface.
- [ ] Model declared `Event` members with declarations, signatures, visibility,
  and source spans.
- [ ] Model `WithEvents` bindings and their event handler authoring surface.
- [ ] Verify default-member/direct object-expression behavior, including
  `VB_UserMemId = 0`, before inferring direct object usage such as
  `textValue = p`.
- [ ] Extend source member signatures with richer parameter metadata,
  declaration spans, return/write types, docs, and provenance.
- [ ] Add document/UserForm designer-backed members to the same object-member
  contract used by source-backed classes and host metadata.

Definition of done:

- UserForm/document member and event surfaces are deterministic and do not invent
  controls, events, or members from names alone.

## Workstream C: External Object/Member Metadata

Purpose: let downstream developers describe referenced libraries, add-ins, and
host extensions that XLIDE cannot parse from workbook source.

- [ ] Define a versioned external object/member metadata schema.
- [ ] Support member names, kinds, signatures, parameter docs/types, return
  types, examples, mutability, exhaustiveness, and provenance.
- [ ] Define reload behavior and validation diagnostics for malformed metadata.
- [ ] Define deterministic precedence:
  - workbook source symbols win for workbook-owned members
  - inline docs enrich source symbols
  - external metadata describes explicitly declared external/extension members
  - curated host/runtime metadata remains the built-in fallback
- [ ] Add completion, hover, signature help, member-call diagnostics, assignment
  diagnostics, and no-diagnostic controls for external metadata.
- [ ] Ship downstream developer documentation with schema examples,
  troubleshooting, provenance rules, and verification steps.

Definition of done:

- A downstream developer can author metadata, reload it, verify `object.`
  completion, and troubleshoot missing members without reading XLIDE source.

## Workstream D: Host Metadata Completeness

Purpose: make generated host metadata broad, auditable, and safe enough to power
hard diagnostics type by type.

- [ ] Normalize the repo-local `reference/` dump corpus into generated metadata
  that production extension code can consume without reading `reference/` at
  runtime.
- [ ] Preserve Office/version/source provenance for every type, member,
  signature, return type, enum, event, default member, and writable/read-only
  fact.
- [ ] Diff generated dumps against curated metadata and official documentation
  where available.
- [ ] Add oracle spot checks for behavior a reference dump cannot answer.
- [ ] Produce coverage reports by host/library type so gaps are visible.
- [ ] Resolve remaining Excel object receiver chains beyond the v2 simple
  return-type and collection-default `Item` paths.
- [ ] Promote generated host types into hard `member-not-found` only
  type-by-type after coverage reports and representative oracle controls prove
  the surface complete enough for red diagnostics.
- [ ] Require every host metadata expansion to add coverage for completion,
  hover/signature docs where applicable, member-call arity/type, assignment
  validation, `member-not-found`, and no-diagnostic controls for incomplete or
  non-exhaustive surfaces.

Definition of done:

- Host metadata has auditable provenance and hard absence diagnostics only come
  from exhaustive receiver surfaces.

## Workstream E: Workbook-To-Workbook Transfer

Purpose: support explicit module transfer between workbooks without crossing
analysis scopes or guessing user intent.

- [ ] Add source workbook and destination workbook selection.
- [ ] Add module/class selection with a side-by-side preview.
- [ ] Add conflict handling for existing destination modules.
- [ ] Add backup/snapshot hooks before destination workbook mutation where
  practical.
- [ ] Preserve multi-workbook analysis isolation; transfer previews must not
  imply cross-workbook project binding.
- [ ] Record write-audit entries and changed/skipped/failed summaries.

Definition of done:

- Workbook-to-workbook transfer is explicit, previewed, auditable, and
  recoverable.

## Workstream F: Performance and Scale

Purpose: make XLIDE faster on large workbooks without changing diagnostic
meaning, skipping modules, or introducing stale symbols.

Safety rule: performance work must prove equivalent results before it changes
the analyzer, project index, completion, or workbook sync behavior. Prefer
measurement, stale-result protection, progress, and cancellation before
incremental parsing/indexing.

- [ ] Add large-workbook fixture coverage or a deterministic synthetic project
  generator that exercises parser, symbol graph, diagnostics, completion setup,
  and workbook-wide analysis.
- [ ] Add stress tests for many modules, large modules, and diagnostic-heavy
  modules without changing expected diagnostics.
- [ ] Add timing/budget assertions for pure analyzer paths where results are
  stable enough across machines.
- [ ] Add stale-result protection for live diagnostics so older async analysis
  cannot overwrite newer document results.
- [ ] Add cache invalidation rules for source text, workbook identity, module
  identity, workbook state, settings, metadata version, and project symbol graph
  changes.
- [ ] Ensure live diagnostics never block typing; any optimization must preserve
  the same diagnostic set for the same source/settings/project inputs.
- [ ] Add cancellation for long-running workbook analysis, test discovery, and
  sync preview work where VS Code can abandon stale operations safely.
- [ ] Add status/progress reporting for work that exceeds the v2 latency
  thresholds in `docs/xlide_performance_budgets.md`.
- [ ] Add incremental parsing/indexing only after fixture and cache-invalidation
  coverage can prove equivalent symbols and diagnostics.

Definition of done:

- XLIDE stays responsive on large real-world workbooks, and slow paths have
  measurable budgets, cancellation, and visible progress.
- Performance improvements are correctness-preserving and do not silently skip
  diagnostics, project symbols, workbook modules, or metadata.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/roadmap_version_3.x.md`
- `docs/spec/MS-VBAL.version.md`
- `docs/spec/MS-VBAL.verification-map.md`
- `docs/xlide_vba_language_service_roadmap.md`
- `docs/type_analysis_corpus_coverage.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/xlide_vba_analysis_test_strategy.md`
- `docs/xlide_vba_com_test_runner.md`
- `docs/xlide_development_principles.md`
- `docs/xlide_performance_budgets.md`
- `docs/xlide_external_member_metadata.md`
- `docs/excel_reference_coverage.md`
- `scripts/generate-excel-reference-metadata.mjs`
- `scripts/generate-office-reference-metadata.mjs`
- `src/analyzer/symbols/projectIndex.ts`
- `src/analyzer/completion/memberAccess.ts`
- `src/analyzer/host/excelReferenceMembers.ts`
- `src/analyzer/host/excelObjectModel.ts`
- `src/vbaLanguageProviders.ts`
- `src/vbaModuleAnalysis.ts`
- `src/vbaProjectAnalysis.ts`
- `src/vbaWorkbookAnalysis.ts`
- `tests/fixtures/vbaProjects/`
- `syntax_corpus/README.md`
- `syntax_corpus/managed_backlog.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
- `syntax_corpus/oracle/README.md`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
