# XLIDE Roadmap Version 2.1.0

Forward backlog for work intentionally moved out of Version 2.0.x closeout.
Version 2.0.x is closed around its launch-hardening scope; Version 2.1.0 is
ordered by expected developer-experience impact, with "red squiggly"
completeness as the primary product bias. The highest-priority work expands
deterministic hard diagnostic coverage, keeps the daily edit loop as a baseline
expectation, and then broadens supporting language features and workbook
workflows.

## North Star

Version 2.1.0 should make XLIDE's red squiggles feel increasingly complete and
trustworthy while preserving the core rule: no hard diagnostics from guesses.
Every new binder, metadata, object, or workflow surface must be backed by source
facts, generated metadata with provenance, the Microsoft specification, or
focused VBE/Excel oracle evidence.

## Priority Model

Developer-experience priority is based first on hard-diagnostic impact:

- Whether it closes an important false-negative gap for compile-equivalent or
  deterministic runtime diagnostics.
- Whether it increases trust in red squiggles without increasing false
  positives.
- Whether it gives `member-not-found`, call validation, assignment validation,
  `Set`/object diagnostics, or type mismatch diagnostics a stronger proven
  receiver/signature/type base.
- How often a VBA developer feels the improvement during normal editing.
- How many later roadmap features it unblocks.
- Whether the work can be delivered as deterministic, testable vertical slices.
- Whether it reduces support burden by making state, provenance, and failures
  visible to the user.

Recommended sequencing:

1. Deepen binding and expression resolution because it improves completions,
   hover, navigation, rename, and hard-diagnostic precision from the same facts.
2. Broaden host metadata type by type after coverage and provenance make absence
   diagnostics safe.
3. Expand object/event authoring where XLIDE can prove source or designer facts.
4. Give downstream developers a supported metadata path for gaps XLIDE cannot
   infer from workbook source.
5. Add workbook-to-workbook transfer after the core authoring loop is stronger.

Responsiveness, stale-result safety, cancellation, cache correctness, progress
reporting, and measured performance remain non-negotiable engineering standards
for every 2.1.0 slice. They should be implemented as part of the relevant
feature work and tracked against `docs/xlide_performance_budgets.md`, not
treated as a separate roadmap priority.

## Related Open Roadmaps and Evidence

Version 2.1.0 workstreams should point to the more specific sub-roadmaps and
evidence files that already own detailed backlog, provenance, and verification
decisions. Use this roadmap for product-level scope, and use the files below as
the active detail layer.

| Area | Source | Use in 2.1.0 |
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
| Development principles | `docs/xlide_development_principles.md` | Keep 2.1.0 changes deterministic, provenance-backed, scoped, and aligned with XLIDE's source-of-truth rules. |

When a 2.1.0 workstream changes one of these areas, update both this roadmap
and the linked sub-roadmap or evidence file in the same change. If this roadmap
and a more specific evidence file disagree, the more specific file wins until
this roadmap is corrected.

## Priority 1: Expression Binding and Name Resolution

Purpose: make the core language service explainable enough to power complete,
low-noise hard diagnostics across the daily authoring loop. Completion, hover,
go-to-definition, rename, and find references should consume the same proven
facts rather than a separate looser resolver.

Developer-experience impact:

- Reduces false positives and false negatives caused by unresolved or
  mis-prioritized names.
- Gives hard diagnostics a shared resolver for ambiguous names, call targets,
  assignment targets, member receivers, type positions, and `New` positions.
- Gives developers consistent answers for "what does this name mean here?"
- Unlocks safer diagnostics and richer editor features across later object,
  host, and external metadata work.

Scope:

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
- [ ] Add external-reference constants and globals as an explicit
  metadata-backed identifier source, with provenance and ambiguity behavior.
- [ ] Add deterministic arbitrary-expression binding only for proven expression
  shapes:
  - parenthesized expressions
  - known function/property return values
  - known member chains
  - collection/default-member calls where metadata proves the returned type
  - `With` receivers and nested `With` receiver stacks
- [ ] Add workbook fixture scenarios covering shadowing, duplicates,
  multi-workbook isolation, live-source overlays, and unresolved external names.

Progress:

- [x] Call-context binder slice: project-visible exported variables, constants,
  enum types, and enum members now resolve as `non-callable-call` targets
  instead of vague `unknown-call` misses. The rule stays silent when a visible
  procedure shares the same bare name or when duplicate non-callable project
  bindings make the target ambiguous.
- [x] Expression-call binder slice: unique exported project `Function`
  signatures now trigger `expression-call-requires-parens` for parenless
  argument calls in expression context, including module-qualified
  standard-module calls. Ambiguous bare project `Function` names stay silent.
- [x] Option Explicit project-call slice: known module-qualified project
  procedures no longer flag their standard-module qualifier as an undeclared
  variable in expression reads, including `Set item = ModuleName.Function()`
  and parameterless `Set item = ModuleName.Function` right-hand sides.
- [x] Option Explicit qualified-value slice: known module-qualified exported
  standard-module constants, variables, enum types, and enum members no longer
  flag their module qualifier as undeclared in expression reads, while unknown
  qualifiers still surface a useful `Variable not defined` diagnostic.
- [x] Module-qualified IntelliSense slice: standard module names now appear as
  qualifier completions, and `ModuleName.` member completion, hover, signature
  help, canonical casing, and source definitions use the same exported
  standard-module surface for procedures, Declares, globals, constants, enum
  types, and enum members.
- [x] Qualified type-name slice: declaration type positions now resolve and
  complete `ModuleName.TypeName` for visible project UDTs/enums/classes, with
  hover, semantic coloring, canonical casing, definition/reference matching,
  and `New` diagnostics all sharing the qualified type resolver.
- [x] Standard-module tree rename slice: renaming a standard module from the
  Explorer tree now rewrites bound `ModuleName.MemberName` and
  `ModuleName.TypeName` qualifier tokens, including self-references retargeted
  to the renamed module URI, while comments, strings, and unrelated modules
  remain untouched.
- [x] Live red-squiggly hardening slice: diagnostics and project indexing now
  default conditional activity to modern Windows Office (`VBA7 = True`,
  `Win64 = True`, `Win32 = False`, `Mac = False`) while preserving explicit
  compiler-constant overrides, so paired `#If VBA7` and `#If Win64`
  declarations do not leak duplicate or ambiguous symbols into the active
  analysis path. `Option Explicit` also recognizes the `VBA` namespace and
  common compare aliases, and typed Function/Property Get fallthrough or
  recovered conditionally split `#If VBA7` Function headers no longer emit a
  return-assignment warning.
- [x] Workbook-analysis fixture slice: machine-readable project fixtures can now
  assert flattened workbook diagnostics. The first stdvba canary keeps one real
  `Property Get ... End Function` structural error visible while guarding
  against false positives from module-qualified constants, project-visible
  ByRef return assignment helpers, default Win64 conditional branches, and
  verified `Worksheet.Buttons` host metadata. The binder/overlay canary adds
  ambiguous bare procedure no-diagnostic coverage, module-qualified arity
  evidence, and same-workbook live-source overlays with cross-workbook overlay
  isolation.

Definition of done:

- The binder can explain why a name resolves, why it is ambiguous, or why XLIDE
  intentionally stays silent.
- Any new hard diagnostic that depends on name resolution can point to the
  resolver evidence it used.

## Priority 2: Host Metadata Completeness

Purpose: broaden generated Excel/Office metadata safely enough to power hard
diagnostics type by type, with richer language features following from the same
proven metadata.

Developer-experience impact:

- Expands red-squiggle coverage for common Excel/Office object-model mistakes:
  missing members, invalid calls, invalid assignments, and object/scalar misuse.
- Turns reference coverage into visible product confidence instead of an opaque
  internal corpus.
- Keeps red `member-not-found` diagnostics limited to receiver surfaces that are
  proven exhaustive.

Scope:

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

Progress:

- [x] Worksheet host-surface gap: `Excel.Worksheet.Buttons` is now part of the
  curated exhaustive Worksheet surface with regression coverage for
  `ActiveSheet.Buttons` and declared `Worksheet` receivers.
- [x] Host receiver-chain triage slice: exhaustive `Workbook`/`Worksheet`
  absence diagnostics now have regression coverage through
  `Workbooks(1).Worksheets(1)`, keeping common chained Worksheet receivers on
  the same proven `member-not-found` surface while later collection/object types
  remain gated on generated coverage and oracle controls.
- [x] Range host-surface promotion slice: generated `Excel.Range` metadata is
  now marked exhaustive for hard `member-not-found`, with completion
  exhaustiveness coverage plus diagnostics for `ActiveCell`, declared `Range`,
  and chained `Worksheet.Range(...)` receivers while known Range members stay
  quiet.

Definition of done:

- Host metadata has auditable provenance and hard absence diagnostics only come
  from exhaustive receiver surfaces.

## Priority 3: Object Member and Event Authoring

Purpose: make source-backed classes, document modules, UserForms, and events
feel native in the editor without inventing members from names alone.

Developer-experience impact:

- Adds proven source/designer member facts that can safely power
  `member-not-found`, event signature, and assignment diagnostics.
- Improves completion and handler authoring for common VBA UI workflows.
- Makes event signatures, `WithEvents`, and designer-backed controls easier to
  discover and less error-prone.
- Extends object intelligence using facts developers already maintain in the
  workbook.

Scope:

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

## Priority 4: External Metadata Authoring and Reload

Purpose: let downstream developers describe referenced libraries, add-ins, and
host extensions that XLIDE cannot parse from workbook source.

Developer-experience impact:

- Lets teams add proven API surfaces that can safely power hard diagnostics for
  private or third-party dependencies when the metadata is exhaustive enough.
- Gives teams a practical path to improve completion, hover, signature help, and
  diagnostics for private or third-party dependencies.
- Lets advanced users close gaps without waiting for a full XLIDE release.
- Makes missing or malformed metadata visible through validation and reload
  feedback instead of silent failure.

Scope:

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

## Priority 5: Workbook-To-Workbook Transfer

Purpose: support explicit module transfer between workbooks without crossing
analysis scopes or guessing user intent.

Developer-experience impact:

- Improves a useful project-maintenance workflow, especially when moving modules
  between real workbooks.
- Keeps workbook mutation explicit, previewed, auditable, and recoverable.
- Has lower red-squiggle impact than binding and metadata, so it should follow
  the core authoring improvements unless customer demand changes.

Scope:

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

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/roadmap_version_2.1.0.md`
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
