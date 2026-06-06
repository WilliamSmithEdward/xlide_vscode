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
- [x] Parameterless value-call binder slice: single-name expression inference
  now resolves zero-argument Function/Property Get references as return values
  when source/project/runtime signatures prove the target. This covers bare,
  parenthesized, and module-qualified project forms while ambiguous bare project
  functions remain silent.
- [x] Nested `With` receiver-stack slice: member resolution now tracks the
  active procedure-local `With` stack lexically and resolves nested
  `With .Member` receivers outer-to-inner. Completion and hard diagnostics use
  the same leading-dot receiver for source-backed class members, UDT fields, and
  promoted host surfaces.
- [x] Parenthesized member-receiver slice: member resolution now unwraps
  parenthesized receiver expressions such as `(ws.Range("A1")).Value` and
  `(p.Child).Save` when the inner expression shape is already deterministic,
  giving completion and hard diagnostics the same receiver type without guessing.
- [x] Member-expression type-inference slice: argument and assignment
  diagnostics now consume proven source-backed and host member return types in
  RHS expressions, including deterministic member-call returns such as
  `ActiveSheet.Range("A1")`. Bare member functions with required arguments stay
  unknown unless an actual call supplies arguments.
- [x] Host collection default-`Item` inference slice: assignment diagnostics now
  treat called host collection-valued properties such as
  `Worksheet.ListObjects("Tests")` and `ListObject.ListColumns("Passed")` as
  returning the collection's proven default `Item` type when the property itself
  has no parameter signature. Signature-backed property calls such as
  `Range("A1")` keep their declared return type.
- [x] Source-shadowed runtime callable slice: argument count/type diagnostics,
  nested return-type inference, expression-call parenthesis diagnostics, and
  runtime-only `Call` syntax checks now use source-name provenance before
  falling back to runtime metadata. Untyped locals, parameters, module
  constants/globals, enum members, and other non-callable source names suppress
  bare runtime signatures instead of producing hard runtime-shaped squiggles.
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
  assert flattened workbook diagnostics, including exact module/line/severity
  and quick-fix evidence for expected problems. The first stdvba canary keeps
  one real `Property Get ... End Function` structural error visible while
  guarding against false positives from module-qualified constants,
  project-visible ByRef return assignment helpers, default Win64 conditional
  branches, and verified `Worksheet.Buttons` host metadata. The binder/overlay
  canary adds ambiguous bare procedure no-diagnostic coverage, module-qualified
  arity evidence, source-backed `member-not-found` for missing or private
  standard-module qualified members, and same-workbook live-source overlays with
  cross-workbook overlay isolation. The project-type canary locks workbook-scale class, UDT,
  enum, and ambiguous type-name binding into the same exact diagnostic harness:
  unambiguous source-backed class/UDT receivers can produce hard
  `member-not-found`, ambiguous type names produce a targeted type diagnostic,
  ambiguous receivers do not cascade into guessed member diagnostics, and
  module-qualified enum reads stay quiet under `Option Explicit`. The
  external-reference canary keeps common unresolved `Scripting.Dictionary`,
  `ADODB.Stream`, `CreateObject`, `TextCompare`, and `VBA.CStr` shapes quiet
  until explicit external metadata exists.

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

Progress:

- [x] Event declaration module-kind slice: `Event` declarations in standard
  modules now produce a compile-equivalent red diagnostic from module-kind
  evidence, while class/document/UserForm declarations and inactive conditional
  branches stay quiet. A workbook-analysis module-kind canary records this next
  to the existing non-red wrong-module event-handler guidance.
- [x] WithEvents declaration slice: parser-backed `WithEvents` facts now produce
  red diagnostics for standard-module declarations, procedure-local
  declarations, `As New`, and array declarators. Module-level object-module
  `WithEvents` declarations stay quiet, and event-source type compatibility
  remains deferred until reference/object metadata can prove it.
- [x] Friend declaration slice: parser-backed `Friend` modifiers now produce
  red diagnostics for standard-module procedures and module variable
  declarations. Object-module `Friend` procedures and inactive conditional
  branches stay quiet, while broader `Friend` visibility/binding semantics
  remain deferred to project-level resolution.
- [x] Implements placement slice: token-backed `Implements` statements now
  produce red diagnostics when declared in standard modules, inside procedures,
  or after a procedure in an object module. Declaration-section object-module
  `Implements` statements and inactive conditional branches stay quiet, while
  interface member completeness remains deferred to the project binder.
- [x] Module declaration order slice: parsed module declarations (`Declare`,
  `Event`, module variables/constants, `Type`, and `Enum`) now produce a red
  diagnostic when they appear after an active procedure, including active
  `#If VBA7` branches and the selected branch of valid `#If`/`#ElseIf`/`#Else`
  declaration blocks. The diagnostic message names the active
  conditional-compilation branch so only-one-branch squiggles are easier to
  interpret. Later procedures remain accepted, and inactive conditional
  branches stay quiet.
- [x] Attribute placement slice: parser recovery now distinguishes exported
  member metadata attributes from executable-body `Attribute` statements.
  Unindented `Attribute ProcedureName...` lines targeting the current procedure
  remain accepted as exported-source metadata only after a module-level
  exported `VB_` attribute marker such as `Attribute VB_Name` is present. Other
  procedure-body `Attribute` statements produce a red placement diagnostic
  backed by VBE `CodeModule.AddFromString` compile evidence. A dedicated
  exported-file import oracle/source-mode split remains the right owner for
  broader visible `Attribute` behavior.
- [x] Parameter-list hardening slice: parsed procedure parameter flags now
  produce a red diagnostic when `ParamArray` appears in the same parameter list
  as any `Optional` argument, and when an explicit `ParamArray` element type is
  not `Variant`. `Property Let`/`Property Set` declarations with no final value
  parameter and `Property Set` declarations whose final value parameter is a
  known intrinsic scalar now also produce red property-shape diagnostics.
  Paired `Property Get` and `Property Let`/`Property Set` declarations for the
  same property now also require matching index-argument count, array shape,
  passing mode, and known scalar/Variant types before the setter's final value
  parameter. The existing `ParamArray` final-position and
  required-after-optional rules remain separate.
- [x] RaiseEvent target slice: settled `RaiseEvent` statements now produce a red
  diagnostic when the target event is not an active `Event` declaration in the
  same module. Declared same-module events and inactive conditional branches
  stay quiet, while event argument/signature validation remains deferred.
- [x] Procedure-label hardening slice: active procedure-local label declarations
  now produce a red diagnostic when the same named label or normalized decimal
  line label is declared more than once in one procedure. Separate procedures
  keep independent label scopes, and inactive conditional branches stay quiet.
- [x] Enum member uniqueness slice: active top-level `Enum` blocks now produce
  a red diagnostic on the second and later case-insensitive duplicate member
  name inside the same enum. Whole `Enum` blocks in inactive
  conditional-compilation branches stay quiet; directives nested inside an
  `Enum` remain deferred until the parser models that source shape explicitly.
- [x] Ambiguous enum member reference slice: focused VBE oracle probes show that
  duplicate member names across separate `Enum` blocks compile until the shared
  name is referenced bare. Value-position unqualified reads now produce a red
  `ambiguous-enum-member` diagnostic when they resolve to multiple visible enum
  member definitions, while declaration-only duplicates, qualified reads,
  same-module bindings, and local shadows stay quiet.
- [x] Conditional branch-order slice: peer `#ElseIf`/`ElseIf` branches after
  `#Else`/`Else`, plus duplicate `#Else`/`Else` branches, now produce red
  diagnostics. Conditional-compilation branch order is checked structurally,
  while normal `If` blocks still respect inactive conditional-compilation
  regions. Nested blocks keep independent branch state, and module-declaration
  placement diagnostics are suppressed inside malformed conditional-compilation
  blocks so the primary branch-order error does not create misleading cascade
  squiggles.
- [x] Module declaration recovery refinement: nested `Type` and `Enum` blocks
  inside procedures are now consumed as invalid procedure-body statements, so
  `module-declaration-in-procedure` reports the opening `Type`/`Enum` and
  misleading `End Sub` outside-procedure cascades stay quiet.
- [x] For/Next block hardening slice: parser-backed `ForBlock` metadata now
  records simple opener and closer control variables, and active `Next name`
  statements now produce a red diagnostic when the name does not match the
  active `For` or `For Each` control variable. Omitted `Next` variables,
  matching names, nested loops, and inactive branches stay quiet.
- [x] For Each control-variable type slice: active `For Each` loops now produce
  a red diagnostic when the control variable is a known intrinsic scalar, an
  array variable, or a project UDT/Enum control when project type metadata is
  available. `Variant`, `Object`, host/project object-looking types, implicit
  `Variant`, ambiguous types, and inactive branches stay quiet.
- [x] For Each source type slice: parser-backed `ForBlock` metadata now records
  the `In` source expression, and active `For Each` loops now produce a red
  diagnostic when a simple source name resolves to a known intrinsic scalar.
  Arrays, `Variant`, `Object`, object-looking/unresolved sources, member chains,
  and inactive branches stay quiet until deeper enumerable inference is proven.
- [x] Array ReDim hardening slice: parsed variable declarations now distinguish
  dynamic array declarators from fixed-bound arrays, and active `ReDim` or
  `ReDim Preserve` statements now produce a red diagnostic when the target
  resolves in procedure/module scope to a fixed-size array. Dynamic arrays,
  undeclared targets, local dynamic shadows, and inactive branches stay quiet.
- [x] Array assignment binding slice: active scalar assignments now produce a
  red diagnostic when a simple RHS name resolves to a dynamic or fixed array
  variable and the target resolves to a known intrinsic scalar. `Variant`
  targets, array targets, indexed element reads, unresolved names, and inactive
  branches stay quiet.
- [x] Array bound intrinsic slice: active `LBound`/`UBound` calls now produce a
  red diagnostic when the first argument is a simple name resolving to a known
  intrinsic scalar. Dynamic/fixed arrays, `Variant`, unresolved names,
  member-expression arguments, and inactive branches stay quiet.
- [x] ReDim Preserve dimension slice: straight-line active `ReDim` shapes now
  feed a deterministic runtime diagnostic when a later `ReDim Preserve` changes
  a known non-final dimension, the known dimension count, or a known
  final-dimension lower bound. Last-dimension upper-bound-only resizes stay
  quiet, and shapes learned inside nested blocks do not leak outward.
- [x] Array Erase hardening slice: active `Erase` statements now reject
  target-list entries that are clearly arbitrary expressions, such as literals
  or arithmetic. Binder-backed Erase checks now also reject simple targets that
  resolve to known non-Variant scalar declarations such as `Object` or `Long`;
  arrays, `Variant`, implicit Variant, unresolved names, non-simple targets, and
  inactive branches stay quiet.
- [x] Type-declaration character slice: legacy suffix-only declarations such as
  `name$`, `total&`, and `GetName$()` now normalize to the base name and inferred
  suffix type. VBE-rejected suffix-plus-`As` declarations now report a red
  `type-declaration-character-as-clause` diagnostic for variable declarations,
  Const declarations, parameters, UDT fields, and Functions, while the
  VBE-accepted `Property Get Name$() As String` control stays quiet.
- [x] Set scalar-target regression: `Set` assignments now have an exact
  oracle-backed guard for scalar `String` targets receiving object expressions
  such as `Set text = New Collection`; the declared target type owns the error.

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
