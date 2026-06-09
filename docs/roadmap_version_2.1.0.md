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

External metadata authoring/reload and workbook-to-workbook transfer are tracked
in `docs/roadmap_version_2.2.0.md`.

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

- [x] Define the exact bare-identifier precedence ladder for expression
  contexts:
  - procedure locals and parameters
  - procedure return variables
  - same-module private/public declarations
  - exported standard-module members
  - enum members and UDT/type names where expression use is valid
  - class/document/UserForm code names
  - runtime, host, and external reference symbols
- [x] Model shadowing separately for declaration contexts, call contexts,
  assignment targets, member receivers, type positions, and `New` positions.
- [x] Preserve no-diagnostic behavior for ambiguous exported names and unknown
  external references until the resolver can prove the target.
- [x] Add external-reference constants and globals as an explicit
  metadata-backed identifier source, with provenance and ambiguity behavior for
  verified VBA runtime constants, generated Excel/Office enum constants, and
  proven host globals. Unknown COM/reference-library symbols remain quiet until
  Priority 2 metadata proves them.
- [x] Add deterministic arbitrary-expression binding only for proven expression
  shapes:
  - parenthesized expressions
  - known function/property return values
  - known member chains
  - collection/default-member calls where metadata proves the returned type
  - `With` receivers and nested `With` receiver stacks
- [x] Add workbook fixture scenarios covering shadowing, duplicates,
  multi-workbook isolation, live-source overlays, and unresolved external names.

Status: **complete for 2.1.0**. The source-name resolver now owns
value/call/assignment/member/type/`New` binding. Procedure labels are
intentionally kept in the procedure-local control-flow label surface, because
they are not expression identifiers and need different syntax rules.
Unknown/non-proven host and reference-library symbols remain no-diagnostic by
policy until Priority 2 metadata promotes them.

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
- [x] Context-aware bare-identifier resolver slice: `ProjectIndex` now exposes
  a shared resolver that records the active context, winning precedence tier
  (`local`, `module`, or `project`), ambiguity, definitions, and explanation for
  bare identifiers. Go-to-definition and reference-scope resolution consume this
  resolver instead of carrying a separate precedence ladder. Runtime fallback
  diagnostics now also receive the project-visible source-name surface, so
  exported workbook names such as constants or enum members suppress bare
  runtime-shaped diagnostics while explicit `VBA.` calls remain runtime-bound.
- [x] Procedure-return binding closure: the shared resolver now treats the
  enclosing `Function`/`Property Get` name as its hidden return variable in
  expression, assignment-target, and member-receiver contexts, while leaving
  call contexts bound to the callable so recursive calls still resolve
  correctly. Declaration-site go-to-definition continues to resolve the actual
  procedure/accessor declarations.
- [x] Option Explicit and call-diagnostic resolver handoff: undeclared-variable,
  `unknown-call`, and `non-callable-call` now ask the shared source resolver for
  local/module/project bindings at the current procedure site. This keeps local
  and module declarations ahead of exported procedures, lets project-visible
  symbols satisfy Option Explicit without relying only on flat known-name sets,
  and preserves ambiguity silence for duplicate project call targets.
- [x] Ambiguous enum resolver handoff: `ambiguous-enum-member` now reports from
  the shared expression binding result instead of a separate local/project name
  ladder. It remains deliberately narrow: only enum-member-only ambiguous
  bindings across multiple visible Enum containers produce the red diagnostic,
  while local/module non-enum shadows and qualified reads stay quiet.
- [x] Const-assignment resolver handoff: `const-assignment` now resolves bare
  assignment targets through the shared assignment-target binding. Local
  declarations correctly shadow module-level constants, visible exported
  standard-module constants are caught when project context proves them, and
  ambiguous exported constants remain quiet rather than guessing.
- [x] Runtime integer-constant resolver handoff: `division-by-zero` and
  `runtime-argument-value` now use a scoped integer-constant lookup that asks
  the shared expression resolver before folding bare source names. Exported
  standard-module constants and enum members still feed deterministic runtime
  diagnostics, while local/module variables or other non-constant declarations
  with the same name suppress folding instead of producing false positives.
- [x] Verified external integer-constant slice: `division-by-zero` and
  `runtime-argument-value` now fold numeric values from verified VBA runtime
  constants and generated Excel/Office host enum constants when no source
  binding exists. Source declarations still shadow those external constants,
  ambiguous source bindings stay quiet, and qualified `VBA.` runtime constants
  participate in the same scoped lookup.
- [x] Verified external constant value-type slice: argument compatibility and
  object `Set` RHS checks now infer scalar types for verified VBA runtime
  constants and generated Excel/Office host enum constants when no source
  binding exists. Numeric enum constants behave as scalar `Long` values while
  preserving the declared enum owner in diagnostic messages; source
  declarations and ambiguous source bindings still suppress external inference.
- [x] Verified external object-global value-type slice: object `Set` RHS checks
  now infer bare host globals such as `Application`, `ActiveSheet`, and
  `ActiveCell` as object-valued expressions when no source binding shadows the
  name. Compatible/generic object targets stay quiet, while incompatible host
  object targets produce `assignment-object-type-mismatch`.
- [x] Assignment-target typing resolver handoff: `assignment-type-mismatch`,
  `set-required`, and `set-requires-object` now resolve bare assignment targets
  through the shared assignment-target binding before consulting legacy local
  type maps. Visible exported standard-module globals participate in scalar and
  object assignment diagnostics, while untyped local shadows and ambiguous
  exported globals stay quiet.
- [x] Member-receiver scalar resolver handoff: `scalar-member-access` now
  resolves bare member receivers through the shared member-receiver binding
  before consulting legacy local type maps. Visible exported standard-module
  scalar globals participate in invalid-qualifier diagnostics, while untyped
  local shadows, ambiguous exported globals, and non-bare member-chain tokens
  stay quiet.
- [x] Member-receiver external shadowing slice: member completion and hard
  `member-not-found` diagnostics now resolve local/module declarations before
  runtime objects, host globals, workbook code names, or standard-module
  qualifier fallbacks. Typed or untyped local declarations named like Excel
  globals such as `ActiveSheet` therefore use the source binding, and
  late-bound `Object`/`Variant` shadows stay quiet instead of receiving
  host-global absence diagnostics, even when completion can refine them from a
  simple `Set` assignment.
- [x] Array/scalar shape resolver handoff: `array-assignment-to-scalar` and
  `array-bound-requires-array` now resolve simple bare target/source/argument
  shapes through the shared resolver before consulting legacy local shape maps.
  Visible exported standard-module arrays and scalars participate in proven
  array/scalar diagnostics, while local shadows, ambiguous exported globals,
  indexed values, and member-expression arguments stay quiet.
- [x] ReDim target-shape resolver handoff: symbol extraction now preserves
  fixed-array bounds on variable symbols, and `fixed-array-redim` /
  `scalar-redim` resolve bare `ReDim` targets through the shared
  assignment-target binding before consulting legacy local/module maps. Visible
  exported standard-module scalars and fixed arrays participate, while dynamic
  arrays, local dynamic shadows, ambiguous exported globals, and undeclared
  targets stay quiet.
- [x] Erase target-shape resolver handoff: `erase-requires-array` now resolves
  simple bare `Erase` targets through the shared assignment-target binding
  before consulting legacy local shape maps. Visible exported standard-module
  non-Variant scalar/object globals participate, while arrays, Variant and
  implicit Variant targets, local shadows, ambiguous exported globals,
  unresolved names, and non-simple member/index targets stay quiet.
- [x] For Each shape resolver handoff: `for-each-control-variable-type` now
  resolves simple control variables through the shared assignment-target
  binding, and `for-each-source-type` resolves simple `In` source names through
  the shared expression binding before consulting legacy local shape maps.
  Visible exported standard-module scalars/arrays participate where proven,
  while Variant/object-like values, local shadows, ambiguous exported globals,
  unresolved names, and member-chain sources stay quiet.
- [x] Bare and module-qualified expression value-type resolver handoff:
  argument validation, ByRef exactness, assignment compatibility, and `Set` RHS
  object/scalar checks now resolve simple bare value identifiers through the
  shared expression binding before consulting legacy local type maps, and
  resolve `ModuleName.ValueName` reads through the matching source-backed
  module surface. Visible exported standard-module globals participate in
  proven object/scalar and ByRef diagnostics, while local shadows, ambiguous
  exported globals, untyped values, Variant values, unknown module qualifiers,
  and callable value references with required arguments stay quiet.
- [x] Incomplete member cascade resolver handoff: `invalid-expression-syntax`
  now asks the shared member-receiver resolver before reporting a generic
  trailing-dot incomplete-member diagnostic. Known scalar receivers, including
  visible exported standard-module scalar globals, are left to
  `scalar-member-access`; local shadows, ambiguous exported globals, object-like
  receivers, and unknown receivers continue through the syntax fallback.
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
- [x] Host `New` creatability deferral slice: `invalid-new-type-name` no longer
  hard-red-squiggles Excel host object-model types such as `Worksheet` after
  `New` without explicit creatability metadata or oracle evidence. Source-proven
  non-creatable project types, primitives, UDTs, Enums, and document modules
  remain red, while host/reference-library creatability stays deferred.
- [x] Live editor responsiveness slice: type-name completions in obvious `As`
  positions now return built-in/host candidates before project-wide context
  loading, semantic type coloring reuses a short-lived project-type snapshot
  instead of rebuilding the workbook on every token request, stale-but-current
  semantic colors stay visible while project types refresh in the background,
  automatic canonical casing uses fast local facts plus cached project context
  rather than waiting on workbook indexing, edited lines get a short idle
  canonical-casing pass as a fallback to cursor-leave events, diagnostics no
  longer re-run on every cursor movement, active-tab changes no longer re-run
  diagnostics for every open VBA document, live diagnostics publish a fast
  module-local pass before the full workbook pass settles, hover and signature
  help answer from cached/current-module facts while warming project context in
  the background, VBE-style Smart Enter and cursor-leave formatting add empty
  `()` to parameterless `Sub`, `Function`, and `Property Get` headers, and
  save-time symbol refresh uses the saved editor text instead of a bridge
  readback.
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
  analysis path. Module `#Const` values now drive active-branch semantic checks
  such as ByRef exactness in live mode; inactive branches stay quiet, with
  strict all-branches analysis kept as a separate future policy. `Option
  Explicit` also recognizes the `VBA` namespace and common compare aliases, and
  typed Function/Property Get fallthrough or recovered conditionally split
  `#If VBA7` Function headers no longer emit a return-assignment warning.
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
  until explicit external metadata exists. Workbook-analysis member diagnostics
  also preserve the shared source-binding precedence: local `ActiveSheet`
  shadows stay late-bound and quiet, while typed source-backed receivers still
  report real missing members instead of falling back to host globals.

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

- [x] Normalize the repo-local `reference/` dump corpus into generated metadata
  that production extension code can consume without reading `reference/` at
  runtime.
- [x] Preserve Office/version/source provenance for promoted generated types,
  members, signatures, return types, enum constants, and event exclusions.
- [x] Diff generated dumps against curated metadata and official documentation
  where available through coverage reports, generated provenance, and
  representative controls.
- [x] Add oracle/unit spot checks for behavior a reference dump cannot answer,
  including event exclusion, workbook/worksheet/range absence, and mixed
  sheet-item deferral.
- [x] Produce coverage reports by host/library type so gaps are visible.
- [x] Resolve remaining Excel object receiver chains beyond the v2 simple
  return-type and collection-default `Item` paths.
- [x] Promote generated host types into hard `member-not-found` only
  type-by-type after coverage reports and representative oracle controls prove
  the surface complete enough for red diagnostics.
- [x] Require every host metadata expansion to add coverage for completion,
  hover/signature docs where applicable, member-call arity/type, assignment
  validation, hard `member-not-found` when the receiver surface is exhaustive,
  and no-diagnostic controls for incomplete or non-exhaustive surfaces.

Status: **complete for the v2.1.0 promoted host metadata scope**. The generated
Excel/Office reference pipeline now feeds production metadata without runtime
`reference/` reads. It promotes 62 Excel runtime surfaces for completion,
hover, signature help, and receiver chains, and marks the 36 proven exhaustive
surfaces as hard-diagnostic surfaces with provenance. The hard-diagnostic set
covers the core application/workbook/worksheet/range surfaces, sheet/workbook
collections, windows, names, tables/list objects, charts/chart objects, shapes,
common range formatting objects, hyperlinks, areas, styles, page setup,
validation, and format conditions. Broader third-party/reference-library
authoring and reload remains in v2.2.0 external metadata work.

`WorksheetFunction` and Pivot object families are promoted metadata surfaces,
but remain outside hard `member-not-found`: development-oracle compile probes
accept unknown members on those receivers, so absence diagnostics there would
be false positives.

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
- [x] Promoted host-surface closure: generated `Application`, `Workbooks`,
  `Worksheets`, and `Sheets` metadata now join `Workbook`, `Worksheet`, and
  `Range` as exhaustive host surfaces. Missing members on those receivers now
  produce hard `member-not-found`, known generated members stay quiet, and
  Excel events remain excluded from object-member surfaces.
- [x] High-value host-family promotion: generated metadata now promotes
  `Window`/`Windows`, `Name`/`Names`, `ListObject`/`ListObjects`,
  `ListRow`/`ListRows`, `ListColumn`/`ListColumns`, `Chart`/`Charts`,
  `ChartObject`/`ChartObjects`, `Shape`/`Shapes`, `Font`, `Interior`,
  `Border`/`Borders`, `Areas`, `Hyperlink`/`Hyperlinks`,
  `FormatCondition`/`FormatConditions`, `Style`/`Styles`, `PageSetup`, and
  `Validation` into the same exhaustive hard-diagnostic path. Mixed
  `Sheets(index)` worksheet/chart receivers now report unknown members because
  every possible candidate receiver surface is exhaustive.
- [x] Pivot and WorksheetFunction metadata promotion: generated metadata now
  promotes `WorksheetFunction` plus Pivot table/field/item/cache/filter,
  calculated field/item, and cube field families for completion, hover,
  signature help, and receiver-chain inference. Focused development-oracle
  controls showed VBE accepts unknown members on representative
  `WorksheetFunction`, `PivotTable`, and `PivotField` receivers, so these
  surfaces intentionally remain non-exhaustive and do not produce hard
  `member-not-found`.

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
  Unindented `Attribute ProcedureName.VB_* = ...` lines targeting the current
  procedure are accepted as exported-source metadata only in the member
  metadata slot before executable body statements, even when upstream source
  has stripped the module-level `Attribute VB_Name` header. Other
  procedure-body `Attribute` statements produce a red placement diagnostic
  backed by VBE `CodeModule.AddFromString` compile evidence. The parser also
  preserves accepted member attributes on the procedure symbol so
  `VB_UserMemId` metadata can flow to source-backed member surfaces.
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
  member definitions through the shared source resolver, while declaration-only
  duplicates, qualified reads, same-module/local/module shadows, and mixed
  non-enum bindings stay quiet.
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
  available. Simple controls resolve through the shared assignment-target
  binding, so visible exported standard-module bindings participate while
  `Variant`, `Object`, host/project object-looking types, implicit `Variant`,
  local shadows, ambiguous exported globals, and inactive branches stay quiet.
- [x] For Each source type slice: parser-backed `ForBlock` metadata now records
  the `In` source expression, and active `For Each` loops now produce a red
  diagnostic when a simple source name resolves to a known intrinsic scalar.
  Simple source names resolve through the shared expression binding, so visible
  exported standard-module bindings participate while arrays, `Variant`,
  `Object`, object-looking/unresolved sources, local shadows, ambiguous exported
  globals, member chains, and inactive branches stay quiet until deeper
  enumerable inference is proven.
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
- [x] ReDim impossible-bounds slice: active `ReDim` and `ReDim Preserve`
  dimensions now produce a deterministic runtime diagnostic when explicit
  literal-style lower and upper bounds are both known and the lower bound is
  greater than the upper bound. Equal, increasing, upper-only, variable-backed,
  inactive, and known compile-invalid targets stay quiet.
- [x] Unallocated dynamic array access slice: straight-line local non-Static
  dynamic arrays now produce a deterministic runtime diagnostic when indexed or
  passed to intrinsic `LBound`/`UBound` before `ReDim`, or after `Erase`.
  Module-level arrays, parameters, `Static` locals, fixed arrays, helper-call
  arguments, assignment-touched arrays, nested runtime-block touches, unresolved
  names, non-intrinsic member calls, and inactive branches stay quiet.
- [x] Runtime conversion value slice: bare and `VBA.`-qualified `CDate` calls
  now produce a deterministic runtime diagnostic when the first argument is an
  empty string literal or plain ASCII non-date text. Date-looking strings,
  non-ASCII/localized-looking text, variables, expressions, non-`CDate`
  conversions, source-shadowed bare `CDate` names, and inactive branches stay
  quiet.
- [x] CVErr Error Variant inference slice: intrinsic `CVErr(...)` and explicit
  `VBA.CVErr(...)` now feed the shared argument/assignment type compatibility
  path as Error Variant values. Scalar targets produce deterministic runtime
  type-mismatch diagnostics, Variant targets stay quiet, and source-shadowed
  bare `CVErr` calls keep their source return type.
- [x] Null literal inference slice: `Null` now feeds the same shared
  argument/assignment type compatibility path as a dedicated Null value. Scalar
  targets produce deterministic runtime error 94 diagnostics, while Variant
  targets stay quiet.
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
- [x] Object variable not set slice: straight-line local non-Static object
  variables now produce a deterministic runtime diagnostic when used as member
  or `With` receivers before any proven `Set`, or after `Set ... = Nothing`.
  Parameters, module-level objects, `Static` locals, helper-call arguments,
  branch-ambiguous state, unresolved object types, and inactive branches stay
  quiet.

Definition of done:

- UserForm/document member and event surfaces are deterministic and do not invent
  controls, events, or members from names alone.

## Deferred To Version 2.2.0

External metadata authoring/reload and workbook-to-workbook transfer have moved
to `docs/roadmap_version_2.2.0.md` so Version 2.1.0 can stay focused on the
red-squiggly, binding, host metadata, and source-backed object/event authoring
workstreams.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/roadmap_version_2.1.0.md`
- `docs/roadmap_version_2.2.0.md`
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
