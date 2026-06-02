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
- When an implementation intentionally leaves a syntax, typing, host, or
  realtime edge out of scope, record that gap in this roadmap or the linked
  corpus/type coverage backlog before moving on.
- New source-backed or metadata-backed language-service surfaces must preserve
  available XML documentation end-to-end. If completion, hover, signature help,
  go-to-definition, diagnostics, tests, or sidebar views expose a documented
  symbol without showing or preserving its documentation where that surface can
  reasonably display it, that is a tracked gap, not an acceptable final state.
- Curated object/member metadata must not become a one-off exception path. Any
  added curation must be tracked with provenance and tests, and object-access
  business rules must apply consistently across source-backed classes,
  user-defined types, document modules, UserForms, curated host objects, and
  external metadata.
- Hard object member diagnostics require an exhaustive member surface for the
  receiver. Curated subsets can power completion, hover, and soft guidance, but
  they cannot prove `member-not-found` unless the curated source explicitly
  declares and verifies exhaustiveness for that type.
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
- Problems output uses metadata-derived XLIDE source labels such as
  `XLIDE/VBE`, `XLIDE/runtime`, `XLIDE/risk`, and `XLIDE/style`, while the
  diagnostic code remains the stable rule id used by quick fixes.
- Same-module callable signatures are parsed for parameters and return types.
- High-confidence argument count and type mismatch diagnostics exist.
- Non-callable call statements are flagged, including bare variable statements
  that VBE Compile rejects.
- `ProjectIndex` exposes deterministic visible project type names for the
  current module: object module names plus visible `Type` and `Enum`
  declarations, preserving duplicates for ambiguity diagnostics.
- `ProjectIndex` exposes deterministic bare identifier names for the current
  module, so Option Explicit diagnostics can distinguish undeclared assignment
  targets from same-module declarations, exported standard-module globals, enum
  members, and document/UserForm code names.
- Navigation/reference/rename scope uses the same project binding rule for
  exported enum members as Option Explicit visibility: enum members inherit
  visibility from their containing `Enum`, while `Private Enum` members stay
  module-private and shadow exported names only inside their module.
- Bare identifier completion now consumes source-backed visible project
  declarations from `ProjectIndex.visibleIdentifierSymbols()`, so exported
  standard-module globals, enums, enum members, types, and procedures use the
  same cross-module visibility rule as diagnostics and navigation.
- Source-backed member completion/hover/diagnostics/navigation now consume
  `ProjectIndex.projectMemberSurfaces(moduleName)`, which combines object-module
  members with visible UDT field surfaces. UDT fields are exhaustive, writable,
  definition-backed members, so `Dim p As TPoint : p.` and `With p : .X` behave
  like ordinary VBA member access.
- Type names now share one resolver for completion, hover, and VS Code semantic
  tokens in declaration type positions and `New` expressions. Resolved VBA
  primitive types, Excel host types, classes, document modules, UserForms, UDTs,
  and enums can be colored without hardcoding project names into static grammar;
  expression-level `New` completion offers known creatable project classes and
  UserForms.
- Workbook class member completion has a first deterministic source-backed
  slice: variables declared as project classes can offer public/default-public
  class members and public fields at `object.`, including chaining through
  class members whose return type is another known project class.
- Source-backed workbook class property assignments now participate in
  deterministic assignment validation when XLIDE knows the receiver member and
  setter type. Read-only property assignment is compile-equivalent red; typed
  writable property assignment can be deterministic-runtime red; known object
  variables and object-valued source-backed members now require `Set`, and
  deterministic `Set` assignments reject incompatible project object types while
  respecting explicit `Implements` declarations. `Function` and `Property Get`
  return names now participate in the same assignment rules.
- Source-backed workbook class receiver/member binding now reports
  `member-not-found` when an unambiguous project class member surface proves the
  member is absent.
- Object member surfaces now carry an explicit exhaustiveness contract. Hard
  absence diagnostics require an exhaustive receiver surface; source-only
  document modules, UserForms, and curated host objects stay silent until their
  host/designer metadata is complete enough to prove absence.
- `Excel.Workbook` now has a dump-backed exhaustive member surface generated
  from `reference/excel/json/Workbook.json`, so `ThisWorkbook.DoesntExist`
  participates in the unified hard `member-not-found` rule while the extension
  still does not read `reference/` at runtime.
- Known source-backed and host/reference member signatures now feed
  `argument-count` and argument-type diagnostics for valid parenthesized member
  call contexts such as `Set wb = Workbooks.Open(...)`, explicit `Call`
  statements, and non-empty statement forms like
  `ActiveSheet.Range("A1")`.
- Verified runtime signatures with explicit parameter lists now feed
  `argument-count` too, so calls such as `MsgBox()` are flagged for omitting the
  required `Prompt` while unparenthesized runtime statement signatures remain
  outside arity diagnostics until they have structured metadata.
- Source-backed workbook class member resolution now feeds go-to-definition for
  `object.Member` and current-object `Me.Member` references.
- Project type-name resolution now feeds go-to-definition and references for
  `As Person`/`New Person`; class modules resolve to their module top, while
  reference search lists actual type-use tokens rather than the implementation.
- Type-name diagnostics now use the same project/host/primitive resolver and
  shared type-position scanner, so valid workbook/Excel types stay quiet while
  invalid reserved/runtime names, visible project non-type declarations, and
  ambiguous visible project type names are flagged in `As`, `New`, return,
  parameter, UDT field, `TypeOf ... Is`, and `Implements` positions.
- Source-backed workbook class member references now resolve through the same
  member binder before textual fallback, so same-named members in different
  classes do not share a reference set.
- Source-backed workbook class member rename now uses that same binder, including
  declarations and function/property return-name occurrences. Project-defined
  class component rename is tree-only because VBA derives the class name from
  the class module/component name; the tree rename action calls pyOpenVBA
  `renameModule` and updates bound `As`/`New`/
  `TypeOf ... Is`/`Implements` references.
- `unknown-call` now consumes current-module-visible procedure names, so a
  `Private Sub` in another standard module or a public class member no longer
  suppresses a bare-call diagnostic.
- `undeclared-variable` now covers project-backed `Option Explicit` write/read
  positions: bare assignment/`Set` targets, RHS and call-argument reads,
  control-flow block headers, member receivers, and indexed bases. Missing
  `Option Explicit` continues to allow implicit Variant assignment.
- `missing-return-assignment` warns when a `Function` or `Property Get` never
  assigns its return variable. This is intentionally a type-safety warning:
  VBA falls through to the default value rather than raising a VBE compile error.
- `invalid-declaration-name` flags unbracketed reserved VBA identifiers such as
  `Dim` or `In` when they are used as procedure, variable, parameter, `Type`,
  `Enum`, field, or enum-member declaration names.
- Runtime metadata is curated explicitly; parameter types are not inferred from
  parameter names.
- Inline documentation comments support descriptive metadata plus optional
  `type`, `unit`, and `value` hints; object-module/class docs use a single
  documented module-header convention: a `'''` block directly above the first
  `Option` directive. Pressing Enter on a `'''` line continues the documentation
  comment marker, and Backspace clears empty auto-continued `''' ` or `' `
  markers in one keypress. Tab on those empty auto-continued comment markers
  escapes the marker and indents normally.
- Smart Enter auto-blocks safe VBA block openers through the same structural
  block helper used by linting, including procedures, `If ... Then`, `With`,
  `For`, `Do`, `While`, `Select Case`, `Type`, `Enum`, and `#If`; `With`
  uses the default `xlide.editor.blockLayout = "comfy"` layout: spacer line,
  one editable body line one real tab deeper, spacer line, then the closer.
  The optional `"compact"` layout removes spacer lines for both Smart Enter and
  block snippets. `With` starts the editable body line with `.` for immediate
  member completion.
- Block keyword completions remain full-block scaffolds for explicit Tab-driven
  shortcuts, while Smart Enter handles the line-by-line opener workflow. Active
  close-keyword suggestions consume the same smart-block stack as Smart Enter.
  Loop snippets mirror the iterator into `Next` through a transform rather than
  a second linked placeholder, so leaving the iterator field exits the
  cross-line selection behavior cleanly. Keyword snippets preserve literal tab
  indentation while still inheriting the current line's base indent; Smart
  Enter uses the same literal-tab block unit for every block archetype.
- Simple `For` / `For Each` iterator names now stay paired after insertion:
  editing the iterator in the opener or matching `Next name` updates the other
  side without relying on snippet mode.
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
- [x] Use category/equivalence fields when reporting workbook lint summaries.

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
- [x] Complete full `syntax_corpus` digestion into managed backlog categories
  (syntax, realtime recovery, type analysis, runtime resolution, project
  binding, host behavior, completion context, UserForm/designer symbols,
  limits, and legacy edges). Treat every Markdown case as raw material until
  promoted through spec, oracle, or deterministic XLIDE-owned evidence. See
  `syntax_corpus/managed_backlog.md`.
- [ ] Promote the `limits-boundaries` backlog into deterministic fixture
  builders before adding hard diagnostics for continuation-count, physical and
  logical line length, string literal/fixed-string size, identifier/module-name
  length, argument-count, array-dimension, or Excel host limits. Basic `_` line
  continuation tokenization/logical-line handling exists today; boundary
  diagnostics remain pending until backed by spec or oracle evidence.

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
- [x] Add compile-equivalent diagnostics for unbracketed reserved identifiers in
  declaration-name positions, while accepting bracketed foreign names.
- [x] Add scalar member-access diagnostics for declared intrinsic scalar
  receivers (`String`, numeric, `Boolean`, `Date`) after focused oracle
  verification: named scalar members are VBE Compile `Invalid qualifier`
  errors, and trailing scalar dots are VBE Compile `Syntax error`s.
- [x] Add compile-equivalent diagnostics for extra same-statement tokens after
  complete declaration type names, such as `Dim s As String junk`, after focused
  oracle verification of the representative `Dim` case.
- [x] Model fixed-length string declaration suffixes (`As String * n`) as an
  explicit parser/symbol fact for variable declarations and UDT fields, while
  normalizing `asType` to `String`. Hover, source-backed UDT member hover, and
  object-module public-member diagnostics now consume the shared fact instead of
  reparsing `asType` text; trailing-token diagnostics consume a recognized suffix
  before flagging extra same-statement tokens.
- [x] Add oracle-backed compile-equivalent diagnostics for literal fixed-length
  String sizes outside the VBE-verified `1..65526` range. Controls prove
  `* 1` and `* 65526` compile, while `* 0` and `* 65527` are rejected as
  `Invalid length for fixed-length string`.
- [ ] Finish fixed-length string behavior validation beyond the first-class
  declaration model:
  - accepted declaration shapes across class/document/UserForm modules beyond
    current standard-module controls
  - nonliteral length expressions and constant resolution
  - assignment/truncation behavior
  - interaction with scalar member access and type-declaration suffixes such as
    `$`
- [x] Add module-kind-sensitive diagnostics for object-module public
  declarations VBE rejects, including `Public Const`, public fixed-length
  strings, public arrays, public UDTs, and public `Declare` statements. The
  `object-module-public-member` rule uses module metadata before emitting hard
  errors because standard modules have different rules, and each branch has
  focused VBE oracle coverage.

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
- [x] Resolve public procedures across standard modules for module-qualified
  calls and ambiguous duplicate behavior in the first deterministic slice:
  ambiguous bare exported names stay silent, while `ModuleName.ProcedureName`
  resolves through the named standard module only.
- [x] Model procedure visibility for bare cross-module calls: same-module
  procedures are visible, exported standard-module procedures are visible, and
  private/object-module procedures are not treated as bare global callees.
- [x] Expose visible project-defined type names from `ProjectIndex` as binder
  groundwork: current-module `Type`/`Enum`, cross-module non-`Private`
  `Type`/`Enum`, and class/document/UserForm module names.
- [x] Expose visible bare identifier names from `ProjectIndex` for
  project-backed `Option Explicit` assignment-target diagnostics: same-module
  declarations, exported standard-module globals, visible enum members, and
  document/UserForm code names.
- [x] Extend `undeclared-variable` to conservative read-reference coverage:
  RHS and call-argument identifiers, block-header expressions, member receivers,
  and indexed bases, while skipping type-name, label, named-argument, and
  unresolved external-style call positions.
- [x] Resolve built-in VBA runtime constants and generated Excel enum constants
  in completion, hover, and project-backed `Option Explicit` diagnostics so
  common names such as `vbOKOnly`, `vbFalse`, and `xlUp` do not false-positive.
- [ ] Model full identifier shadowing and arbitrary-expression binding,
  including external-reference constants/globals and ambiguous external-reference
  behavior.
- [x] Resolve `As` type names against project classes, UDTs, enums, and host
  object types before flagging broad unknown type names.
- [x] Resolve enums and enum members across modules: public/default-public
  standard-module enum members now bind through `ProjectIndex` for go to
  definition, qualified definition, references, and rename scope; private enum
  members remain module-private.
- [x] Resolve UDT names across modules through the shared project type resolver
  used by completion, hover, semantic tokens, diagnostics, and type-definition
  navigation.
- [x] Add a shared project-analysis helper for project-aware test fixtures,
  lint surfaces, diagnostics, and editor providers, covering `ProjectIndex`
  construction, live current-module overlay, visibility-filtered names,
  exported signatures, visible project types, source-backed member surfaces,
  and live-IntelliSense tolerance for temporarily invalid modules.
- [x] Add shared workbook identity helpers for workbook cache keys,
  workbook-path comparisons, and case-insensitive module-name keys; route symbol
  index caches, member-completion caches/open-document overlays, command
  open-document overlays, and signature-drop warning suppression through that
  path.
- [ ] Add explicit multi-workbook isolation fixtures: two open workbooks with
  overlapping module, class, document-module, procedure, and member names must
  keep completion, hover, diagnostics, semantic coloring, references, rename,
  lint, and live-source overlays scoped to the workbook encoded in the module
  URI.
- [ ] Add machine-readable workbook-level fixture files for larger project
  analysis scenarios.
- [x] Keep current project-signature diagnostics stable under module order
  changes.

Definition of done:

- Cross-module procedure calls are validated deterministically.
- Shadowing behavior is explicit and tested.

## Workstream F: Object and Member Types

Purpose: validate Excel/VBA object use where receiver type is known.

- [x] Track simple preceding `Set` assignments to known object expressions for
  member completion, refining generic `Object`/`Variant` receivers while keeping
  declared host/project object types authoritative.
- [x] Build the first deterministic workbook class-member model from source:
  public/default-public methods and properties plus public fields, with return
  types, setter/write types, mutability, visibility, and inline XML
  documentation. Public constants are excluded from object-module member
  surfaces because VBE rejects them before they can become members.
- [x] Feed workbook class-member resolution into `object.` completion and
  return-type chaining for variables declared as project classes, such as
  `Dim p As Person: p.`.
- [x] Surface inline XML documentation for source-backed workbook class members
  in `object.` completion and member hover, including properties such as
  `p.Age`.
- [x] Add oracle-backed diagnostics for source-backed property assignment:
  nonnumeric string to typed writable property is deterministic runtime error
  13, and assignment to a read-only property is a compile-equivalent error.
- [x] Add oracle-backed `member-not-found` diagnostics for unambiguous
  source-backed workbook class receivers.
- [x] Add the first unified object-member exhaustiveness contract: source-backed
  class module surfaces can prove absence; document modules, UserForms, and
  curated host object surfaces cannot produce hard `member-not-found` until
  their full host/designer surface is modeled.
- [x] Promote the first dump-backed exhaustive host type: `Excel.Workbook`
  metadata generated from `reference/excel/json/Workbook.json`, enabling
  hard `member-not-found` for `ThisWorkbook.<missing>` through the same
  object-member rule contract.
- [x] Replace the one-off Workbook promotion script with a generalized Excel
  reference generator: `npm run generate:reference:excel` now scans
  `reference/excel/json`, emits promoted runtime metadata under
  `src/analyzer/host/excelReferenceMembers.ts`, and writes the deterministic
  coverage report at `docs/excel_reference_coverage.md`.
- [x] Preserve generated Excel reference member signatures and documentation in
  the promoted host metadata, enriching overlapping curated members through the
  shared completion/hover/signature-help member surface.
- [x] Generate Excel enum constants from `reference/excel/json` into the host
  metadata and expose them through the shared host resolver for completion,
  hover, and high-confidence diagnostics.
- [x] Promote generated non-exhaustive core Excel surfaces for completion and
  chaining: `Application`, `Worksheet`, `Range`, `Workbooks`, `Worksheets`, and
  `Sheets` now merge dump-backed member names, signatures, and documentation
  with curated return/chaining metadata.
- [x] Promote generated `Excel.Worksheet` to an exhaustive host diagnostic
  surface, so unknown members on `ActiveSheet` and declared `Worksheet`
  variables participate in hard `member-not-found` diagnostics.
- [x] Resolve collection-default `Item` calls in simple receiver chains, covering
  patterns such as `ThisWorkbook.Worksheets(1).Range("A1").` and
  `Workbooks(1).Worksheets(1).Range("A1").`; indexed `Sheets` routes through a
  merged worksheet/chart item surface for completion without treating
  non-exhaustive host surfaces as absence proof.
- [x] Reuse the shared member-completion binder for object member arity/type
  diagnostics when a source-backed or host/reference signature is known,
  including parenthesized member calls and parenless member call statements.
- [x] Split member-call signature validation from VBA call-statement syntax. VBE
  rejects standalone zero-argument calls such as `myFunction()`,
  `ThisWorkbook.CanCheckIn()`, `Application.Calculate()`, and
  `ActiveSheet.Range()` when they use empty parentheses without `Call`; the same
  call remains valid in expression context or with explicit `Call`. Known
  same-module and exported standard-module procedures use the project procedure
  surface, while non-empty standalone member/property forms such as
  `ActiveSheet.Range("A1")` and parenless member call statements such as
  `p.Save "ok"` stay on the normal signature-validation path.
- [x] Exclude host events from object-member surfaces. VBE rejects calls such as
  `ThisWorkbook.AfterSave True` with `Method or data member not found`, so event
  names are reserved for document-module handler authoring instead of appearing
  as object methods/properties.
- [x] Oracle-verify late-bound member access for `Object` and `Variant`
  receivers. VBE compile accepts unknown members on those receivers even after a
  simple `Set` assignment to a known host object, so hard member diagnostics opt
  out of completion-only `Set` refinement; a yellow
  `XLIDE(late-bound-invocation)` guidance rule remains separate settings-gated
  future work.
- [x] Oracle-verify public class module variables as object members:
  `Public Age As Integer` is indexed as a valid writable member, assignment/type
  diagnostics consume that writable field fact, and VBE oracle controls confirm
  nonnumeric string assignment raises runtime error 13 while compatible
  assignments compile/run. Public constants are not modeled as read-only object
  members because VBE rejects `Public Const` in object modules at declaration
  time.
- [x] Explore VBA default members, including exported attributes such as
  `Attribute Value.VB_UserMemId = 0`, before inferring direct object usage like
  `textValue = p`. The parser keeps dotted attribute names, the symbol graph
  attaches member attributes with source spans, and project member surfaces expose
  a `defaultMember` fact for `VB_UserMemId = 0`; direct object-expression
  inference remains deferred until an import-capable VBE oracle path can verify
  the runtime/compile behavior.
- [x] Define and implement deterministic class/module-level documentation for
  workbook object modules. VBA has no source-level `Class Person` declaration,
  so XLIDE uses one explicit convention: a module-header `'''` block directly
  above the first `Option` directive. `ProjectIndex.visibleTypeNames()` and
  `projectMemberSurfaces(moduleName)` preserve those docs so class-name type
  completion and hover show the same documentation as source-backed member
  surfaces.
- [x] Add first-class document-module event handler authoring for workbook and
  worksheet document modules. `src/analyzer/completion/eventHandlers.ts`
  scopes the metadata by module type and drives completion/insertions from one
  source of truth, not a parallel snippet list:
  - `ThisWorkbook` modules offer workbook handlers such as
    `Private Sub Workbook_Open()` and workbook event signatures.
  - Worksheet document modules offer worksheet handlers such as
    `Private Sub Worksheet_Change(ByVal Target As Range)` and
    `Worksheet_SelectionChange(ByVal Target As Range)`.
  - Existing handlers are not re-suggested, handler stubs are not offered inside
    procedure bodies, and `Private Sub ...` authoring inserts only the missing
    declaration tail.
  - `listModules` now carries an optional `documentType` (`workbook`,
    `worksheet`, `chart`) from hidden `VB_Base` CLSID metadata so chart modules
    are not treated as worksheets when the bridge can prove the subtype.
  - Event names do not appear in `object.` member completion, hover, signature
    help, or callable member diagnostics; those surfaces expose object
    properties/methods only.
  - Event-handler-shaped procedures in the wrong module now receive non-red
    `event-handler-module-scope` guidance: they may compile as ordinary
    procedures, but Excel will not wire them as events from that module.
- [ ] Extend document-module event handler authoring beyond the first
  workbook/worksheet slice:
  - Chart document modules offer chart handlers from the chart event surface.
  - UserForm modules offer form/control event handlers only when designer-backed
    metadata can prove the control/event surface.
- [ ] Extend the source member model to declared `Event` members, richer
  signatures, declaration spans, `WithEvents` bindings, and document/UserForm
  designer-backed members.
- [x] Feed workbook class-member resolution into signature help for
  source-backed method/function members, including inline XML summary and
  parameter docs.
- [x] Feed workbook class-member resolution into go-to-definition and
  member-call diagnostics.
- [x] Resolve `Me.Member` through the same source-backed current object module
  surface used for normal object receivers, merging with the known host surface
  when the current module has one.
- [ ] Extend external metadata files as an explicit object/member metadata
  source for referenced libraries, add-ins, team APIs, and host extensions that
  XLIDE cannot parse from workbook source.
- [ ] Ensure external object metadata can provide member names, kinds,
  signatures, parameter docs/types, return types, examples, and provenance.
- [ ] Drive host/object member curation toward completeness through the
  repo-local `reference/` dump corpus, not hand-entered guesses. In this repo,
  "reference dumps" means the checked-in material under `reference/`, including
  `reference/index.json`, `reference/members.json`, and the library folders such
  as `reference/excel/`, `reference/vba/`, `reference/office/`,
  `reference/msforms/`, `reference/vbide/`, and related COM/library dumps:
  - ingest the `reference/` dump corpus into a normalized metadata format
    (Excel has the first generator slice)
  - keep `reference/` as development context only; production extension code
    must not read it at runtime, and promoted metadata must be generated or
    checked in under `src/` with provenance
  - record Office/version/source provenance for every type, member, signature,
    return type, enum, event, default member, and writable/read-only fact
  - diff generated dumps against curated metadata and official docs where
    available
  - add oracle spot checks for behavior that a reference dump cannot answer
  - produce coverage reports by host type so gaps are visible
    (Excel has the first report)
  - mark a host type `exhaustive` only after its dump-backed surface is complete
    enough to prove member absence for that type/version
- [x] Keep generated/reference metadata separate from source-backed workbook
  symbols, then merge through the unified object-member contract. Source stays
  authoritative for workbook-owned classes/modules; dump-backed metadata becomes
  authoritative only for the exact referenced host/library surface it describes.
- [ ] Define a unified object-member rule contract used by source-backed
  classes, document modules, UserForms, curated host objects, and external
  metadata. The contract must identify whether a member surface is exhaustive,
  whether assignments are writable/read-only, which value type is accepted, and
  which diagnostics are allowed from that evidence.
- [ ] Require every new curated host/object metadata expansion to add coverage
  for completion, hover/signature docs where applicable, member-call arity/type,
  assignment validation, `member-not-found` behavior, and no-diagnostic
  controls for incomplete or non-exhaustive surfaces.
- [ ] Define deterministic precedence:
  source symbols win for workbook-owned members; inline docs enrich source;
  external metadata describes explicitly declared external/extension members;
  curated host/runtime metadata remains the built-in fallback.
- [ ] Resolve the remaining curated Excel object model receiver chains beyond
  simple return-type and collection-default `Item` paths.
- [ ] Promote dump-backed exhaustive Excel object types into hard
  `member-not-found` only type-by-type, after coverage reports and representative
  oracle controls prove the surface is complete enough for red diagnostics.
- [x] Add `member-not-found` only when receiver type is known and source-backed
  workbook class member metadata is unambiguous.
- [x] Add `set-required` and `set-forbidden`-family checks for deterministic
  object assignments. Known object variables and source-backed object-valued
  members (`Property Set` and public fields) now require `Set`; `Set` on
  source-backed scalar members reuses the existing stable
  `set-requires-object` code instead of adding a parallel spelling; incompatible
  known object RHS types now use `assignment-object-type-mismatch`. Function and
  Property Get return names now use the same scalar/object assignment path.
- [ ] Add downstream developer documentation/how-to for object member
  completion and external object metadata before shipping this workflow.

Definition of done:

- Object diagnostics do not guess from names.
- Host metadata has auditable provenance.
- A downstream developer can author external object/member metadata, reload it,
  verify `object.` completion, and troubleshoot missing members without reading
  XLIDE source.

## Workstream G: Realtime Experience

Purpose: keep the live editor useful while the user is mid-keystroke.

- [ ] Suppress hard errors for incomplete expressions where VBE behavior is not
  yet deterministically knowable.
- [x] Add a live-syntax state for incomplete member access, including trailing
  receiver dots such as `ThisWorkbook.` and bare leading dots inside active
  `With` blocks. While the cursor is in the incomplete edit, keep completion
  active and suppress the hard syntax diagnostic; once the edit settles,
  focus leaves the line, or workbook validation runs, promote the completed
  invalid form through an oracle-backed diagnostic. Oracle cases now cover
  `thisworkbook_trailing_dot_compile`,
  `bare_leading_member_access_inside_with_compile`, and the valid
  `leading_member_access_inside_with_compile` control.
- [ ] Make diagnostic ranges precise and stable.
- [x] Emit semantic tokens and hover for resolved type names in declaration type
  positions (`As Person`, `As Worksheet`, `As Currency`, function returns,
  parameters, UDT fields, and local/module variables) and `New Person`
  expressions. Primitive, host, and project names all flow through the same
  resolver used by type completion.
- [x] Apply safe canonical casing when typing a boundary after an identifier,
  moving the cursor away, or leaving the editor, so `person` in `As Person` /
  `New Person` and known members/runtime names normalize even when completion is
  dismissed.
- [x] Share structural block knowledge between linting and Smart Enter so safe
  block openers auto-insert their matching closer instead of maintaining
  procedure-only editor logic.
- [x] Use the same smart-block stack for close-keyword completion while keeping
  full block snippets available for explicit completion shortcuts.
- [x] Synchronize simple `For` / `For Each` iterator names between openers and
  matching `Next name` statements after either side is edited.
- [x] Conform Smart Enter and keyword snippet block archetypes to one
  literal-tab indentation rule, so `While`, `Do`, `If`, `With`, procedure,
  `Select Case`, `Type`, `Enum`, `#If`, `For`, and `For Each` bodies do not
  drift into separate space/tab behavior.
- [x] Project block keyword snippets from the shared `src/vbaLinter.ts`
  smart-block catalogue, and test that snippets, Smart Enter, and static
  language-configuration indentation stay aligned as one contract.
- [x] Conform static VS Code language-configuration indentation/folding regexes
  to the shared smart-block opener contract, including complete-opener gates
  for `For`, `For Each`, `Do While`, `Do Until`, `While`, `With`,
  `Select Case`, `If`, and `#If`.
- [x] Consolidate shared VBA source-text helpers (`stripVba`,
  identifier validation, comment/string-safe identifier occurrence search,
  line-start offsets, and leading whitespace) so providers, workbook lint,
  code actions, member completion, and class-module commands do not carry
  local duplicate regex pipelines.
- [x] Use metadata categories to tune Problems output and future filters.
- [x] Route workbook lint summaries through the shared diagnostic metadata
  catalogue, including structural block-balance codes, so command output and
  agent JSON report category/equivalence without local bucketing.
- [x] Route live diagnostics, current-module lint, and workbook lint through a
  shared module lint core so structural diagnostics, semantic diagnostics, and
  lint suppression directives cannot drift by surface.
- [x] Route current-module lint, workbook lint, and project-aware diagnostic
  fixtures through shared project-analysis option derivation so cross-module
  procedure, identifier, type, non-type, and member-surface rules cannot drift by
  surface.
- [x] Route live diagnostics, type semantic tokens, member completion, hover,
  and type completion through shared project-analysis project construction and
  option derivation so editor-facing project context cannot drift by provider.
- [x] Consolidate explicit `Call` argument syntax and standalone
  empty-parentheses call-statement diagnostics into
  `src/analyzer/call/callContext.ts`, alongside signature-help call-site
  detection and completion parenthesis insertion.
- [x] Route explicit `Call` argument-list quick fixes through the same
  `callContext` range helper used by diagnostics, including chained receivers
  and named arguments.
- [x] Route standalone empty-parentheses removal and invalid runtime `Call`
  rewrite quick fixes through shared `callContext` edit spans.
- [ ] Keep signature help, hover, completion, and diagnostics sharing the same
  symbol/type model. Signature help now reuses the member-completion route for
  member-call signatures; diagnostics now consume it for known member-call
  arity/type checks, and go-to-definition/references/rename now consume it for
  source-backed members, including current-object `Me.Member` references.
  Type-name completion, hover, semantic coloring, go-to-definition, references,
  and runtime-function misuse diagnostics now share one type resolver and
  type-position scanner; the VS Code completion provider now supplies project
  type candidates from `ProjectIndex.visibleTypeNames()` instead of a parallel
  type-name cache/builder. Call-site classification now lives in
  `src/analyzer/call/callContext.ts`, so signature help, completion parenthesis
  insertion, and bare-call diagnostics share the same
  parenthesized/parenless/explicit `Call` rules.

Definition of done:

- Typing incomplete code does not create avoidable noise.
- Completed invalid code produces deterministic, useful diagnostics.

## Workstream H: Developer VBA Test Runner

Purpose: let workbook developers write and run deterministic tests for their own
VBA code from XLIDE, using Excel COM as the execution host.

- [ ] Define an explicit VBA test discovery contract. Discovery must be
  annotation-driven or manifest-driven, not naming-heuristic-driven.
- [ ] Add a developer-facing `xlide.runVbaTests` command that runs selected
  tests through Excel COM.
- [ ] Run tests against a disposable workbook/session by default so test runs do
  not mutate the developer's open workbook unexpectedly.
- [ ] Reuse the workbook close/reopen/reset discipline from macro execution and
  warn when a workbook cannot be safely reopened in XLIDE's context.
- [ ] Add a small VBA assertion/support module or equivalent injected test
  runtime for assertions such as equality, truth, expected error, and expected
  no error.
- [ ] Support rich explicit test metadata:
  - tags/categories
  - skip reason
  - expected failure (`xfail`) reason
  - per-test timeout
  - owner or requirement id
  - expected error metadata
  - output/state assertions
- [ ] Support developer-friendly test selection and execution modes:
  - run all
  - run current test
  - run current module
  - include/exclude tags
  - rerun failed
  - fail fast
  - machine-readable automation mode
- [ ] Support setup/teardown patterns:
  - per-test setup and teardown
  - per-module setup and teardown
  - workbook/session setup and teardown
  - deterministic cleanup failure reporting
- [ ] Capture deterministic test results:
  - pass/fail/skip
  - expected failure (`xfail`)
  - unexpected pass (`xpass`)
  - compile errors
  - runtime errors, including error number and description
  - assertion failures
  - explicit test log/output written through the XLIDE test API
  - timeout and teardown failures
- [ ] Support tests that assert expected output, expected state, expected thrown
  error, and expected absence of errors.
- [ ] Return machine-readable JSON results for automation and render a concise
  Problems/Test Results view in VS Code.
- [ ] Keep the product test runner separate from the Excel/VBE oracle. The
  oracle validates XLIDE behavior; the test runner validates user VBA projects.
- [ ] Add fixture tests before enabling broad adoption.
- [ ] Fully document the downstream developer workflow before calling the test
  runner shipped:
  - how to author tests
  - how to mark test procedures explicitly
  - metadata for tags, skips, expected failures, timeouts, owners, and
    requirement ids
  - assertion API reference
  - expected-error and expected-output patterns
  - setup/teardown and test data patterns
  - filtering, rerun failed, and fail-fast workflows
  - workbook/session lifecycle
  - COM/Excel trust requirements
  - timeout and cleanup behavior
  - command palette and automation usage
  - result JSON schema
  - troubleshooting and known host limitations

Definition of done:

- A developer can author VBA tests, run them through Excel COM from XLIDE, and
  receive deterministic pass/fail/error/output results.
- The runner is opt-in, visible to the user, timeout-bounded, and safe against
  silent workbook mutation.
- The full workflow is documented for downstream workbook developers, including
  examples they can copy into real projects.

## Workstream I: Lint Suppression Directives

Purpose: give developers deterministic, VBA-comment-compatible control over
XLIDE lint diagnostics without changing VBA execution behavior.

- [x] Document the proposed suppression syntax in
  `docs/xlide_vba_lint_suppression_comments.md`.
- [x] Parse suppression directives only from explicit XLIDE comment directives.
- [x] Support module-level suppression.
- [x] Support next-member suppression for `Sub`, `Function`, `Property`,
  `Type`, and `Enum` blocks.
- [x] Support line-level and next-line suppression.
- [x] Support paired arbitrary block suppression.
- [x] Support optional diagnostic-code lists so developers can suppress one rule
  without hiding unrelated diagnostics.
- [x] Add directive diagnostics for malformed first-slice directives, late
  `disable-file`, and unknown-code directives without guessing the user's
  intent.
- [x] Add directive diagnostics for unbalanced block directives.
- [x] Preserve a suppressed-diagnostic count so ignored problems can be audited.
- [x] Add unit tests for file, line, next-line, code-list, ignored-comment,
  structural, malformed, and unknown-code first-slice cases.
- [x] Add unit tests for remaining next-member/block directive scopes, nesting
  edges, and unbalanced directive cases.

Definition of done:

- Suppression behavior is lexical, deterministic, auditable, and fully
  compatible with VBA because directives are comments.
- Suppressions hide XLIDE diagnostics only; they never make invalid VBA valid
  and never affect COM/test/oracle execution.

## Workstream J: XLIDE Activity Bar and Sidebar Panel

Purpose: make XLIDE feel like a first-class VS Code extension with a polished
Activity Bar icon and a full sidebar command/status surface for workbook
development.

- [x] Add the planned sidebar design/spec to `docs/xlide_sidebar_panel.md`.
- [ ] Add a dedicated XLIDE Activity Bar container using standard VS Code
  contribution points.
- [ ] Create a slick monochrome SVG Activity Bar icon that follows VS Code's
  icon style: simple line geometry, mask-friendly, readable at 24px, and
  theme-neutral.
- [ ] Expand the current workbook/module explorer into a full XLIDE sidebar
  with deterministic status sections and command surfaces.
- [ ] Keep the existing XLIDE workbook/module tree in the VS Code Explorer
  section even after the dedicated XLIDE Activity Bar/sidebar ships. Explorer is
  the file/navigation surface; the XLIDE sidebar is the product shell for
  health, commands, configuration, tests, lint summaries, and status.
- [ ] Add a unified configuration section/menu collection in the XLIDE sidebar:
  show effective settings, their source layer, validation state, quick actions,
  and links to edit user, workspace, and local workspace config.
- [ ] Add setup health checks with pass/warn/fail/unknown states for:
  - active XLIDE workbook/project context
  - workbook source sync/export mapping
  - Excel COM availability on Windows
  - trust access to the VBA project object model
  - macro/security prerequisites relevant to run/test workflows
  - workbook open/reopen safety
  - lint engine readiness
  - VBA test runner readiness once implemented
  - optional metadata/doc-comment support
- [ ] Add primary action buttons:
  - lint current module
  - lint workbook
  - run current test
  - run all tests
  - export/sync modules
  - open/reopen workbook
  - refresh project
- [ ] Add polished secondary panels for:
  - Problems summary by severity
  - Test summary by pass/fail/skip/xfail/xpass
  - Recent XLIDE operations and logs
  - Setup recommendations and quick fixes
  - Workbook/module metadata
- [ ] Keep status checks deterministic. If XLIDE cannot prove a requirement is
  met or missing, show `Unknown` with a concrete action instead of guessing.
- [ ] Add telemetry-free local persistence for collapsed sections and selected
  workbook context.
- [ ] Add VS Code UI/integration tests or fixture-backed provider tests for the
  sidebar model before broadening the UI.

Definition of done:

- XLIDE has a recognizable Activity Bar icon and a sidebar that surfaces the
  main workbook workflows without needing command-palette spelunking.
- Setup health, lint, test, sync, and workbook actions are available from one
  coherent place.
- Sidebar status is deterministic, accessible, theme-safe, and does not depend
  on heuristic project guesses.

## Workstream K: Safety, Trust, and Recovery

Purpose: make workbook mutation, COM execution, and macro/test workflows safe,
auditable, and recoverable for real user projects.

- [ ] Define XLIDE's trust model for workbook read, write, run, and test
  workflows.
- [ ] Require explicit user action for operations that mutate workbook contents
  or execute VBA.
- [ ] Add backup/snapshot strategy before workbook mutation where practical.
- [ ] Add an audit trail for XLIDE writes:
  - module changed
  - command/tool that changed it
  - timestamp
  - source path/workbook path
  - success/failure
- [ ] Add crash/timeout recovery for COM operations.
- [ ] Add workbook lock/open-state checks before write/run/test operations.
- [ ] Warn when a workbook is open outside XLIDE's controllable context.
- [ ] Add explicit workbook-to-workbook module transfer workflows for copying or
  exporting selected modules/classes from a source workbook into a destination
  workbook, with source/destination selection, preview, conflict handling,
  backup/snapshot hooks, and no implicit cross-workbook project analysis.
- [ ] Surface "what changed?" summaries after sync/write operations.
- [ ] Document trust-center, macro security, and VBA project access
  requirements without hiding them behind vague failures.

Definition of done:

- A developer can see when XLIDE is about to mutate or execute something, what
  it changed afterward, and how to recover from failed COM/workbook operations.
- Safety behavior is explicit and deterministic; XLIDE does not infer consent.

## Workstream L: Settings and Profiles

Purpose: give individuals and teams controlled configuration without weakening
the deterministic analyzer contract.

- [ ] Define workspace/project configuration for XLIDE.
- [ ] Separate team-shared settings from local-machine settings.
- [ ] Implement a deterministic configuration resolver with explicit provenance
  for every effective setting:
  - built-in defaults declared in the extension schema
  - user/machine defaults through VS Code settings
  - workspace/team config stored in a versionable workspace file such as
    `.xlide/settings.json`
  - workspace-local machine overrides stored separately, such as
    `.xlide/settings.local.json`, for paths, COM behavior, and other local
    choices that should not be committed
  - explicit command/session overrides for one-off operations
- [ ] Make built-in defaults configurable by users and overrideable by
  workspace config without weakening diagnostic determinism.
- [x] Add a deterministic VS Code extension setting, contributed through
  `package.json`, for editor block layout in Smart Enter and block snippets:
  - default `comfy`: spacer line, editable indented body line, spacer line,
    closer
  - optional `compact`: editable indented body line immediately followed by the
    closer
  - both modes must flow through the shared smart-block catalogue/helper so Tab
    snippets and Enter auto-blocking cannot diverge.
  - this setting is user/workspace editor preference, not `.xlide` workbook or
    project configuration.
- [ ] Surface all configuration through the unified XLIDE sidebar menu
  collection, while preserving normal VS Code Settings integration for users
  who prefer native settings UI.
- [ ] Add rule severity overrides with guardrails:
  - error to warning only when the rule permits downgrade
  - warning to off
  - no override that converts unknown behavior into a red diagnostic
- [ ] Add enabled/disabled rule-set profiles.
- [ ] Add COM/test-runner settings:
  - Excel visibility
  - timeouts
  - workbook reset behavior
  - trusted test folders
  - result output path
- [ ] Add sidebar/profile UI for active configuration.
- [ ] Add configuration validation diagnostics for malformed settings.
- [ ] Document precedence and conflict handling for defaults, user/machine,
  workspace/team, workspace-local, and command/session override layers.

Definition of done:

- Teams can share stable XLIDE behavior through project settings while each
  developer can keep local COM/test paths and machine-specific choices local.
- Settings cannot make analyzer behavior heuristic or unverifiable.

## Workstream M: Code Actions and Quick Fixes

Purpose: turn deterministic diagnostics into useful repairs where XLIDE can
prove the edit is safe.

- [x] Add code-action infrastructure for analyzer diagnostics.
- [ ] Add deterministic quick fixes for:
  - [x] add missing `Option Explicit`
  - [x] insert required `Call` parentheses
    through the shared `callContext` range helper
  - [x] add expression-call parentheses when a parenless function call is the
    whole assignment RHS
  - [x] remove invalid empty parentheses from standalone zero-argument calls
    through shared `callContext` edit spans
  - [x] rewrite runtime calls such as `Call DoEvents()` to bare statement syntax
    through shared `callContext` edit spans
  - [x] add missing block closer
  - [x] move misplaced `Option` statements before declarations
  - [x] split local `Dim` initializers into declaration plus assignment
  - [x] add `Set` when the assignment target is known object type
  - [x] remove `Set` when the assignment target is known scalar type
  - [ ] add missing required argument placeholder only when explicitly requested
  - [x] add lint suppression comment for a selected diagnostic
  - [ ] create procedure stub for unresolved calls only when the target location is
    explicit
- [ ] Add source actions:
  - [x] lint current module
  - [ ] export/sync current module
  - [ ] run current test when test runner exists
- [ ] Add tests for every generated edit, including formatting and range
  stability.
- [x] Avoid quick fixes for uncertain, host-dependent, or incomplete-code cases
  in the first shipped quick-fix slice; the resolver returns no action unless
  the diagnostic span proves a mechanical edit.

Definition of done:

- Code actions are precise, previewable, and backed by tests.
- XLIDE offers no quick fix when the correct edit would require guessing intent.

## Workstream N: Performance and Scale

Purpose: keep XLIDE responsive on large workbooks and during active typing.

- [ ] Define performance budgets for:
  - keystroke diagnostics
  - module parse/analyze
  - workbook-wide lint
  - project index rebuild
  - sidebar health refresh
- [ ] Add incremental parsing/indexing where deterministic and measurable.
- [ ] Add cache invalidation rules for source text, workbook state, metadata,
  and project symbol graph changes.
- [ ] Ensure live diagnostics never block typing.
- [ ] Add large-workbook fixture coverage.
- [ ] Add stress tests for many modules, large modules, and many diagnostics.
- [ ] Add cancellation for long-running lint/test/sync operations.
- [ ] Add status/progress reporting for work that exceeds user-visible latency
  thresholds.

Definition of done:

- XLIDE remains responsive on large real-world workbooks, and slow paths have
  measurable budgets, cancellation, and progress.
- Performance improvements remain deterministic and do not skip diagnostics
  silently.

## Workstream O: Release and Documentation Polish

Purpose: make XLIDE understandable and trustworthy for users who were not part
of development.

- [ ] Add marketplace-ready assets:
  - icon
  - screenshots
  - feature bullets
  - short walkthrough media or GIFs
- [ ] Add getting-started documentation.
- [ ] Add a sample workbook/project repo for demos and regression examples.
- [ ] Add feature walkthroughs:
  - opening a workbook
  - editing modules
  - linting
  - running macros
  - running VBA tests
  - using doc comments and metadata
  - adding external object/member metadata for `object.` completion
  - interpreting setup health
- [ ] Ship a full downstream developer how-to for external object/member
  metadata, including schema, examples, precedence, reload behavior, and
  troubleshooting.

Definition of done:

- A new developer can install XLIDE, understand what works on their platform,
  run the main workflows, and troubleshoot common setup issues without reading
  source code.

## Workstream P: Support Bundle and Diagnostics

Purpose: make user support and self-debugging possible without guessing or
asking for sensitive workbook contents.

- [ ] Add "Export XLIDE Support Bundle" command.
- [ ] Include non-sensitive diagnostic data:
  - extension version
  - VS Code version
  - platform
  - enabled XLIDE settings
  - setup health states
  - recent XLIDE command log
  - lint/test summary counts
  - COM availability and failure categories
  - workbook metadata summary without source code by default
- [ ] Add opt-in inclusion of anonymized lint/test reports.
- [ ] Add opt-in inclusion of selected logs.
- [ ] Redact workbook paths or allow path redaction.
- [ ] Add a local "Copy Diagnostics" quick action for setup failures.
- [ ] Document what the support bundle contains before export.

Definition of done:

- Users can produce a useful local support snapshot without exposing workbook
  source by default.
- The bundle helps diagnose setup, COM, lint, test, and sync issues
  deterministically.

## Immediate Next Steps

1. Use `syntax_corpus/managed_backlog.md` and
   `docs/type_analysis_corpus_coverage.md` to choose the next verified
   corpus additions for the project-wide binder.
2. Continue extending type-name diagnostics only where the binder can prove the
   shape, such as verified `New` creatability and qualified reference-library
   type names; keep broad unknown external reference names deferred.
3. Promote small `CANARY_*` cases through observe-only oracle fixtures when
   they become relevant to analyzer behavior.
4. Keep the VBA test runner as a planned workstream until its specs and fixture
   coverage are ready, and continue promoting small `CANARY_*` cases only when
   they become relevant to analyzer behavior.
5. Treat the XLIDE sidebar as the future product shell for lint, test, setup,
   sync, and workbook actions.
6. Track safety, settings, code actions, performance, release polish, and
   support diagnostics as product-maturity gates before a broad release.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/type_analysis_corpus_coverage.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/xlide_vba_linting_test_strategy.md`
- `docs/xlide_vba_lint_suppression_comments.md`
- `docs/xlide_vba_com_test_runner.md`
- `docs/xlide_external_member_metadata.md`
- `docs/xlide_sidebar_panel.md`
- `docs/xlide_development_principles.md`
- `src/vbaModuleLint.ts`
- `src/analyzer/diagnostics/ruleMetadata.ts`
- `syntax_corpus/README.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
- `syntax_corpus/managed_backlog.md`
- `syntax_corpus/oracle/README.md`
- `syntax_corpus/oracle/run_excel_vbe_oracle.py`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
- `tests/corpusProvenance.test.ts`
