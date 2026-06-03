# Syntax Corpus Managed Backlog

This file digests the current Markdown corpus into managed backlog categories.
It is a planning index, not an authority for hard diagnostics.

Every item below remains `pending-verification` until promoted through one of:

- MS-VBAL or another Microsoft primary source.
- A focused Excel/VBE oracle fixture.
- Deterministic XLIDE-owned metadata with unit or integration tests.

Do not use a Markdown expectation alone to justify a red diagnostic.

## Source Digest

| Source | Primary Role | Backlog Categories |
| --- | --- | --- |
| `26_class_and_userform_deep_edges.md` | Deep class-module and UserForm-module coverage for lifecycle events, `Me`, `Friend`, events, WithEvents, Implements, properties, designer symbols, default instances, exported metadata, realtime states, and completion contexts | `module-context`, `project-binder`, `object-member`, `userform-designer`, `completion-context`, `roundtrip-io`, `canary-verdicts`, `realtime-recovery` |
| `27_semantic_runtime_resolution_edges.md` | Semantic/runtime hardening for deterministic runtime faults, error-handler suppression, coercion, For Each restrictions, structured-branch targets, references, host event binding, public API visibility, WithEvents restrictions, macro discoverability, predeclared instances, conditional reachability, realtime semantic suppression, and canary verdicts | `runtime-resolution`, `type-analysis`, `project-binder`, `module-context`, `host-behavior`, `object-member`, `canary-verdicts`, `realtime-recovery` |
| `excel_vba_realtime_analysis_test_corpus.md` | Broad baseline corpus for visible realtime syntax analysis | `syntax-core`, `realtime-recovery`, `type-analysis`, `host-behavior`, `module-context`, `userform-designer`, `legacy-edges` |
| `excel_vba_analysis_additional_edge_cases.md` | Additional parser, type, host, class, conditional compilation, API, file I/O, and tokenizer cases | `syntax-core`, `type-analysis`, `object-member`, `host-behavior`, `legacy-edges`, `limits-boundaries`, `tokenizer` |
| `excel_vba_analysis_final_hardening_cases.md` | Final hardening layer for module-kind, directives, signatures, assignments, properties, labels, arrays, types, Implements, host warnings, and realtime partial states | `syntax-core`, `project-binder`, `type-analysis`, `object-member`, `host-behavior`, `realtime-recovery`, `module-context` |
| `excel_vba_analysis_limits_boundary_cases.md` | Generated and hand-written limits for continuations, line length, strings, names, modules, argument count, arrays, and host limits | `limits-boundaries`, `syntax-core`, `realtime-recovery`, `host-behavior` |
| `xlide_vba_legacy_visible_corpus_edges.md` | Legacy-visible VBA grammar and recovery cases | `legacy-edges`, `syntax-core`, `realtime-recovery`, `type-analysis` |
| `xlide_vba_visible_analysis_corpus_recommendations.md` | Recommended user-visible fixture organization and diagnostic range strategy | `diagnostic-ranges`, `realtime-recovery`, `module-context`, `host-behavior`, `roundtrip-io` |
| `xlide_vba_realtime_analysis_final_corpus_addendum.md` | Latest hardening addendum for Excel syntax traps, legacy control transfer, completion contexts, canaries, UserForm symbols, and casing | `host-behavior`, `legacy-edges`, `completion-context`, `canary-verdicts`, `userform-designer`, `casing` |

## Category Index

### `syntax-core`

Purpose: grammar, tokenization, declaration placement, block structure, call
shape, directives, procedure signatures, literals, labels, preprocessor syntax,
and parser traps.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: `MOD_*`, `LEX_*`, `CONT_*`,
  `LIT_*`, `PROC_*`, `PROP_*`, `CTRL_*`, `PP_*`, `API_*`, `TRAP_*`, `BAD_*`.
- `excel_vba_analysis_additional_edge_cases.md`: `OPT_*`, `ATTR_*`, `DEF_*`,
  `PROC_009` through `PROC_018`, `CALL_*`, `ID_*`, `SEP_*`, `CASE_*`,
  `PP_*`, `API_*`, `FILE_*`, `TOK_*`.
- `excel_vba_analysis_final_hardening_cases.md`: `DIRECTIVE_*`, `SIG_*`,
  `LABEL_*`, `COND_*`, and syntax-shaped pieces of `TYPE_*`.
- `excel_vba_analysis_limits_boundary_cases.md`: line continuation,
  identifier/name, module-name, and argument-count boundary cases.
- `xlide_vba_legacy_visible_corpus_edges.md`: line labels, `DefType`, file I/O,
  `Option Compare`, `AddressOf`, date and `#` ambiguity, colon separators.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: `EXCEL_SYNTAX_*` and
  `CANARY_*` grammar-sensitive cases.

Promotion path:

- Use MS-VBAL first for pure grammar.
- Use focused compile-oracle cases when VBE behavior is host/compiler-specific
  or corpus language says "invalid or warning", "host-specific", or
  "mode-sensitive".

Near-term candidates:

- `CANARY_001` procedure declaration missing parentheses.
- `CANARY_002` duplicate `Option Explicit`.
- `CANARY_003` and `CANARY_004` type suffix plus `As` clause.
- `CANARY_005` descending `DefType` range.
- `TRAP_003` and `TRAP_004` `Call` keyword parentheses behavior.

### `realtime-recovery`

Purpose: incomplete code while typing should produce stable, local diagnostics
or no diagnostic; it must not flood the editor with cascades.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: `WITH_003`, incremental
  `RT_001` through `RT_004`, and high-value `BAD_*` recovery snippets.
- `excel_vba_analysis_additional_edge_cases.md`: `RT_005` through `RT_009`.
- `excel_vba_analysis_final_hardening_cases.md`: `RT_001` through `RT_007`.
- `excel_vba_analysis_limits_boundary_cases.md`: `CONT_LIMIT_007` and realtime
  acceptance rules.
- `26_class_and_userform_deep_edges.md`: `RT_CLASS_*` and `RT_FORM_*`.
- `27_semantic_runtime_resolution_edges.md`: `RT_SEM_*`.
- `xlide_vba_legacy_visible_corpus_edges.md`: `negative_recovery_fuzz_seed.md`.
- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_realtime_incomplete_states.md`.

Promotion path:

- Unit tests should drive these first because the behavior is XLIDE editor
  policy.
- Use oracle only when the completed form of the same code becomes a new
  syntax or runtime rule.

Near-term candidates:

- Remaining incomplete member-access states beyond the promoted oracle-backed
  `receiver.` and active-`With` bare-dot cases.
- Incomplete named argument while typing.
- Unterminated string range stability.
- Missing block closer with recovery below the broken block.
- Partial class property, event declaration, `Implements`, and UserForm handler
  states.

### `diagnostic-ranges`

Purpose: keep squiggles precise and stable; a correct diagnostic on the wrong
range is still a poor editor result.

Sources:

- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_diagnostic_ranges.md`.
- `excel_vba_realtime_analysis_test_corpus.md`: invalid and incomplete cases
  whose expected range should be narrowed.
- `excel_vba_analysis_final_hardening_cases.md`: assignment, signature, label,
  With, and realtime hardening cases.

Promotion path:

- Unit tests with explicit span markers.
- No oracle needed unless the diagnostic itself depends on VBE behavior.

Near-term candidates:

- Bad line continuation range on `_`.
- Invalid identifier start range on the identifier.
- `Set` used with scalar range on the `Set` statement.
- Missing required argument range on the call site or empty slot.

### `type-analysis`

Purpose: scalar compatibility, expression typing, procedure signatures,
assignment compatibility, optional arguments, `ByRef`, `ParamArray`, arrays,
UDTs, enums, and return assignment.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: declarations, procedures,
  properties, expressions, arrays, UDTs, enums, and invalid corpus cases.
- `excel_vba_analysis_additional_edge_cases.md`: procedure signatures,
  `IsMissing`, `ByRef`, `Set`, `Is`, arrays, UDTs, enums, properties, and
  For Each control variables.
- `excel_vba_analysis_final_hardening_cases.md`: signature, assignment,
  property consistency, arrays, types, enums, and object-module restrictions.
- `27_semantic_runtime_resolution_edges.md`: `COERCE_*`, `FOREACH_*`,
  `API_VIS_*`, `WITHEVENTS_TYPE_*`, and conditional semantic cases.
- `excel_vba_analysis_limits_boundary_cases.md`: fixed-length strings, argument
  limits, array dimensions, naming boundaries.
- `xlide_vba_legacy_visible_corpus_edges.md`: `Let`, `Set`, `LSet`, `RSet`,
  `Mid`, optional arguments, `ByRef`, `ParamArray`, date literals.

Promotion path:

- Use `docs/type_analysis_corpus_coverage.md` as the detailed type matrix.
- Add valid, invalid, and unknown/no-diagnostic unit tests for each rule.
- Use runtime oracle for deterministic runtime-error claims.
- Use compile oracle for signature and declaration forms where MS-VBAL is not
  enough or Excel host context matters.

Near-term candidates:

- `ByRef` exactness and parenthesized expression behavior. First scalar compile
  slice is promoted through `byref-argument-type-mismatch`; continue with
  object references, arrays, Variant behavior, named arguments, and runtime
  mutation behavior.
- Optional/default parameter validity.
- `ParamArray` placement and element typing.
- Return assignment inside `Function` and `Property Get`.
- Object assignment without `Set` where receiver type is statically known.

### `runtime-resolution`

Purpose: deterministic runtime faults, runtime-error suppression, reachability,
and runtime canaries whose result may justify a red deterministic-runtime
diagnostic only after focused evidence.

Sources:

- `27_semantic_runtime_resolution_edges.md`: `RUNTIME_*`, `ERROR_FLOW_*`,
  `COERCE_*`, `COND_SEM_*`, and `RT_SEM_*`.
- `docs/vba_runtime_type_mismatch_oracle_matrix.md`: existing verified runtime
  type-mismatch matrix for nonnumeric string coercion.
- `syntax_corpus/oracle/vbe_oracle_cases.json`: promoted runtime oracle
  fixtures.

Promotion path:

- Add observe-only runtime oracle fixtures before changing analyzer behavior.
- Record compile acceptance separately from runtime failure.
- Red diagnostics require deterministic proof that the statement will raise if
  reached and must use authoritative language.
- Suppressed, unreachable, host-dependent, locale-dependent, or value-dependent
  faults are no-diagnostic or yellow only unless the model proves otherwise.
- Corpus severity examples are raw material; XLIDE severity follows
  `docs/roadmap_version_2.x.md`.

Near-term candidates:

- Division, integer division, and `Mod` by literal zero.
- Negative literal arguments to curated runtime functions such as `Left$` and
  `String$`.
- `On Error Resume Next` suppression boundaries.
- Conditional reachability for deterministic runtime faults.
- `Null`, `Empty`, and plus/ampersand coercion controls.

### `project-binder`

Purpose: workbook-aware symbol resolution: public procedures, visibility,
module-qualified calls, duplicate public names, shadowing, classes, UDTs,
enums, and module order stability.

Sources:

- `excel_vba_analysis_final_hardening_cases.md`: `EXPLICIT_*`, `SYMBOL_*`,
  `TYPE_*`, `IMPL_*`, module-kind-sensitive validation.
- `excel_vba_analysis_additional_edge_cases.md`: class, interface, event,
  WithEvents, UDT, enum, property, and Excel module-context examples.
- `26_class_and_userform_deep_edges.md`: `FRIEND_*`, `OBJECT_MEMBER_*`,
  `EVENT_*`, `WITHEVENTS_*`, `IMPLEMENTS_*`, `PROP_CLASS_*`, and default
  instance cases.
- `27_semantic_runtime_resolution_edges.md`: `REF_*`, `API_VIS_*`,
  `MACRO_VIS_*`, `PREDECLARED_*`, and project-visible reference-resolution
  cases.
- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_scope_and_shadowing.md` and `vbe_module_context_events.md`.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`:
  module-context and completion cases that need project metadata.

Promotion path:

- Build deterministic multi-module fixtures first.
- Use oracle only when a new binder rule claims VBE compile behavior.
- Prefer no diagnostic for ambiguous duplicate exported names until shadowing
  rules are explicitly modeled and tested.

Near-term candidates:

- Module-qualified public procedure calls.
- Public/private procedure visibility.
- Duplicate exported procedure ambiguity.
- Project classes, UDTs, and enums as valid `As` type names.
- Module order invariance.

### `module-context`

Purpose: module-kind-aware behavior for standard modules, class modules,
worksheet modules, ThisWorkbook, and UserForms. This includes lifecycle
procedures, `Me`, `Friend`, events, and declarations that are valid only in
some module kinds.

Sources:

- `26_class_and_userform_deep_edges.md`: `CLASS_LIFE_*`, `FORM_LIFE_*`,
  `ME_*`, `FRIEND_*`, class/UserForm event and module-kind cases.
- `27_semantic_runtime_resolution_edges.md`: `HOST_EVENT_SIG_*`,
  `MACRO_VIS_*`, `PREDECLARED_*`, and module-kind pieces of `WITHEVENTS_TYPE_*`.
- `excel_vba_analysis_final_hardening_cases.md`: module-kind-sensitive
  validation, object-module restrictions, Implements, host event stubs.
- `excel_vba_realtime_analysis_test_corpus.md`: worksheet, workbook, class,
  event, WithEvents, and UserForm baseline cases.
- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_module_context_events.md`.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: worksheet/UserForm
  completion and module-context examples.

Promotion path:

- Add fixture metadata for `moduleKind`, `moduleName`, and host-generated
  symbols.
- Use compile oracle for module-kind restrictions that are not purely parser
  facts.
- Keep module-kind warnings yellow unless VBE rejection is verified.

Near-term candidates:

- `Me` in standard modules versus object modules.
- `Friend` procedure validity by module kind.
- Class/UserForm lifecycle procedure signatures.
- Event and WithEvents declarations by module kind.

### `object-member`

Purpose: object/reference typing, `Set` required/forbidden, default members,
member existence, class/document module members, and known receiver chains.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: object declaration and property
  examples.
- `excel_vba_analysis_additional_edge_cases.md`: `EXPR_010` through `EXPR_015`,
  `PROP_006` through `PROP_009`, class and WithEvents edge cases.
- `excel_vba_analysis_final_hardening_cases.md`: `ASSIGN_*`, `PROP_*`,
  `WITH_*`, `IMPL_*`.
- `26_class_and_userform_deep_edges.md`: object-module members, events,
  WithEvents declarations, Implements, property consistency, UserForm usage, and
  default instances.
- `27_semantic_runtime_resolution_edges.md`: `RUNTIME_007`, `WITHEVENTS_TYPE_*`,
  `PREDECLARED_*`, and object/event binding cases.
- `xlide_vba_visible_analysis_corpus_recommendations.md`: host object patterns
  and module-context event cases.

Promotion path:

- Model concrete receiver types first.
- Add `member-not-found`, `set-required`, and `set-forbidden` only when the
  receiver/target type is known.
- Use Excel object metadata only from curated, audited sources.

Near-term candidates:

- `Set` assignment to known object variables.
- Object assignment without `Set` when the target type is known object.
- Property Get returning an object without `Set`.
- Leading dot outside `With` versus inside nested `With`.

### `host-behavior`

Purpose: Excel/VBA host syntax and semantic behavior that is not generic VBA
grammar: `Range`, `Cells`, `Evaluate`, `Application.Run`, worksheet/workbook
events, formula strings, object-model receiver chains, and host warnings.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: `XL_*`.
- `excel_vba_analysis_additional_edge_cases.md`: `XL_009` through `XL_020`.
- `excel_vba_analysis_final_hardening_cases.md`: `HOST_*`.
- `excel_vba_analysis_limits_boundary_cases.md`: `EXCEL_LIMIT_*`.
- `27_semantic_runtime_resolution_edges.md`: `HOST_EVENT_SIG_*`,
  `MACRO_VIS_*`, reference-sensitive host binding, and Application event cases.
- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_excel_host_common_patterns.md`.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`:
  `EXCEL_SYNTAX_*`.

Promotion path:

- Keep host warnings yellow unless a compile/runtime error is proven.
- Use live Excel/VBE oracle for host/compiler verdicts.
- Use curated Excel object-model metadata for receiver/member facts.

Near-term candidates:

- `[A1]` bracket evaluate shorthand versus bracketed identifiers.
- Bang operator forms: `rs!Name` and `rs![Order Date]`.
- `TypeOf ... Is ...` parsing and malformed `TypeOf`.
- Unqualified `Range` as optional yellow host guidance only.

### `completion-context`

Purpose: preserve enough parser and binder state to answer completions,
signature help, and hover while the user is mid-expression.

Sources:

- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: `COMP_001` through
  `COMP_010`.
- `xlide_vba_visible_analysis_corpus_recommendations.md`: recommended fixture
  marker shape for incomplete states and diagnostic ranges.
- `excel_vba_analysis_additional_edge_cases.md`: `RT_006`, `RT_009`.
- `26_class_and_userform_deep_edges.md`: `COMP_CLASS_*` and `COMP_FORM_*`.

Promotion path:

- Editor/provider tests with cursor markers.
- No red diagnostics for incomplete completion contexts unless the completed
  syntax is independently verified invalid.

Near-term candidates:

- Worksheet variable member completion.
- `ThisWorkbook.` completion.
- Nested `With` leading-dot completion.
- Named argument completion.
- Class `Me` completion and UserForm designer-control completion.

### `userform-designer`

Purpose: symbols that are valid because of UserForm designer metadata or
host-generated module context, not because they appear in the code pane.

Sources:

- `excel_vba_realtime_analysis_test_corpus.md`: `FORM_*`, worksheet/workbook
  event cases, class events and WithEvents.
- `excel_vba_analysis_additional_edge_cases.md`: UserForm, class, event, and
  WithEvents edge cases.
- `excel_vba_analysis_final_hardening_cases.md`: module-kind-sensitive
  validation and event stubs.
- `26_class_and_userform_deep_edges.md`: `FORM_LIFE_*`, `FORM_SYMBOL_*`,
  `FORM_USE_*`, UserForm default instance cases, and UserForm completion.
- `27_semantic_runtime_resolution_edges.md`: UserForm default-instance and
  predeclared-instance cases under `PREDECLARED_*`.
- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `vbe_module_context_events.md`.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`:
  `FORM_SYMBOL_*`.

Promotion path:

- Add fixture metadata for module kind and designer symbols.
- Do not mark unknown controls red without designer metadata proving the symbol
  set.
- Keep code-only parse ambiguous cases silent or advisory.

Near-term candidates:

- Known UserForm control symbol.
- `Me.` in worksheet and UserForm modules.
- UserForm Initialize event.
- Standard-module event stubs as no syntax error, optional semantic warning.
- UserForm default instance, `Load`, `Show`, `Hide`, and `Unload` patterns.

### `limits-boundaries`

Purpose: generated and boundary tests for language limits and host limits.

Sources:

- `excel_vba_analysis_limits_boundary_cases.md`: continuation count, physical
  and logical line lengths, string and fixed-string sizes, identifier length,
  module-name length, argument limits, array dimension limits, Excel formula
  and worksheet-function limits.
- `excel_vba_analysis_additional_edge_cases.md`: Declare/PtrSafe/LongPtr,
  LongLong, API declarations, array and ParamArray boundaries.
- `excel_vba_realtime_analysis_test_corpus.md`: Declare and 64-bit
  compatibility baseline.

Promotion path:

- Generated tests should be deterministic and checked into fixtures or created
  by deterministic builders.
- XLIDE-owned deterministic fixture builders live in
  `tests/helpers/vbaLimitBoundaryFixtures.ts`; use them before promoting
  boundary cases into analyzer diagnostics.
- Oracle only for actual VBE limit verdicts.
- Yellow warnings for platform or host limits unless a hard failure is proven.

Near-term candidates:

- 60 argument boundary.
- 60 array dimensions boundary.
- Fixed-length string nonliteral length expressions and assignment/truncation
  behavior.
- Continuation-count limit.

### `legacy-edges`

Purpose: valid legacy VBA should not be rejected just because it is unusual.
Legacy constructs may receive yellow style guidance only when policy explicitly
allows it.

Sources:

- `xlide_vba_legacy_visible_corpus_edges.md`: line numbers, `DefType`, `Let`,
  `LSet`, `RSet`, `Mid`, file I/O, optional/ByRef/ParamArray, `Option Compare`,
  `AddressOf`, date literals, colon separators, negative recovery fuzz seeds.
- `excel_vba_realtime_analysis_test_corpus.md`: `ERR_*`, `TRAP_*`, legacy
  control flow and parser traps.
- `excel_vba_analysis_additional_edge_cases.md`: `ERR_005` through `ERR_010`,
  file I/O and legacy branching.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`:
  `LEGACY_TRANSFER_*`.

Promotion path:

- Compile oracle for debated legacy syntax.
- Runtime oracle only when the proposed diagnostic is about deterministic
  runtime behavior.
- Yellow style warnings require explicit XLIDE policy; never convert legacy
  discomfort into red syntax errors.

Near-term candidates:

- `GoSub`/`Return` valid forms.
- `Return` without `GoSub` context.
- `On expression GoTo` and `On expression GoSub`.
- `On Error GoTo -1`.

### `tokenizer`

Purpose: highlighting and lexing traps that should not leak into diagnostics.

Sources:

- `excel_vba_analysis_additional_edge_cases.md`: `TOK_001` through `TOK_006`.
- `excel_vba_realtime_analysis_test_corpus.md`: strings, comments, date
  literals, type suffixes, keyword-like identifiers.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: casing and bracketed
  identifier preservation.

Promotion path:

- Lexer/tokenizer unit tests first.
- No oracle needed unless tokenization affects a hard VBE verdict.

Near-term candidates:

- Apostrophe inside string is not comment.
- `Rem` inside identifier is not comment.
- Date literal versus preprocessor hash.
- Bracketed identifiers and type-declaration suffix preservation.

### `roundtrip-io`

Purpose: hidden/exported metadata, attributes, and internal IO should be tested
without surfacing non-user-editable content as live editor diagnostics.

Sources:

- `xlide_vba_visible_analysis_corpus_recommendations.md`:
  `internal-io-only/roundtrip_preservation.md`.
- `26_class_and_userform_deep_edges.md`: `EXPORT_META_*`.
- `excel_vba_realtime_analysis_test_corpus.md`: exported `Attribute` examples.
- `excel_vba_analysis_additional_edge_cases.md`: `ATTR_*`.
- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: `CANARY_006` and
  `CANARY_007`.

Promotion path:

- IO roundtrip tests and export/import fixture tests.
- Live editor diagnostics should treat user-typed attributes separately from
  exported metadata.

Near-term candidates:

- Exported module/member attributes before code body.
- User-typed `Attribute` inside live editor body.
- Attribute in exported member metadata position.

### `canary-verdicts`

Purpose: settle ambiguous or corpus-disputed behavior with small focused
fixtures before implementing analyzer behavior.

Sources:

- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: `vbe_canary_verdicts.md`.
- `26_class_and_userform_deep_edges.md`: suggested VBE canary matrix for class
  and UserForm module-kind behavior.
- `27_semantic_runtime_resolution_edges.md`: VBE canary matrix for runtime,
  coercion, event binding, references, macro discoverability, and predeclared
  instance behavior.
- Debated cases pulled from any pending Markdown source.
- Existing `syntax_corpus/oracle/vbe_oracle_cases.json` promoted fixtures.

Promotion path:

- Add observe-only oracle fixture first.
- Run only the focused case.
- Promote to accepted/rejected only after the oracle returns a stable verdict.
- Update `diagnostic_influence_audit.json` only if the case drives an active
  diagnostic.

Near-term candidates:

- The canary list under `CANARY_*`.
- Any future "should this error?" discussion where VBA behavior is not already
  proven.
- Type runtime probes that decide red deterministic-runtime diagnostics.

### `casing`

Purpose: canonical casing, symbol casing propagation, bracketed identifier
preservation, type-declaration suffix preservation, and host member casing.

Sources:

- `xlide_vba_realtime_analysis_final_corpus_addendum.md`: `CASING_001` through
  `CASING_006`.
- `excel_vba_analysis_additional_edge_cases.md`: keyword, identifier, and type
  suffix traps.
- `excel_vba_realtime_analysis_test_corpus.md`: keyword-like identifiers and
  type suffix examples.

Promotion path:

- Formatter/completion tests, not analyzer diagnostics, unless a casing rule is
  explicitly introduced as XLIDE guidance.
- Preserve strings, comments, bracketed identifiers, and suffixes.

Near-term candidates:

- Canonical keyword casing.
- Declaration casing propagated to references.
- Preserve string literal and comment content.
- Host member casing from metadata.

## Promotion Queue

1. `canary-verdicts`: promote the small `CANARY_*` compile cases first. These
   are cheap, high-signal, and prevent wrong assumptions.
2. `project-binder`: continue the public-procedure slice with
   module-qualified calls, visibility, duplicate public names, and order
   invariance.
3. `module-context`: add fixture metadata for module kind before class,
   UserForm, event, or `Me` diagnostics.
4. `type-analysis`: add a `ByRef` matrix and optional/default parameter matrix
   before broad unknown-type diagnostics.
5. `realtime-recovery`: add marker-based incomplete-state tests for trailing
   dots, incomplete calls, line continuation, and unterminated strings.
6. `host-behavior`: verify Excel bracket shorthand, bang operator, and
   `TypeOf` traps before adding parser or host diagnostics.
7. `completion-context`: add cursor-marker fixtures for member completion and
   named-argument completion.
8. `userform-designer`: introduce fixture metadata for designer symbols before
   any unknown-control warning.
9. `limits-boundaries`: add deterministic fixture builders for generated limit
   cases.
10. `legacy-edges`: use the oracle to protect valid legacy syntax from false
   red diagnostics.

## Stop Rules

- If behavior is disputed, add an observe-only oracle fixture before changing
  analyzer behavior.
- If a rule depends on project symbols that the binder cannot see yet, keep the
  case in `project-binder` and emit no hard diagnostic.
- If a case is an XLIDE style or host guidance preference, it is yellow at most.
- If a corpus case says "invalid or warning", "host-specific",
  "mode-sensitive", or "verify", it is not diagnostic-driving evidence.
