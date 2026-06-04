# Type Analysis Corpus Coverage

Purpose: track the wider "not yet verified, but on the list" corpus coverage
for XLIDE's VBA type analysis work. This document is a planning matrix, not an
authority for hard diagnostics.

The existing Markdown corpus is useful discovery material. A case becomes
diagnostic-driving evidence only after it is promoted to a verified source such
as MS-VBAL/spec evidence, an asserted Excel/VBE oracle case, or deterministic
XLIDE-owned metadata with tests.

The cross-category Markdown digestion lives in
`syntax_corpus/managed_backlog.md`; this file is the narrower type-analysis
view of that backlog.

## Status Legend

- **Verified**: enough asserted evidence exists for the current rule surface.
- **Partial**: some representative cases exist, but the area is not systematic.
- **Pending**: examples exist in legacy Markdown corpus or discussion notes, but
  they are not verified.
- **Missing**: the area needs explicit examples.
- **Needs oracle/spec**: behavior is likely VBA-specific enough that we should
  verify before using it for red diagnostics.

## Coverage Matrix

| Area | Status | Current Seeds | Gaps To Add | Verification Path |
| --- | --- | --- | --- | --- |
| Same-module procedure parameter types | Verified for current slice | `tests/vbaDiagnostics.test.ts`, oracle `normal_function_call` | More procedure kinds and properties | Parser tests plus oracle only for disputed behavior |
| Same-module function return types | Partial | Nested return-type tests plus Function/Property Get return-name assignment diagnostics | Typed functions across modules, return values through alias/default-member expressions | Unit tests first; oracle for disputed compile/runtime behavior |
| Scalar assignment coercion | Partial | Runtime oracle matrix for numeric, Boolean, String controls; Byte/Integer decimal literal boundary controls now prove in-range min/max values run and out-of-range values raise Overflow | Date, Currency edge formats, Decimal, Byte/Integer suffix/hex bounds, Long/LongLong/LongPtr overflow vs type mismatch | Focused runtime oracle cases |
| Scalar argument coercion | Partial | Currency parameter, native `Left` length, numeric-string controls, Byte/Integer decimal literal boundary controls for ByVal arguments | Boolean arguments, Date arguments, more native signatures, ByRef arguments, suffix/hex/octal numeric literal bounds beyond current decimal slice | Runtime oracle and curated metadata tests |
| Runtime argument value bounds | Verified for current slice | Runtime oracle cases for negative and zero `Left`/`Left$`, `Right`/`Right$`, `String`/`String$`, and `Space`/`Space$` bounds plus `Mid`/`Mid$` Start/Length and `Replace` Start/Count bounds; analyzer diagnostics for out-of-bound integer literals, reducible integer expressions, same-module/procedure `Const` values, current-module Enum members, and visible exported standard-module `Const`/Enum members in settled procedure statements with bare-call and named-argument handling | Broader runtime-function bounds, additional named-argument edge cases, module-qualified constants, hidden-helper constant dependencies, and `On Error Resume Next`/reachability policy | Runtime oracle for each function family plus conservative unit tests |
| Division by zero | Verified for current slice | Runtime oracle cases for `/ 0`, `\ 0`, `Mod 0`, `/ 2`, hex/octal zero, nonzero hex, module/local decimal/hex/octal `Const` zero, zero/nonzero `Const` expressions, and zero/nonzero Enum member values; analyzer diagnostics for these zero divisors in settled procedure statements, including nested expressions | Flow/reachability, `On Error Resume Next`, cross-module enum constants, and broader expression-value folding | Runtime oracle for each operator family plus conservative unit tests |
| Nonnumeric string arithmetic | Verified for current slice | Runtime oracle `1 + "string"`, reversed operands, grouping, multiplication, string-plus-string `+` valid control | Division, integer division, exponentiation, `Mod`, comparisons, variables with constant values | Runtime oracle for each new operator family |
| String concatenation operators | Partial | Runtime-valid controls `1 & "string"` and `"string" + "string"` plus inference tests | Object operands, arrays, `Null`, `Empty`, unknown/Variant behavior, variables with known String type under `+` | Runtime oracle where behavior affects diagnostics |
| Unknown and `Variant` suppression | Partial | Unit tests for arguments, assignments, arithmetic, concatenation; oracle-backed `Option Explicit` assignment controls; project-backed undeclared-variable tests for RHS/call-argument reads, control-flow headers, member receivers, indexed bases, runtime constants, and generated Excel enum constants | `Empty`, `Null`, `Nothing`, external-reference constants/globals, flow-sensitive identifier binding, ambiguous external-reference behavior | Unit tests plus oracle if runtime behavior drives red/yellow policy |
| Object argument mismatch | Partial | Oracle `string_literal_to_object_argument`, same-module return tests | Passing `Nothing`, object variables, class instances, host objects | Compile oracle and project binder tests |
| `Set` object assignment rules | Partial | Oracle `set_scalar_integer_assignment`, unit tests for missing `Set` on known object variables, Function/Property Get return names, and source-backed object members, `Set` used against scalar variables/members, incompatible project object RHS types, scalar RHS object assignment, and explicit `Implements` compatibility | Host object RHS compatibility matrix, non-project interface edge cases, broader Property Let/Set declaration consistency | Compile oracle and binder tests |
| Scalar member access | Verified for current slice | Oracle `string_scalar_named_member_compile`, `integer_scalar_named_member_compile`, trailing-dot compile controls, unit tests | Array element receivers, fixed-length strings, parenthesized scalar expressions, default properties | Focused oracle for new receiver forms, unit tests for declared-variable binder |
| Invalid `As` type names | Partial | Oracle `As Int`; `ProjectIndex.visibleTypeNames()` for project classes/document/userform modules and visible UDT/Enum names; known non-type declaration diagnostics; ambiguous visible type diagnostics; broad unknown deferred | Host/reference type catalog, misspelling suggestions, external-reference ambiguity behavior | Project binder tests plus compile oracle for disputed names |
| Broad numeric family | Partial | Some `Integer`, `Double`, `Currency`, `Long` examples; oracle-backed Byte/Integer decimal literal overflow for assignments and ByVal arguments | Long/LongLong/LongPtr, Single, Decimal, suffixes, hex/octal literals, fixed-width overflow beyond Byte/Integer | Runtime oracle for coercion/overflow; parser tests for suffixes |
| Boolean compatibility | Partial | Boolean assignment true/invalid-string runtime cases | Boolean parameters, numeric-to-Boolean, comparisons, `And`/`Or`/`Not` operands | Runtime oracle for coercion, parser/unit tests for operators |
| Date compatibility | Missing | Legacy corpus mentions date literals | String-to-Date, numeric-to-Date, date parameters, locale-sensitive strings | Needs oracle/spec before diagnostics |
| String family | Partial | Fixed-length declaration parser/symbol/hover tests; oracle-backed literal bounds for `String * 1`, `String * 65526`, `String * 0`, and `String * 65527`; deterministic fixed-length `Const` expression diagnostics; runtime argument value diagnostics for `Left`/`Left$`, `Right`/`Right$`, `String`/`String$`, `Space`/`Space$`, `Mid`/`Mid$`, and `Replace`; fixed-length string and string-size legacy corpus cases | Fixed-length assignment truncation, non-deterministic length expression semantics, broad type-suffix call normalization | Spec and oracle mix |
| Optional/default parameter typing | Partial | Arity tests, optional argument oracle seed, and oracle-backed default-value slice: numeric/numeric-string `Long`, string `String`, and `True` Boolean controls accepted; nonnumeric string defaults for `Long` and `Boolean` rejected and surfaced by `parameter-default-type-mismatch` | Broader default value expressions, Date defaults, object/array invalid defaults, missing optional argument slots, named optional args | Compile oracle for headers and call syntax |
| ParamArray typing | Pending | Legacy corpus examples, arity tests | Non-Variant ParamArray, ParamArray not last, argument element inference | Spec first, oracle for VBE messages |
| ByRef compatibility | Partial | Oracle-backed compile matrix for `ByRef Long`: exact `Long` variable accepted, literal accepted, parenthesized `Long` variable expression accepted, unparenthesized `Integer` variable rejected, and matching `ByVal Long` control accepted; analyzer diagnostic for known source-backed scalar variable exactness | Object references, arrays, Variant behavior, named arguments, runtime mutation behavior | Compile/runtime oracle required |
| Named arguments | Partial | Unit tests for named argument mapping | Named/positional mixing, duplicate named args, optional named omissions | Compile oracle for call syntax; unit tests for binder mapping |
| Return assignment inside Function | Partial | Unit tests for scalar Function return assignment, object Function return requiring `Set`, incompatible object return assignment, compatible object return assignment, and Property Get object return assignment | Property Let/Set declaration consistency, object return values through default members, host object return compatibility | Compile/runtime oracle and binder tests |
| Comparisons | Missing | Roadmap only | Numeric/string/Date/Object comparisons, `Like`, `Is`, `Is Not` | Needs operator matrix and oracle for edge behavior |
| Arrays | Pending | Legacy corpus mentions arrays and `Array()` | Dynamic/fixed arrays, indexed element type, array parameter compatibility, `ParamArray` elements | Parser/binder tests and oracle for compile errors |
| Enums | Partial | Parser and symbol graph cover declarations/members; runtime oracle-backed current-module enum member values now feed deterministic integer folding for `division-by-zero` | Cross-module enum constants, enum-to-integer compatibility, unknown enum names, duplicate/ambiguous enum member behavior | Spec plus binder tests and focused runtime oracle where values drive diagnostics |
| UDTs | Pending | Legacy corpus mentions UDTs | UDT declarations, member access, object members in UDTs, cross-module UDT names | Spec plus compile oracle for edge cases |
| Classes and document modules | Partial | `ProjectIndex.visibleTypeNames()`, `ProjectIndex.projectClassMembers()`, member-completion tests for `Dim p As Person: p.` and `Me.`, source-backed object member go-to-definition including `Me.Member`, inline-doc completion/hover tests for `p.Age`, source-backed class member signature help with inline summary/parameter docs, source-backed member-call arity/type diagnostics including parenless call statements, property and public-field assignment diagnostics for typed writable members and read-only properties, `member-not-found` diagnostics for source-backed class receivers, oracle coverage for public fields as writable object members, `Public Const` exclusion from object-member surfaces, module-kind-sensitive public object-module declaration diagnostics, late-bound `Object`/`Variant` unknown-member compile controls and no-diagnostic tests even after `Set` assignment, exported `VB_UserMemId = 0` default-member attribute extraction with source spans, shared type-name semantic-token/hover tests for project names in `As Person` and `New Person`, creatable-only `As New`/expression-level `New` completion for project classes/UserForms, resolved non-creatable `New` diagnostics for primitives, host types, document modules, UDTs, and enums, workbook/worksheet document-module event-handler completion with scoped insertions and duplicate/inside-procedure suppression, non-red wrong-module event-handler diagnostics for known workbook/worksheet handlers, optional bridge `documentType` for workbook/worksheet/chart modules, class/module docs from module-header comments in type completion/hover, module-level `Implements` metadata for object compatibility, `Property Set`/object-valued public field assignment diagnostics for missing `Set` and incompatible object RHS types, shared project-analysis helper used by current-module analysis, workbook analysis, and diagnostic fixtures | Class instances beyond current member-assignment slice, default-member expression semantics such as `textValue = p`, chart and UserForm/designer-backed event handler authoring, declared `Event` members, `WithEvents` bindings, Friend visibility, document/UserForm designer members, broader member-call diagnostics, external/reference-library class creatability metadata | Machine-readable workbook fixture files and binder tests; oracle only when VBA behavior affects diagnostics |
| Excel object model receiver chains | Partial | Legacy host-pattern corpus, oracle `thisworkbook_unknown_member_compile`, `workbook_event_member_call_compile`, `implicit_member_call_parentheses_compile`, `standalone_zero_arg_method_call_compile`, `standalone_range_property_empty_compile`, and non-empty member/property call controls, generated member surfaces for `Application`/`Workbook`/`Worksheet`/`Range`/`Workbooks`/`Worksheets`/`Sheets`, generated host member signatures/docs in completion, hover, signature help, parenthesized member-call arity/type diagnostics, parenless member-call arity diagnostics, and generated Excel enum constants in completion/hover/Option Explicit suppression, host type-name hover and semantic coloring in `As Worksheet`, exhaustive host diagnostics for `Excel.Workbook` and `Excel.Worksheet`, tests for `ThisWorkbook.doesnotexist`, `ThisWorkbook.AfterSave` as a non-member event call, `ActiveSheet.asdf`, declared `Worksheet` missing members, `Worksheets(index).Range`, merged worksheet/chart `Sheets(index).Range`, completion-only generic `Object`/`Variant` refinement from simple `Set` assignments, generated non-exhaustive members, object-member event exclusion, and empty-parentheses standalone member-call syntax, plus `docs/excel_reference_coverage.md` | Broader Excel object surfaces, property-vs-method call semantics beyond collection-default `Item`, consistent object-access diagnostics across host/source/external metadata, promotion of additional dump-backed host types to exhaustive only after controls | Generated/reference dump metadata plus integration tests; hard diagnostics only after exhaustive provenance or oracle-backed controls |
| Native VBA runtime metadata breadth | Partial | Curated metadata, `Left` coercion oracle, oracle-backed argument-value bounds for `Left`/`Left$`, `Right`/`Right$`, `String`/`String$`, `Space`/`Space$`, `Mid`/`Mid$`, and `Replace`, `IIf` compile-valid and eager branch-evaluation runtime oracle cases, runtime argument-count coverage for explicit parameter-list signatures such as missing `MsgBox` Prompt, runtime constants in completion/hover/Option Explicit suppression, and primitive type-name hover/semantic coloring in `As Currency` without treating type names as runtime calls | More functions with typed params and returns, statements vs functions, shadowing, runtime statements whose signatures are not parenthesized, broad `$` intrinsic-call normalization, deterministic diagnostics for eager `IIf` branch faults only when the branch expression itself is proven fatal | Primary docs for signatures, oracle for behavior disputes |
| External `.vbref.xml` metadata types | Partial | Docs type-hint strategy exists | Metadata-driven callable signatures, precedence vs source, stale metadata behavior | Unit tests, no oracle unless VBA behavior is involved |
| Inline doc-comment type hints | Partial | Docs and language-service tests | Type hints as signature gap-fillers, no override of concrete source types | Unit tests; hard diagnostics only with deterministic source/metadata |
| Cross-module procedure calls and identifiers | Partial | Project signature tests for unique exported standard-module `Sub`/`Function` calls, module-qualified standard-module calls, ambiguous bare-name no-diagnostic controls, module-order stability, visibility-filtered bare procedure names, `ProjectIndex.visibleIdentifierNames()` for Option Explicit assignment targets, and conservative read-reference/indexed-base diagnostics | Full identifier shadowing, arbitrary-expression binding, class/document member binding, workbook fixture builder | Workbook-level fixture builder plus analyzer tests |
| Realtime incomplete expressions | Partial | Module-analysis tests suppress active dangling member access, trailing binary operators, and unmatched opening parentheses within the active colon-separated statement | Partial calls, partial strings, editor/provider integration tests, workbook-analysis settled-state controls | Editor/provider tests and targeted analyzer tests |

## Immediate Corpus Backlog

These are the highest-value additions before the next major binder slice:

1. **Workbook fixture files**: build on the shared project-analysis helper with
   machine-readable multi-module fixtures so cross-module calls, classes, UDTs,
   and enums can be tested deterministically at workbook scale.
2. **ByRef oracle matrix**: first scalar exactness slice is promoted; continue
   with object references, arrays, Variant behavior, named arguments, and
   runtime mutation behavior.
3. **Date coercion matrix**: accepted literals, rejected strings, and
   locale-sensitive cases marked as no-diagnostic until deterministic.
4. **Object assignment matrix**: host object RHS compatibility, non-project
   interfaces, default-member object returns, and broader Property Get/Let/Set.
5. **Fixed-length string matrix**: the parser/symbol graph now treats
   recognized `As String * n` declaration suffixes as explicit metadata and
   trailing-token detection consumes that suffix. Still verify allowed module
   contexts, length boundaries, assignment/truncation behavior, scalar member
   access, and type-declaration suffix interactions before adding further hard
   diagnostics.
6. **Operator matrix**: arithmetic, integer division, `Mod`, comparisons,
   `Like`, `Is`, and Boolean operators.
7. **Native metadata expansion matrix**: prioritize common VBA runtime
   functions with typed parameters and returns, one oracle probe per behavior
   family when signatures alone do not answer compatibility.
8. **Document-module event handler matrix**: extend the completed
   workbook/worksheet completion slice to chart and UserForm event signatures;
   keep testing completion/insertions, exact casing, module-type scoping,
   wrong-module guidance for newly added event surfaces, and continued exclusion
   from object-member completion/callable surfaces.
9. **Class/object member edge matrix**: verify remaining object-module
   declaration restrictions and default-member expression semantics through an
   import-capable oracle path before these cases drive hard diagnostics.
10. **Unified object-member rule matrix**: for every added curated object
    surface, record completion coverage, member-not-found policy,
    writable/read-only assignment behavior, type mismatch behavior, doc display,
    and no-diagnostic controls for incomplete metadata.
11. **Reference dump coverage matrix**: extend the generated
    `docs/excel_reference_coverage.md` approach beyond Excel as each library
    becomes relevant. The report must keep tracking which object types are
    dump-backed, which members/signatures/events/enums/default members were
    imported, which event rows are coverage-only rather than object members,
    which gaps remain, and whether the type is eligible for
    `exhaustive` hard diagnostics.

## Adequacy Assessment

The current corpus is adequate as a discovery backlog and adequate for the
narrow Phase 1/2 type rules already implemented. It is not yet adequate as a
complete type-analysis coverage plan for the project-wide binder, object/member
typing, ByRef compatibility, arrays, enums, UDTs, classes, document modules, or
Excel object model diagnostics.

Before a new type rule ships, this matrix should have at least one valid case,
one invalid case, one unknown/no-diagnostic case, and a named verification path
for that rule family.
