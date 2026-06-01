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
| Scalar assignment coercion | Partial | Runtime oracle matrix for numeric, Boolean, String controls | Date, Currency edge formats, Decimal, Byte bounds, overflow vs type mismatch | Focused runtime oracle cases |
| Scalar argument coercion | Partial | Currency parameter, native `Left` length, numeric-string controls | Boolean arguments, Date arguments, more native signatures, ByRef arguments | Runtime oracle and curated metadata tests |
| Nonnumeric string arithmetic | Verified for current slice | Runtime oracle `1 + "string"`, reversed operands, grouping, multiplication, string-plus-string `+` valid control | Division, integer division, exponentiation, `Mod`, comparisons, variables with constant values | Runtime oracle for each new operator family |
| String concatenation operators | Partial | Runtime-valid controls `1 & "string"` and `"string" + "string"` plus inference tests | Object operands, arrays, `Null`, `Empty`, unknown/Variant behavior, variables with known String type under `+` | Runtime oracle where behavior affects diagnostics |
| Unknown and `Variant` suppression | Partial | Unit tests for arguments, assignments, arithmetic, concatenation; oracle-backed `Option Explicit` bare assignment target controls, including missing-`Option Explicit` implicit Variant assignment | `Empty`, `Null`, `Nothing`, broad read-reference undeclared-variable cases, indexed assignment targets, built-in constants/globals needed before broader identifier scanning | Unit tests plus oracle if runtime behavior drives red/yellow policy |
| Object argument mismatch | Partial | Oracle `string_literal_to_object_argument`, same-module return tests | Passing `Nothing`, object variables, class instances, host objects | Compile oracle and project binder tests |
| `Set` object assignment rules | Partial | Oracle `set_scalar_integer_assignment`, unit tests for missing `Set` on known object variables, Function/Property Get return names, and source-backed object members, `Set` used against scalar variables/members, incompatible project object RHS types, scalar RHS object assignment, and explicit `Implements` compatibility | Host object RHS compatibility matrix, non-project interface edge cases, broader Property Let/Set declaration consistency | Compile oracle and binder tests |
| Scalar member access | Verified for current slice | Oracle `string_scalar_named_member_compile`, `integer_scalar_named_member_compile`, trailing-dot compile controls, unit tests | Array element receivers, fixed-length strings, parenthesized scalar expressions, default properties | Focused oracle for new receiver forms, unit tests for declared-variable binder |
| Invalid `As` type names | Partial | Oracle `As Int`; `ProjectIndex.visibleTypeNames()` for project classes/document/userform modules and visible UDT/Enum names; broad unknown deferred | Host/reference type catalog, known non-type declarations in type position, misspellings, ambiguity handling | Project binder tests plus compile oracle for disputed names |
| Broad numeric family | Pending | Some `Integer`, `Double`, `Currency`, `Long` examples | Byte, LongLong, LongPtr, Single, Decimal, fixed-width overflow, type-declaration suffixes | Runtime oracle for coercion/overflow; parser tests for suffixes |
| Boolean compatibility | Partial | Boolean assignment true/invalid-string runtime cases | Boolean parameters, numeric-to-Boolean, comparisons, `And`/`Or`/`Not` operands | Runtime oracle for coercion, parser/unit tests for operators |
| Date compatibility | Missing | Legacy corpus mentions date literals | String-to-Date, numeric-to-Date, date parameters, locale-sensitive strings | Needs oracle/spec before diagnostics |
| String family | Pending | Fixed-length string and string-size legacy corpus cases | Fixed-length assignment truncation, invalid fixed lengths, type suffix `$` | Spec and oracle mix |
| Optional/default parameter typing | Pending | Arity tests and optional argument oracle seed | Default value type compatibility, missing optional argument slots, named optional args | Compile oracle for headers and call syntax |
| ParamArray typing | Pending | Legacy corpus examples, arity tests | Non-Variant ParamArray, ParamArray not last, argument element inference | Spec first, oracle for VBE messages |
| ByRef compatibility | Missing | None systematic | ByRef exactness vs coercion, literals to ByRef, parenthesized ByRef expressions | Compile/runtime oracle required |
| Named arguments | Partial | Unit tests for named argument mapping | Named/positional mixing, duplicate named args, optional named omissions | Compile oracle for call syntax; unit tests for binder mapping |
| Return assignment inside Function | Partial | Unit tests for scalar Function return assignment, object Function return requiring `Set`, incompatible object return assignment, compatible object return assignment, and Property Get object return assignment | Property Let/Set declaration consistency, object return values through default members, host object return compatibility | Compile/runtime oracle and binder tests |
| Comparisons | Missing | Roadmap only | Numeric/string/Date/Object comparisons, `Like`, `Is`, `Is Not` | Needs operator matrix and oracle for edge behavior |
| Arrays | Pending | Legacy corpus mentions arrays and `Array()` | Dynamic/fixed arrays, indexed element type, array parameter compatibility, `ParamArray` elements | Parser/binder tests and oracle for compile errors |
| Enums | Pending | Legacy corpus mentions enums | Enum declaration, enum member values, enum-to-integer compatibility, unknown enum names | Spec plus binder tests |
| UDTs | Pending | Legacy corpus mentions UDTs | UDT declarations, member access, object members in UDTs, cross-module UDT names | Spec plus compile oracle for edge cases |
| Classes and document modules | Partial | `ProjectIndex.visibleTypeNames()`, `ProjectIndex.projectClassMembers()`, member-completion tests for `Dim p As Person: p.` and `Me.`, source-backed object member go-to-definition including `Me.Member`, inline-doc completion/hover tests for `p.Age`, source-backed class member signature help with inline summary/parameter docs, source-backed member-call arity/type diagnostics, property and public-field assignment diagnostics for typed writable members and read-only properties, `member-not-found` diagnostics for source-backed class receivers, oracle coverage for public fields as writable object members, `Public Const` exclusion from object-member surfaces, module-kind-sensitive public object-module declaration diagnostics, late-bound `Object`/`Variant` unknown-member compile controls and no-diagnostic tests even after `Set` assignment, exported `VB_UserMemId = 0` default-member attribute extraction with source spans, shared type-name semantic-token/hover tests for project names in `As Person` and `New Person`, workbook/worksheet document-module event-handler completion with scoped insertions and duplicate/inside-procedure suppression, non-red wrong-module event-handler diagnostics for known workbook/worksheet handlers, optional bridge `documentType` for workbook/worksheet/chart modules, class/module docs from module-header comments in type completion/hover, module-level `Implements` metadata for object compatibility, `Property Set`/object-valued public field assignment diagnostics for missing `Set` and incompatible object RHS types | Class instances beyond current member-assignment slice, default-member expression semantics such as `textValue = p`, chart and UserForm/designer-backed event handler authoring, declared `Event` members, `WithEvents` bindings, Friend visibility, document/UserForm designer members, broader member-call diagnostics | Project fixture builder and binder tests; oracle only when VBA behavior affects diagnostics |
| Excel object model receiver chains | Partial | Legacy host-pattern corpus, oracle `thisworkbook_unknown_member_compile`, `workbook_event_member_call_compile`, `implicit_member_call_parentheses_compile`, `standalone_zero_arg_method_call_compile`, `standalone_range_property_empty_compile`, and non-empty member/property call controls, generated member surfaces for `Application`/`Workbook`/`Worksheet`/`Range`/`Workbooks`/`Worksheets`/`Sheets`, generated host member signatures/docs in completion, hover, signature help, and member-call arity/type diagnostics, host type-name hover and semantic coloring in `As Worksheet`, exhaustive host diagnostics for `Excel.Workbook` and `Excel.Worksheet`, tests for `ThisWorkbook.doesnotexist`, `ThisWorkbook.AfterSave` as a non-member event call, `ActiveSheet.asdf`, declared `Worksheet` missing members, `Worksheets(index).Range`, merged worksheet/chart `Sheets(index).Range`, completion-only generic `Object`/`Variant` refinement from simple `Set` assignments, generated non-exhaustive members, object-member event exclusion, and empty-parentheses standalone member-call syntax, plus `docs/excel_reference_coverage.md` | Broader Excel object surfaces, property-vs-method call semantics beyond collection-default `Item`, consistent object-access diagnostics across host/source/external metadata, promotion of additional dump-backed host types to exhaustive only after controls | Generated/reference dump metadata plus integration tests; hard diagnostics only after exhaustive provenance or oracle-backed controls |
| Native VBA runtime metadata breadth | Partial | Curated metadata, `Left` coercion oracle, runtime argument-count coverage for explicit parameter-list signatures such as missing `MsgBox` Prompt, and primitive type-name hover/semantic coloring in `As Currency` without treating type names as runtime calls | More functions with typed params and returns, statements vs functions, shadowing, runtime statements whose signatures are not parenthesized | Primary docs for signatures, oracle for behavior disputes |
| External `.vbref.xml` metadata types | Partial | Docs type-hint strategy exists | Metadata-driven callable signatures, precedence vs source, stale metadata behavior | Unit tests, no oracle unless VBA behavior is involved |
| Inline doc-comment type hints | Partial | Docs and language-service tests | Type hints as signature gap-fillers, no override of concrete source types | Unit tests; hard diagnostics only with deterministic source/metadata |
| Cross-module procedure calls and identifiers | Partial | Project signature tests for unique exported standard-module `Sub`/`Function` calls, module-qualified standard-module calls, ambiguous bare-name no-diagnostic controls, module-order stability, visibility-filtered bare procedure names, and `ProjectIndex.visibleIdentifierNames()` for Option Explicit assignment targets | Broader read-use variable visibility, indexed targets, shadowing, class/document member binding, workbook fixture builder | Workbook-level fixture builder plus analyzer tests |
| Realtime incomplete expressions | Missing | General diagnostics tests | Suppression while typing partial calls, partial strings, partial operators | Editor/provider tests and targeted analyzer tests |

## Immediate Corpus Backlog

These are the highest-value additions before the next major binder slice:

1. **`As` type resolver**: build on `ProjectIndex.visibleTypeNames()` to
   separate known project/host types, known non-type declarations, ambiguous
   type names, and unknown external-reference candidates before any broad
   unknown diagnostic ships.
2. **Project fixture builder**: machine-readable multi-module fixtures so
   cross-module calls, classes, UDTs, and enums can be tested deterministically.
3. **ByRef oracle matrix**: exactness, coercion, literals, parenthesized
   expressions, and object references.
4. **Date coercion matrix**: accepted literals, rejected strings, and
   locale-sensitive cases marked as no-diagnostic until deterministic.
5. **Object assignment matrix**: host object RHS compatibility, non-project
   interfaces, default-member object returns, and broader Property Get/Let/Set.
6. **Fixed-length string matrix**: verify `As String * n` declaration grammar,
   allowed module contexts, length boundaries, assignment/truncation behavior,
   interaction with member access, and type-declaration suffixes before any
   analyzer rule treats `String * n` as ordinary trailing declaration tokens.
7. **Operator matrix**: arithmetic, integer division, `Mod`, comparisons,
   `Like`, `Is`, and Boolean operators.
8. **Native metadata expansion matrix**: prioritize common VBA runtime
   functions with typed parameters and returns, one oracle probe per behavior
   family when signatures alone do not answer compatibility.
9. **Document-module event handler matrix**: extend the completed
   workbook/worksheet completion slice to chart and UserForm event signatures;
   keep testing completion/insertions, exact casing, module-type scoping,
   wrong-module guidance for newly added event surfaces, and continued exclusion
   from object-member completion/callable surfaces.
10. **Class/object member edge matrix**: verify remaining object-module
   declaration restrictions and default-member expression semantics through an
   import-capable oracle path before these cases drive hard diagnostics.
11. **Unified object-member rule matrix**: for every added curated object
    surface, record completion coverage, member-not-found policy,
    writable/read-only assignment behavior, type mismatch behavior, doc display,
    and no-diagnostic controls for incomplete metadata.
12. **Reference dump coverage matrix**: extend the generated
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
