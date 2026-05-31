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
| Same-module function return types | Partial | Nested return-type tests | Return values through variables, property returns, typed functions across modules | Unit tests first; oracle for disputed compile/runtime behavior |
| Scalar assignment coercion | Partial | Runtime oracle matrix for numeric, Boolean, String controls | Date, Currency edge formats, Decimal, Byte bounds, overflow vs type mismatch | Focused runtime oracle cases |
| Scalar argument coercion | Partial | Currency parameter, native `Left` length, numeric-string controls | Boolean arguments, Date arguments, more native signatures, ByRef arguments | Runtime oracle and curated metadata tests |
| Nonnumeric string arithmetic | Verified for current slice | Runtime oracle `1 + "string"`, reversed operands, grouping, multiplication, string-plus-string `+` valid control | Division, integer division, exponentiation, `Mod`, comparisons, variables with constant values | Runtime oracle for each new operator family |
| String concatenation operators | Partial | Runtime-valid controls `1 & "string"` and `"string" + "string"` plus inference tests | Object operands, arrays, `Null`, `Empty`, unknown/Variant behavior, variables with known String type under `+` | Runtime oracle where behavior affects diagnostics |
| Unknown and `Variant` suppression | Partial | Unit tests for arguments, assignments, arithmetic, concatenation | `Empty`, `Null`, `Nothing`, implicit Variant under missing `Option Explicit` | Unit tests plus oracle if runtime behavior drives red/yellow policy |
| Object argument mismatch | Partial | Oracle `string_literal_to_object_argument`, same-module return tests | Passing `Nothing`, object variables, class instances, host objects | Compile oracle and project binder tests |
| `Set` to scalar | Partial | Oracle `set_scalar_integer_assignment`, unit tests | `Set` required for object assignment, `Set` forbidden with non-object RHS, property Set/Let/Get behavior | Compile oracle and binder tests |
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
| Return assignment inside Function | Missing | None systematic | Function name assignment type compatibility, object return requires `Set`, Property Get return assignment | Compile/runtime oracle and binder tests |
| Comparisons | Missing | Roadmap only | Numeric/string/Date/Object comparisons, `Like`, `Is`, `Is Not` | Needs operator matrix and oracle for edge behavior |
| Arrays | Pending | Legacy corpus mentions arrays and `Array()` | Dynamic/fixed arrays, indexed element type, array parameter compatibility, `ParamArray` elements | Parser/binder tests and oracle for compile errors |
| Enums | Pending | Legacy corpus mentions enums | Enum declaration, enum member values, enum-to-integer compatibility, unknown enum names | Spec plus binder tests |
| UDTs | Pending | Legacy corpus mentions UDTs | UDT declarations, member access, object members in UDTs, cross-module UDT names | Spec plus compile oracle for edge cases |
| Classes and document modules | Partial | `ProjectIndex.visibleTypeNames()`, `ProjectIndex.projectClassMembers()`, member-completion tests for `Dim p As Person: p.`, inline-doc completion/hover tests for `p.Age`, property assignment diagnostics for typed writable and read-only properties, `member-not-found` diagnostics for source-backed class receivers, semantic-token tests for `As Person` and `New Person` | Class-level/module-level docs, class instances beyond current member-assignment slice, late-bound `Object`/`Variant` receiver behavior, public class fields/constants oracle coverage, default members and `VB_UserMemId = 0` attributes, events, Friend visibility, `Property Set` object assignment, document/UserForm designer members, signature/go-to-definition | Project fixture builder and binder tests; oracle only when VBA behavior affects diagnostics |
| Excel object model receiver chains | Pending | Legacy host-pattern corpus | `Workbook.Worksheets().Range`, member existence, property vs method calls | Curated host metadata plus integration tests |
| Native VBA runtime metadata breadth | Partial | Curated metadata and `Left` coercion oracle | More functions with typed params and returns, statements vs functions, shadowing | Primary docs for signatures, oracle for behavior disputes |
| External `.vbref.xml` metadata types | Partial | Docs type-hint strategy exists | Metadata-driven callable signatures, precedence vs source, stale metadata behavior | Unit tests, no oracle unless VBA behavior is involved |
| Inline doc-comment type hints | Partial | Docs and language-service tests | Type hints as signature gap-fillers, no override of concrete source types | Unit tests; hard diagnostics only with deterministic source/metadata |
| Cross-module procedure calls | Partial | Project signature tests for unique exported standard-module `Sub`/`Function` calls, module-qualified standard-module calls, ambiguous bare-name no-diagnostic controls, module-order stability, visibility-filtered bare procedure names | Broader module variable/type visibility, shadowing, class/document member binding, workbook fixture builder | Workbook-level fixture builder plus analyzer tests |
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
5. **Object assignment matrix**: `Set` required, `Set` forbidden, object return
   assignment, and Property Get/Let/Set.
6. **Fixed-length string matrix**: verify `As String * n` declaration grammar,
   allowed module contexts, length boundaries, assignment/truncation behavior,
   interaction with member access, and type-declaration suffixes before any
   analyzer rule treats `String * n` as ordinary trailing declaration tokens.
7. **Operator matrix**: arithmetic, integer division, `Mod`, comparisons,
   `Like`, `Is`, and Boolean operators.
8. **Native metadata expansion matrix**: prioritize common VBA runtime
   functions with typed parameters and returns, one oracle probe per behavior
   family when signatures alone do not answer compatibility.
9. **Class/object member edge matrix**: verify late-bound `Object` and
   `Variant` receivers, public class fields/constants as writable/read-only
   members, and default members declared through exported
   `VB_UserMemId = 0` attributes before these cases drive hard diagnostics.

## Adequacy Assessment

The current corpus is adequate as a discovery backlog and adequate for the
narrow Phase 1/2 type rules already implemented. It is not yet adequate as a
complete type-analysis coverage plan for the project-wide binder, object/member
typing, ByRef compatibility, arrays, enums, UDTs, classes, document modules, or
Excel object model diagnostics.

Before a new type rule ships, this matrix should have at least one valid case,
one invalid case, one unknown/no-diagnostic case, and a named verification path
for that rule family.
