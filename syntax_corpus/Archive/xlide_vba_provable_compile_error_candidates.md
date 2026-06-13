# XLIDE VBA Provable Compile-Error Candidates

**Status:** Oracle-evaluated 2026-06-13 against the live Excel/VBE oracle (Excel
16.0). Each `PCEC_*` case below now carries a recorded VBE verdict, and its
matching oracle case in `syntax_corpus/oracle/vbe_oracle_cases.json` is now
`vbe-oracle-verified` (was observe-only).

**Key finding:** the oracle **refuted 3 of the 8 candidates** — `ByVal
ParamArray`, an empty `Enum`, and `Const x As Object = Nothing` all *compile
cleanly* in VBE. Shipping rules for those from the Markdown expectation alone
would have produced false positives. This is the no-false-positive rule working
as intended: ground every red diagnostic in evidence, not plausibility.

**Purpose:** capture hard, deterministic VBA compile errors decidable from a
single declaration, signature, call, or module header — no expression binder,
host metadata, or runtime evidence required. Originally discovery material; now
reconciled with real VBE behavior.

---

## Disposition summary (oracle verdicts, 2026-06-13)

| Case | Snippet | VBE verdict | Disposition |
| --- | --- | --- | --- |
| PCEC_001 | duplicate parameter name | rejected — "Duplicate declaration in current scope" | **Already covered** by `duplicate-declaration` |
| PCEC_002 | `ByVal ParamArray` | **accepted** | **Refuted** — no rule (would be a false positive) |
| PCEC_003 | empty `Enum` | **accepted** | **Refuted** — no rule |
| PCEC_004 | empty `Type` | rejected — "User-defined type without members not allowed" | **Shipped** as `empty-type` |
| PCEC_005 | `Const x As Object = Nothing` | **accepted** | **Refuted** — no rule |
| PCEC_006 | duplicate `Option Explicit` | rejected — "Duplicate Option statement" | **Shipped** as `duplicate-option` |
| PCEC_007 | duplicate `Case Else` | rejected — "Case without Select Case" | **Shipped** as `duplicate-case-else` |
| PCEC_008 | positional after named arg | rejected — "Syntax error" | **Confirmed** — binder-gated |

Note the asymmetry the oracle revealed: an empty `Type` is rejected, but an empty
`Enum` is accepted.

---

## Confirmed compile errors — shipped as rules

These three were oracle-confirmed and have since shipped as `vbe-oracle-verified`
diagnostics (`empty-type`, `duplicate-option`, `duplicate-case-else`), each with
tests and an asserted oracle case in `diagnostic_influence_audit.json`.

### PCEC_004 Empty Type block

```vba
Public Type EmptyType
End Type
```

**Oracle verdict (2026-06-13):** rejected — "User-defined type without members
not allowed" (`empty_type_block_compile`, compile-error). A user-defined `Type`
declaring no members is a compile error. Structural; analyzer is currently
silent — a clean, no-false-positive `empty-type` rule (sibling to
`duplicate-type-field`).

- Oracle case: `empty_type_block_compile` (now `vbe-oracle-verified`)
- Candidate rule family: declarations / UDT
- Reference: MS-VBAL 5.2.3.3 (UDT declaration).

### PCEC_006 Duplicate / conflicting Option

```vba
Option Explicit
Option Explicit

Public Sub Demo()
End Sub
```

**Oracle verdict (2026-06-13):** rejected — "Duplicate Option statement"
(`duplicate_option_explicit_compile`, compile-error). A repeated `Option Explicit`
is a compile error. Module-header only; analyzer is currently silent. `CANARY_002`
names this case.

- Oracle case: `duplicate_option_explicit_compile` (now `vbe-oracle-verified`)
- Candidate rule family: declarations / option placement (near `option-after-declaration`)
- Reference: MS-VBAL 5.2.1 (module options).
- Open: confirm `Option Compare Binary` + `Option Compare Text` (conflicting, not
  exact-duplicate) with a separate oracle case before extending the rule to it.

### PCEC_007 Duplicate / mis-ordered Case Else

```vba
Public Sub Demo()
    Dim n As Long
    n = 1
    Select Case n
        Case 1
            Debug.Print 1
        Case Else
            Debug.Print 2
        Case Else
            Debug.Print 3
    End Select
End Sub
```

**Oracle verdict (2026-06-13):** rejected — "Case without Select Case"
(`duplicate_case_else_compile`, compile-error). A `Select` with two `Case Else`
branches is a compile error (VBE reports the second one as outside its Select).
Structural; analyzer has `case-outside-select` but it does not fire on a second
`Case Else` that is structurally inside the block, so this is a gap.

- Oracle case: `duplicate_case_else_compile` (now `vbe-oracle-verified`)
- Candidate rule family: control flow
- Reference: MS-VBAL 5.4.2.10 (Select Case).

---

## Already covered

### PCEC_001 Duplicate parameter name

```vba
Public Sub DupParam(ByVal alpha As Long, ByVal alpha As Long)
End Sub
```

**Oracle verdict (2026-06-13):** rejected — "Duplicate declaration in current
scope" (`duplicate_parameter_name_compile`, compile-error). **Already covered** by
the `duplicate-declaration` rule, which fires on this signature; the oracle case
is now asserted evidence for that rule.

- Oracle case: `duplicate_parameter_name_compile` (now `vbe-oracle-verified`)
- Reference: MS-VBAL 5.3.1 (procedure declarations).

---

## Refuted by the oracle — VBE accepts, no rule

Building a red diagnostic for any of these would be a false positive. The
analyzer correctly stays silent. The oracle cases are retained as
`vbe-oracle-verified` *accepted* controls so the silence stays protected.

### PCEC_002 ByVal ParamArray

```vba
Public Sub TakesArgs(ByVal ParamArray items() As Variant)
End Sub
```

**Oracle verdict (2026-06-13):** **accepted** (`byval_paramarray_compile`,
compile-valid). Contrary to the original expectation, VBE compiles `ByVal
ParamArray`. No rule. The `paramarray-non-variant` / `paramarray-not-last` /
`paramarray-with-optional` family stands; do not add a ByVal check.

### PCEC_003 Empty Enum block

```vba
Public Enum EmptyEnum
End Enum
```

**Oracle verdict (2026-06-13):** **accepted** (`empty_enum_block_compile`,
compile-valid). VBE compiles an empty `Enum`. No rule. (Contrast PCEC_004: an
empty `Type` *is* rejected.)

### PCEC_005 Invalid Const type

```vba
Public Sub Demo()
    Const Handle As Object = Nothing
End Sub
```

**Oracle verdict (2026-06-13):** **accepted** (`const_object_type_compile`,
compile-valid). VBE compiles `Const Handle As Object = Nothing`. No rule for this
form. (A different object/array `Const` initializer might still be rejected; that
would need its own oracle case before any rule.)

---

## Binder-gated (confirmed, deferred)

### PCEC_008 Positional argument after named argument

```vba
Public Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)
End Sub

Public Sub Demo()
    NamedArgs alpha:=1, 2
End Sub
```

**Oracle verdict (2026-06-13):** rejected — "Syntax error"
(`positional_after_named_argument_compile`, compile-error). A positional argument
may not follow a named argument. Confirmed invalid, but detecting it needs parsed
call-argument lists; the MS-VBAL 5.6 expression binder now exists, so this is a
candidate once argument-slot named/positional shape is modeled.

- Oracle case: `positional_after_named_argument_compile` (now `vbe-oracle-verified`)
- Candidate rule family: call arity / argument lists (binder-dependent)
- Reference: MS-VBAL 5.6 (argument lists).

---

## Promotion path

1. ~~Run each observe-only oracle probe against the Excel/VBE oracle~~ **Done
   2026-06-13** — all eight verdicts recorded above and the oracle cases promoted
   to `vbe-oracle-verified`.
2. ~~For each *confirmed* case (PCEC_004, PCEC_006, PCEC_007), implement the
   diagnostic rule~~ **Done** — shipped as `empty-type`, `duplicate-option`, and
   `duplicate-case-else`, each `vbe-oracle-verified` with an asserted oracle case
   and tests in `tests/diagnostics/oracleConfirmedRules.test.ts`.
3. Leave the *refuted* cases (PCEC_002, PCEC_003, PCEC_005) ruleless; their
   accepted oracle controls guard against a future false positive.
