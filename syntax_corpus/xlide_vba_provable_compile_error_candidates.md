# XLIDE VBA Provable Compile-Error Candidates

**Status:** `pending-verification`. Every case below is discovery material, not
diagnostic-driving evidence. Do not use a Markdown expectation here to justify a
red diagnostic until it is promoted through MS-VBAL, a focused VBE oracle case,
or deterministic XLIDE-owned metadata with tests.

**Purpose:** capture hard, deterministic VBA compile errors that are decidable
from a single declaration, signature, call, or module header - no expression
binder, host metadata, or runtime evidence required - and that are **not yet**
covered by an active diagnostic rule, an oracle case, or the existing corpus.
These are red-squiggly candidates that extend rule families XLIDE already ships.

Each case names the matching observe-only oracle probe in
`syntax_corpus/oracle/vbe_oracle_cases.json` (added alongside this file) so the
exact VBE verdict and message can be recorded before promotion.

**Coverage note:** duplicate *named* arguments are already
`vbe-oracle-verified` (`duplicate_named_argument_compile`) and intentionally
omitted here. Its sibling, a positional argument after a named argument, is
included below as binder-gated.

---

## Ship-now group (binder-independent)

These are decidable from declaration/signature/header text alone and obey the
no-false-positive rule today.

### PCEC_001 Duplicate parameter name

```vba
Public Sub DupParam(ByVal alpha As Long, ByVal alpha As Long)
End Sub
```

Expected: invalid. VBE rejects a signature that repeats a parameter name
("Ambiguous name detected" / "Duplicate declaration in current scope"). Pure
signature syntax; sibling to `duplicate-declaration`.

- Oracle probe: `duplicate_parameter_name_compile`
- Candidate rule family: declarations / duplicates
- Reference: MS-VBAL procedure declarations (5.3.1); Ambiguous name detected.

### PCEC_002 ByVal ParamArray

```vba
Public Sub TakesArgs(ByVal ParamArray items() As Variant)
End Sub
```

Expected: invalid. A `ParamArray` parameter must be a `ByRef` `Variant` array;
declaring it `ByVal` is a compile error. Pure signature syntax; slots into the
existing `paramarray-not-last` / `paramarray-with-optional` /
`paramarray-non-variant` family.

- Oracle probe: `byval_paramarray_compile`
- Candidate rule family: paramarray
- Reference: MS-VBAL 5.3.1 (ParamArray); Sub/Function statement docs.

### PCEC_003 Empty Enum block

```vba
Public Enum EmptyEnum
End Enum
```

Expected: invalid. An `Enum` declaring no members is a compile error.
Structural; sibling to `duplicate-enum-member`.

- Oracle probe: `empty_enum_block_compile`
- Candidate rule family: declarations / enum
- Reference: MS-VBAL 5.2.3.4 (Enum declaration).

### PCEC_004 Empty Type block

```vba
Public Type EmptyType
End Type
```

Expected: invalid. A user-defined `Type` declaring no members is a compile
error. Structural.

- Oracle probe: `empty_type_block_compile`
- Candidate rule family: declarations / UDT
- Reference: MS-VBAL 5.2.3.3 (UDT declaration); Type statement docs.

### PCEC_005 Invalid Const type

```vba
Public Sub Demo()
    Const Handle As Object = Nothing
End Sub
```

Expected: invalid. A `Const` must be a simple intrinsic type; an object (or
array) `As` clause is a compile error. Declaration-only; sits next to
`const-assignment`.

- Oracle probe: `const_object_type_compile`
- Candidate rule family: declarations / const
- Reference: MS-VBAL 5.2.3.1 (Const declaration).
- Note: keep quiet when the `As` type name is an unresolved/ambiguous project
  type - only fire on a known non-simple type so this stays no-false-positive.

### PCEC_006 Duplicate / conflicting Option

```vba
Option Explicit
Option Explicit

Public Sub Demo()
End Sub
```

Expected: invalid. A repeated `Option Explicit` (and likewise conflicting
`Option Compare Binary` + `Option Compare Text`) is a compile error.
Module-header only. `CANARY_002` names this case but there is no rule and no
verdict fixture.

- Oracle probe: `duplicate_option_explicit_compile`
- Candidate rule family: declarations / option placement (near
  `option-after-declaration`)
- Reference: MS-VBAL 5.2.1 (module options); Option Explicit / Option Compare
  docs.

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

Expected: invalid. A `Select` block with two `Case Else` branches (or a `Case`
after `Case Else`) is a compile error. Structural; XLIDE has
`case-outside-select` but not this ordering rule.

- Oracle probe: `duplicate_case_else_compile`
- Candidate rule family: control flow
- Reference: MS-VBAL 5.4.2.10 (Select Case).

---

## Binder-gated group

Provable, but only once call-argument lists are parsed. Today arguments are raw
text, so a rule here would risk false positives. Queue behind the MS-VBAL 5.6
expression binder.

### PCEC_008 Positional argument after named argument

```vba
Public Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)
End Sub

Public Sub Demo()
    NamedArgs alpha:=1, 2
End Sub
```

Expected: invalid. A positional argument may not follow a named argument.
Sibling to the verified `duplicate_named_argument_compile`.

- Oracle probe: `positional_after_named_argument_compile`
- Candidate rule family: call arity / argument lists (binder-dependent)
- Reference: MS-VBAL 5.6 (argument lists).

---

## Promotion path

1. Run each observe-only oracle probe against the Excel/VBE oracle and record the
   accepted/rejected verdict plus the exact compile message.
2. For cases backed cleanly by MS-VBAL grammar (PCEC_002 ParamArray, PCEC_003
   empty Enum, PCEC_004 empty Type), a spec citation may suffice; use the oracle
   to confirm the VBE message and severity.
3. Promote a confirmed case into an executable diagnostic fixture with
   `spec-derived` or `vbe-oracle-verified` provenance, add the valid /
   invalid / unknown-no-diagnostic controls, and update
   `diagnostic_influence_audit.json` when it begins driving an active rule.
