# Excel VBA Realtime Linting Addendum: Final Hardening Cases

**Purpose:** Add the remaining high-value tests for a VS Code Excel VBA realtime syntax linter after the basic syntax, edge-case, and limits corpora are already covered.

**Audience:** LLM agent implementing or validating XLIDE-style realtime syntax linting, semantic diagnostics, project-aware symbol analysis, and Excel host-aware warnings.

**Non-negotiable rule:** Treat Microsoft MS-VBAL as the grammar oracle. Treat Microsoft Office VBA docs as the behavior oracle for editor/compiler diagnostics. Do not borrow rules from VB.NET, VBScript, TypeScript, Python, or generic BASIC.

Official references to verify before implementation:

- MS-VBAL lexical rules: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/71e79228-eb45-4f81-8c0d-a224b72a0e47
- Option Explicit: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement
- Option Base: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement
- Option Private Module: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-private-statement
- Invalid inside procedure: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/invalid-inside-procedure
- Ambiguous name detected: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ambiguous-name-detected
- Set statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement
- Let statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/let-statement
- Sub statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement
- Function statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement
- Property Get: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-get-statement
- Property Let: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-let-statement
- Property Set: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-set-statement
- Property inconsistency diagnostics: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/definitions-of-property-procedures-for-the-same-property-are-inconsistent
- Implements: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/implements-statement
- Friend keyword: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/friend-keyword
- Type statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-statement
- ReDim statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/redim-statement
- ParamArray and Optional argument rules: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement
- Label not defined: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/label-not-defined
- With statement: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement

---

## Agent instruction

Add these cases as a fourth suite. Do **not** collapse them into generic parser tests. These are hardening cases for the places a realtime linter usually becomes annoying:

1. Module-kind-sensitive rules.
2. Directive placement.
3. Project-wide symbol conflicts.
4. Procedure signature rules.
5. Object/value assignment rules.
6. Property procedure consistency.
7. Excel host-object ambiguity.
8. Realtime partial-code recovery.

Expected diagnostic levels:

- `syntax-error`: code cannot be parsed as valid VBA.
- `compile-error`: parseable VBA that the VBA compiler rejects.
- `semantic-warning`: suspicious or project-dependent, but not necessarily rejected.
- `host-warning`: Excel-specific risk, not pure VBA grammar.
- `incomplete`: while typing, suppress hard errors until enough context exists.

---

# A. Module-kind-sensitive validation

The same text may be valid in one module kind and invalid or suspicious in another. The linter must know whether the file represents a standard `.bas`, class `.cls`, userform `.frm`, sheet module, or `ThisWorkbook` module.

## MOD_KIND_001 valid standard module public Sub

```vba
Option Explicit

Public Sub RefreshReport()
    Debug.Print "refresh"
End Sub
```

Expected: valid in standard module.

---

## MOD_KIND_002 suspicious event stub in standard module

```vba
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print Target.Address
End Sub
```

Expected: parse valid. Host-aware semantic warning if file is a standard module.

Reason: This is syntactically valid VBA, but Excel will not treat it as the worksheet event handler unless it is in the correct worksheet module.

---

## MOD_KIND_003 valid worksheet event in worksheet module

```vba
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    Debug.Print Target.Address
End Sub
```

Expected: valid in worksheet module.

Agent note: If the linter cannot know the module kind yet, mark this as valid syntax and avoid a false error.

---

## MOD_KIND_004 valid ThisWorkbook event only in ThisWorkbook module

```vba
Option Explicit

Private Sub Workbook_Open()
    Debug.Print "opened"
End Sub
```

Expected: valid syntax. Host-aware warning if found outside `ThisWorkbook`.

---

## MOD_KIND_005 Friend procedure only in class or form module

```vba
Option Explicit

Friend Sub InternalOnly()
    Debug.Print "friend"
End Sub
```

Expected:

- Valid in class module or form module.
- Invalid or compile-error in standard module.

Agent note: `Friend` modifies procedures, not variables or user-defined types.

---

# B. Module directives and declaration placement

## DIRECTIVE_001 valid directive order before procedures

```vba
Option Explicit
Option Base 1
Option Private Module

Public Sub Demo()
    Dim values(3) As Long
    Debug.Print LBound(values)
End Sub
```

Expected: valid.

---

## DIRECTIVE_002 invalid Option Explicit after procedure

```vba
Public Sub Demo()
End Sub

Option Explicit
```

Expected: compile-error.

Reason: `Option Explicit` must appear before procedures.

---

## DIRECTIVE_003 invalid Option Base inside procedure

```vba
Option Explicit

Public Sub Demo()
    Option Base 1
    Dim values(3) As Long
End Sub
```

Expected: compile-error.

Reason: `Option Base` is module-level only.

---

## DIRECTIVE_004 invalid duplicate Option Base

```vba
Option Explicit
Option Base 0
Option Base 1

Public Sub Demo()
End Sub
```

Expected: compile-error.

Reason: `Option Base` can appear only once in a module.

---

## DIRECTIVE_005 invalid module declarations inside procedure

```vba
Option Explicit

Public Sub Demo()
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
End Sub
```

Expected: compile-error.

Reason: `Declare`, `Private`, `Public`, `Option Base`, `Option Explicit`, `Enum`, and `Type` are not valid inside procedure bodies.

---

## DIRECTIVE_006 valid DefType at module level

```vba
Option Explicit
DefLng A-Z

Public Sub Demo()
    Dim value
    Debug.Print TypeName(value)
End Sub
```

Expected: valid.

Agent note: `value` is implicitly affected by the `DefLng` rule if no explicit type is provided. This is semantic behavior, not a parser error.

---

## DIRECTIVE_007 invalid DefType inside procedure

```vba
Option Explicit

Public Sub Demo()
    DefLng A-Z
    Dim value
End Sub
```

Expected: compile-error.

---

# C. Option Explicit and implicit Variant behavior

## EXPLICIT_001 invalid undeclared variable with Option Explicit

```vba
Option Explicit

Public Sub Demo()
    total = 10
End Sub
```

Expected: compile-error.

Reason: `total` is undeclared.

---

## EXPLICIT_002 valid implicit Variant without Option Explicit

```vba
Public Sub Demo()
    total = 10
    Debug.Print total
End Sub
```

Expected: valid syntax, no compile-error for undeclared variable.

Agent note: The linter may emit a style warning, but not a syntax error.

---

## EXPLICIT_003 typo caught only if symbol table is enabled

```vba
Option Explicit

Public Sub Demo()
    Dim customerCount As Long
    customerCount = 10
    Debug.Print customerCout
End Sub
```

Expected: compile-error for `customerCout`.

Realtime behavior: while typing `customerCou`, prefer `incomplete` or delayed diagnostic until identifier boundary or debounce.

---

# D. Declaration semantics that look like syntax trivia

## DECL_001 valid but surprising comma declaration

```vba
Option Explicit

Public Sub Demo()
    Dim a, b As Long
    a = "text"
    b = 42
End Sub
```

Expected: valid.

Semantic note: `a` is Variant, `b` is Long. Do not infer `a As Long` from `b As Long`.

---

## DECL_002 explicit types on each variable

```vba
Option Explicit

Public Sub Demo()
    Dim a As Long, b As Long
    a = 1
    b = 2
End Sub
```

Expected: valid.

---

## DECL_003 invalid New with intrinsic type

```vba
Option Explicit

Public Sub Demo()
    Dim value As New Long
End Sub
```

Expected: compile-error.

Reason: `New` applies to object variables, not intrinsic data types.

---

## DECL_004 valid auto-instancing object variable

```vba
Option Explicit

Private mItems As New Collection

Public Sub Demo()
    mItems.Add "x"
End Sub
```

Expected: valid.

Agent note: Auto-instancing is legal but can be style-warned if desired.

---

## DECL_005 invalid fixed-length string in public object module context

```vba
Option Explicit

Public FixedName As String * 20
```

Expected: project/module-kind-sensitive.

Agent note: Verify exact public/private fixed-string restrictions against MS-VBAL and host compiler. Do not guess. If unsure, classify as `needs-verification` instead of hardcoding.

---

# E. Procedure signature rules

## SIG_001 valid Optional arguments after first Optional

```vba
Option Explicit

Public Sub Demo(Optional ByVal first As Long = 0, Optional ByVal second As String = "")
End Sub
```

Expected: valid.

---

## SIG_002 invalid required argument after Optional

```vba
Option Explicit

Public Sub Demo(Optional ByVal first As Long = 0, ByVal second As Long)
End Sub
```

Expected: compile-error.

Reason: Once an argument is Optional, all following arguments must also be Optional.

---

## SIG_003 valid ParamArray last

```vba
Option Explicit

Public Sub Demo(ByVal caption As String, ParamArray values() As Variant)
    Debug.Print caption, UBound(values)
End Sub
```

Expected: valid.

---

## SIG_004 invalid ParamArray not last

```vba
Option Explicit

Public Sub Demo(ParamArray values() As Variant, ByVal caption As String)
End Sub
```

Expected: compile-error.

---

## SIG_005 invalid ParamArray with ByVal

```vba
Option Explicit

Public Sub Demo(ByVal ParamArray values() As Variant)
End Sub
```

Expected: compile-error.

---

## SIG_006 invalid ParamArray typed as non-Variant

```vba
Option Explicit

Public Sub Demo(ParamArray values() As Long)
End Sub
```

Expected: compile-error.

---

## SIG_007 invalid Optional UDT parameter

```vba
Option Explicit

Private Type Customer
    Id As Long
End Type

Public Sub Demo(Optional ByVal value As Customer)
End Sub
```

Expected: compile-error.

Reason: Optional parameters cannot be user-defined types.

---

## SIG_008 ByRef is default in VBA

```vba
Option Explicit

Public Sub Demo(value As Long)
    value = value + 1
End Sub
```

Expected: valid.

Semantic note: `value` is ByRef by default. Do not assume VB.NET-style ByVal.

---

# F. Set, Let, object assignment, and default members

## ASSIGN_001 valid Set object reference

```vba
Option Explicit

Public Sub Demo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Debug.Print ws.Name
End Sub
```

Expected: valid.

---

## ASSIGN_002 invalid object assignment without Set when statically known

```vba
Option Explicit

Public Sub Demo()
    Dim ws As Worksheet
    ws = ThisWorkbook.Worksheets(1)
End Sub
```

Expected: compile-error if Excel object model symbols are available.

Agent note: If `Worksheet` type is unresolved, downgrade to semantic warning, not syntax-error.

---

## ASSIGN_003 invalid Set for scalar assignment

```vba
Option Explicit

Public Sub Demo()
    Dim name As String
    Set name = "Sheet1"
End Sub
```

Expected: compile-error.

---

## ASSIGN_004 valid implicit Let assignment

```vba
Option Explicit

Public Sub Demo()
    Dim name As String
    name = "Sheet1"
End Sub
```

Expected: valid.

---

## ASSIGN_005 Excel Range default member ambiguity

```vba
Option Explicit

Public Sub Demo()
    Range("A1") = 42
    Debug.Print Range("A1")
End Sub
```

Expected: valid syntax. Optional host-warning only if the project policy requires explicit `.Value` or `.Value2`.

Agent note: Do not flag this as invalid just because no property is named. Excel default members are real.

---

## ASSIGN_006 default member chain should not be over-eagerly rejected

```vba
Option Explicit

Public Sub Demo()
    Dim text As String
    text = Range("A1")
End Sub
```

Expected: valid syntax. Host-aware semantic warning optional.

Reason: This is Excel-default-member behavior. A strict style policy may prefer `Range("A1").Value2`, but parser should not reject it.

---

# G. Property procedure consistency

## PROP_001 valid scalar Property Get and Let

```vba
Option Explicit

Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property
```

Expected: valid.

---

## PROP_002 valid object Property Get and Set

```vba
Option Explicit

Private mSheet As Worksheet

Public Property Get Sheet() As Worksheet
    Set Sheet = mSheet
End Property

Public Property Set Sheet(ByVal value As Worksheet)
    Set mSheet = value
End Property
```

Expected: valid.

---

## PROP_003 invalid Let/Set mismatch for object property

```vba
Option Explicit

Private mSheet As Worksheet

Public Property Get Sheet() As Worksheet
    Set Sheet = mSheet
End Property

Public Property Let Sheet(ByVal value As Worksheet)
    Set mSheet = value
End Property
```

Expected: compile-error or semantic-error.

Reason: Object reference assignment should use `Property Set`, not `Property Let`.

---

## PROP_004 invalid inconsistent Get and Let types

```vba
Option Explicit

Private mCount As Long

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Let Count(ByVal value As String)
    mCount = CLng(value)
End Property
```

Expected: compile-error.

Reason: paired property procedure argument and return types must be consistent.

---

## PROP_005 invalid Optional in Property procedure

```vba
Option Explicit

Private mName As String

Public Property Get Name(Optional ByVal fallback As String = "") As String
    Name = mName
End Property
```

Expected: compile-error.

Reason: Optional and ParamArray parameters are not permitted in Property procedures.

---

# H. Duplicate symbols, shadowing, and ambiguity

## SYMBOL_001 invalid duplicate procedure names in same module

```vba
Option Explicit

Public Sub Demo()
End Sub

Private Sub Demo()
End Sub
```

Expected: compile-error, ambiguous name detected.

---

## SYMBOL_002 valid local shadowing of module variable

```vba
Option Explicit

Private value As Long

Public Sub Demo()
    Dim value As Long
    value = 1
    Debug.Print value
End Sub
```

Expected: valid. Optional style warning.

---

## SYMBOL_003 qualified access to hidden module variable

```vba
Option Explicit

Private value As Long

Public Sub Demo()
    Dim value As Long
    value = 1
    Module1.value = 2
End Sub
```

Expected: project-dependent.

Agent note: Valid only if the module is actually named `Module1` and the member visibility permits the qualification. Do not hardcode this without project metadata.

---

## SYMBOL_004 module-level identifier conflicts with procedure name

```vba
Option Explicit

Private Customer As String

Public Sub Customer()
End Sub
```

Expected: compile-error, ambiguous name detected.

---

## SYMBOL_005 project-level module name conflicts with procedure-local variable

```vba
Option Explicit

Public Sub Demo()
    Dim Module1 As String
    Module1 = "local"
End Sub
```

Expected: valid syntax. Optional warning if project has a module named `Module1`.

Reason: VBA allows shadowing, but it can require qualification to access wider-scope members.

---

# I. Labels, line numbers, GoTo, GoSub, and On Error

## LABEL_001 valid named label and GoTo

```vba
Option Explicit

Public Sub Demo()
    GoTo CleanExit
    Debug.Print "skip"
CleanExit:
    Debug.Print "done"
End Sub
```

Expected: valid.

---

## LABEL_002 valid numeric line labels

```vba
Option Explicit

Public Sub Demo()
10  Debug.Print "start"
20  GoTo 40
30  Debug.Print "skip"
40  Debug.Print "done"
End Sub
```

Expected: valid.

---

## LABEL_003 invalid missing label

```vba
Option Explicit

Public Sub Demo()
    GoTo MissingLabel
End Sub
```

Expected: compile-error.

Reason: labels are procedure-scoped.

---

## LABEL_004 invalid cross-procedure GoTo

```vba
Option Explicit

Public Sub A()
    GoTo SharedLabel
End Sub

Public Sub B()
SharedLabel:
    Debug.Print "B"
End Sub
```

Expected: compile-error.

Reason: labels are visible only inside their own procedure.

---

## LABEL_005 valid On Error target in same procedure

```vba
Option Explicit

Public Sub Demo()
    On Error GoTo CleanFail
    Err.Raise 5
    Exit Sub
CleanFail:
    Debug.Print Err.Number
End Sub
```

Expected: valid.

---

## LABEL_006 valid On Error GoTo 0

```vba
Option Explicit

Public Sub Demo()
    On Error Resume Next
    On Error GoTo 0
End Sub
```

Expected: valid.

---

# J. With blocks and member access recovery

## WITH_001 valid simple With

```vba
Option Explicit

Public Sub Demo()
    With ThisWorkbook.Worksheets(1)
        .Range("A1").Value = 42
        .Name = "Data"
    End With
End Sub
```

Expected: valid.

---

## WITH_002 valid nested With, inner masks outer

```vba
Option Explicit

Public Sub Demo()
    With ThisWorkbook.Worksheets(1)
        With .Range("A1")
            .Value = 42
        End With
    End With
End Sub
```

Expected: valid.

Semantic note: inner With changes the meaning of leading-dot member access.

---

## WITH_003 invalid leading-dot outside With

```vba
Option Explicit

Public Sub Demo()
    .Range("A1").Value = 42
End Sub
```

Expected: syntax-error or compile-error.

---

## WITH_004 realtime incomplete member access after dot

```vba
Option Explicit

Public Sub Demo()
    ThisWorkbook.
End Sub
```

Expected while typing: `incomplete`, not permanent syntax-error.

Agent note: This is where completion popup and linter must cooperate.

---

## WITH_005 realtime incomplete With leading dot

```vba
Option Explicit

Public Sub Demo()
    With ThisWorkbook.Worksheets(1)
        .
    End With
End Sub
```

Expected while typing: `incomplete`, with member-completion context from the With target.

Promotion note: `bare_leading_member_access_inside_with_compile` verifies that
VBE rejects this completed form as Syntax error. Live syntax suppresses the error
only while the cursor is in the incomplete member-access edit state; save,
current-module lint, and workbook validation report it as syntax invalid.

---

# K. Conditional compilation and inactive branches

## COND_001 valid conditional compilation structure

```vba
Option Explicit

#Const DEBUG_MODE = True

Public Sub Demo()
#If DEBUG_MODE Then
    Debug.Print "debug"
#Else
    Debug.Print "release"
#End If
End Sub
```

Expected: valid.

---

## COND_002 inactive branch contains syntax-looking junk

```vba
Option Explicit

#Const DEBUG_MODE = True

Public Sub Demo()
#If DEBUG_MODE Then
    Debug.Print "debug"
#Else
    this is not normal vba syntax
#End If
End Sub
```

Expected: mode-dependent.

Recommended behavior:

- Foreground linter: suppress hard syntax errors in inactive branch if the condition is statically known.
- Background validator: optional warning that inactive branch is not parseable.

Agent note: Verify exact compiler behavior against MS-VBAL conditional compilation grammar before final implementation.

---

## COND_003 incomplete conditional block while typing

```vba
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
```

Expected while typing: `incomplete`, not syntax-error until debounce or file save.

---

# L. Arrays, ReDim, Preserve, and Option Base

## ARRAY_001 valid dynamic array then ReDim

```vba
Option Explicit

Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10)
    values(1) = 42
End Sub
```

Expected: valid.

---

## ARRAY_002 invalid ReDim fixed array

```vba
Option Explicit

Public Sub Demo()
    Dim values(1 To 10) As Long
    ReDim values(1 To 20)
End Sub
```

Expected: compile-error.

---

## ARRAY_003 valid ReDim creates undeclared dynamic array only without Option Explicit

```vba
Public Sub Demo()
    ReDim values(1 To 10)
    values(1) = 42
End Sub
```

Expected: valid without `Option Explicit`.

---

## ARRAY_004 invalid ReDim undeclared with Option Explicit

```vba
Option Explicit

Public Sub Demo()
    ReDim values(1 To 10)
End Sub
```

Expected: compile-error.

---

## ARRAY_005 Preserve can resize only last dimension

```vba
Option Explicit

Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 2, 1 To 2)
    ReDim Preserve values(1 To 3, 1 To 2)
End Sub
```

Expected: compile-error or runtime-error depending compiler acceptance. Treat as semantic diagnostic if analyzable.

Agent note: Verify exact timing of diagnostic in real Excel VBA.

---

## ARRAY_006 ParamArray lower bound unaffected by Option Base

```vba
Option Explicit
Option Base 1

Public Sub Demo(ParamArray values() As Variant)
    Debug.Print LBound(values)
End Sub
```

Expected: valid. Semantic note: `ParamArray` lower bound is zero.

---

# M. Type, Enum, constants, and object module restrictions

## TYPE_001 valid UDT at module level

```vba
Option Explicit

Private Type Customer
    Id As Long
    Name As String
End Type

Public Sub Demo()
    Dim c As Customer
    c.Id = 1
End Sub
```

Expected: valid.

---

## TYPE_002 invalid UDT inside procedure

```vba
Option Explicit

Public Sub Demo()
    Type Customer
        Id As Long
    End Type
End Sub
```

Expected: compile-error.

---

## TYPE_003 invalid line label inside Type block

```vba
Option Explicit

Private Type Customer
10  Id As Long
End Type
```

Expected: compile-error.

---

## TYPE_004 Enum at module level

```vba
Option Explicit

Public Enum ExportMode
    ExportCsv = 1
    ExportJson = 2
End Enum
```

Expected: valid in standard module. Module-kind visibility restrictions may apply in object modules.

Agent note: Verify public/private restrictions by module kind.

---

## TYPE_005 Public Const in object module may be invalid

```vba
Option Explicit

Public Const MaxRows As Long = 1000
```

Expected: module-kind-sensitive.

Agent note: Standard module is different from class/form/sheet modules. Use module metadata before hard-erroring.

---

# N. Implements and interface member shape

## IMPL_001 valid Implements statement in class module

```vba
Option Explicit

Implements IRunner
```

Expected: valid only in a class/form module with a resolvable `IRunner` class/interface.

---

## IMPL_002 suspicious Implements in standard module

```vba
Option Explicit

Implements IRunner
```

Expected: compile-error in standard module.

---

## IMPL_003 missing implemented member

```vba
Option Explicit

Implements IRunner

Private Sub Class_Initialize()
End Sub
```

Expected: project-aware semantic diagnostic only if `IRunner` members are known.

Agent note: Do not flag this from syntax alone. It requires project symbol graph.

---

# O. Excel-specific host warnings that are not syntax errors

## HOST_001 unqualified Range in standard module

```vba
Option Explicit

Public Sub Demo()
    Range("A1").Value = 42
End Sub
```

Expected: valid syntax. Optional host-warning: unqualified `Range` binds through active sheet context.

---

## HOST_002 qualified Range is cleaner

```vba
Option Explicit

Public Sub Demo()
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = 42
End Sub
```

Expected: valid.

---

## HOST_003 ActiveWorkbook vs ThisWorkbook is not syntax

```vba
Option Explicit

Public Sub Demo()
    ActiveWorkbook.Worksheets(1).Range("A1").Value = 42
End Sub
```

Expected: valid syntax. Optional host-warning depending project policy.

---

## HOST_004 Application-qualified call should be understood

```vba
Option Explicit

Public Sub Demo()
    Application.WorksheetFunction.Sum Range("A1:A10")
End Sub
```

Expected: valid syntax.

Agent note: Procedure-call grammar and Excel object model resolution are separate. Do not reject merely because parentheses are omitted in a Sub-style call.

---

# P. Realtime editor partial states

These are not final source files. They are transient states the linter sees while the user is typing.

## RT_001 half-written Sub

```vba
Option Explicit

Public Sub Demo(
```

Expected: `incomplete`.

---

## RT_002 half-written If block

```vba
Option Explicit

Public Sub Demo()
    If True Then
```

Expected: `incomplete`, not syntax-error during typing.

---

## RT_003 half-written string literal

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello
End Sub
```

Expected: `incomplete` while typing, eventually syntax-error after debounce/save.

---

## RT_004 half-written named argument

```vba
Option Explicit

Public Sub Demo()
    MsgBox Prompt:=
End Sub
```

Expected: `incomplete`.

---

## RT_005 half-written member chain

```vba
Option Explicit

Public Sub Demo()
    ThisWorkbook.Worksheets(1).Range("A1").
End Sub
```

Expected: `incomplete` with completion context.

---

## RT_006 partially typed line continuation

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello" & _
```

Expected: `incomplete` until next physical line exists.

---

## RT_007 pasted block with missing End If should recover below error

```vba
Option Explicit

Public Sub Demo()
    If True Then
        Debug.Print "x"

    Debug.Print "after"
End Sub
```

Expected: missing `End If` diagnostic, but parser should recover and continue diagnostics below the block.

---

# Q. Suggested test runner fields

Each fixture should include:

```json
{
  "id": "SIG_004",
  "moduleKind": "standard",
  "code": "...",
  "expected": {
    "parse": "ok",
    "diagnostics": [
      {
        "severity": "error",
        "kind": "compile-error",
        "messageContains": "ParamArray",
        "line": 4
      }
    ]
  },
  "requiresProjectSymbols": false,
  "requiresExcelHostModel": false,
  "realtimeState": false
}
```

Recommended module kinds:

- `standard`
- `class`
- `userform`
- `worksheet`
- `thisworkbook`
- `unknown`

Recommended phases:

- `lex`
- `parse`
- `compile-shape`
- `symbol-table`
- `excel-host`
- `realtime-recovery`

---

# R. The main implementation warning

Do not try to make one diagnostic pass do everything.

Recommended pipeline:

1. **Lexer:** comments, strings, line continuations, attributes, conditional directives.
2. **Tolerant parser:** build a tree even when the user is mid-token or mid-block.
3. **Strict parser:** validate saved or debounced source against MS-VBAL grammar.
4. **Compile-shape checker:** directive placement, declaration placement, block matching, procedure signatures.
5. **Project symbol checker:** duplicate names, shadowing, Implements, property pairs, label targets.
6. **Excel host checker:** worksheet/workbook events, object model calls, unqualified host members, default members.
7. **Style analyzer:** optional warnings like implicit default members, unqualified `Range`, `Dim a, b As Long`, auto-instancing `New`.

A good realtime VBA linter should feel forgiving while typing and strict after the code settles. That means `incomplete` is not a weaker error. It is a different state.
