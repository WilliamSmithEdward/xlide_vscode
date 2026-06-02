# XLIDE VBA Corpus Addendum: Semantic, Runtime, and Resolution Edges

**Recommended file name:** `27_semantic_runtime_resolution_edges.md`

**Purpose:** Add a final optional hardening layer for XLIDE's Excel VBA realtime analysis suite.

This file is not a basic syntax corpus. It targets state-sensitive analysis:

1. Deterministic runtime faults.
2. Runtime faults suppressed by error handling.
3. Variant coercion and operator ambiguity.
4. Project reference resolution.
5. Host event signature binding.
6. Public API visibility problems.
7. `WithEvents` type restrictions.
8. Macro discoverability vs compile validity.
9. Predeclared class instances.
10. Branching into structured blocks.

**Target host:** Excel VBA.

**Primary grammar oracle:** Microsoft **MS-VBAL**.

**Behavior oracle:** live Excel VBE compile/runtime canary.

**Rule:** Do not infer VBA behavior from VB.NET, VBScript, TypeScript, Python, or generic BASIC. If the behavior is uncertain, record a canary verdict instead of guessing.

---

## Diagnostic Classes

Use these diagnostic categories consistently.

```text
syntax-error                    cannot parse as valid VBA
compile-error                   parseable, but VBE compiler rejects it
semantic-warning                suspicious or symbol-dependent
host-warning                    Excel-specific risk, not pure VBA grammar
style-warning                   valid but discouraged
incomplete                      valid partial state during active typing
provable-runtime-error          execution will fail if this statement is reached
suppressed-runtime-fault        execution would fault, but active error handling suppresses it
probable-runtime-warning        likely runtime risk, but not mechanically proven
reference-resolution-error      missing project reference or unresolved external type
host-binding-warning            event/macro binding issue, not raw VBA syntax
needs-verification              must be resolved by MS-VBAL and/or VBE canary
```

---

## Fixture Shape

```jsonc
{
  "id": "RUNTIME_001",
  "title": "Division by zero",
  "moduleKind": "standard",
  "expected": "provable-runtime-error",
  "phase": "runtime-static",
  "diagnostics": [
    {
      "code": "VBA_RUNTIME_DIVISION_BY_ZERO",
      "severity": "warning"
    }
  ],
  "snippet": "..."
}
```

For oracle-backed runtime behavior:

```jsonc
{
  "id": "COERCE_001",
  "title": "String plus number coercion",
  "moduleKind": "standard",
  "expected": "needs-verification",
  "phase": "runtime-canary",
  "vbeCanary": {
    "host": "Excel",
    "hostVersion": "recorded-by-runner",
    "bitness": "recorded-by-runner",
    "compileResult": "accept | reject",
    "runtimeResult": "no-error | runtime-error",
    "errorNumber": null,
    "observedValue": null,
    "notes": ""
  },
  "snippet": "..."
}
```

---

# 1. Deterministic Runtime Faults

## RUNTIME_001 provable runtime error: division by zero

```vba
Public Sub Demo()
    Debug.Print 1 / 0
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_DIVISION_BY_ZERO
```

---

## RUNTIME_002 provable runtime error: integer division by zero

```vba
Public Sub Demo()
    Debug.Print 1 \ 0
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_DIVISION_BY_ZERO
```

---

## RUNTIME_003 provable runtime error: modulo by zero

```vba
Public Sub Demo()
    Debug.Print 1 Mod 0
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_DIVISION_BY_ZERO
```

---

## RUNTIME_004 provable runtime error: negative length to Left$

```vba
Public Sub Demo()
    Debug.Print Left$("abcdef", -1)
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_INVALID_PROCEDURE_CALL
```

Canary note:

```text
Verify exact runtime error number and message.
```

---

## RUNTIME_005 provable runtime error: negative repeat count to String$

```vba
Public Sub Demo()
    Debug.Print String$(-1, "x")
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_INVALID_PROCEDURE_CALL
```

---

## RUNTIME_006 provable runtime error: array constant index out of bounds

```vba
Public Sub Demo()
    Dim values(1 To 3) As Long
    Debug.Print values(4)
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_SUBSCRIPT_OUT_OF_RANGE
```

Caution:

```text
Only mark provable when bounds and index are statically known.
```

---

## RUNTIME_007 provable runtime error: Nothing dereference

```vba
Public Sub Demo()
    Dim ws As Worksheet
    Set ws = Nothing
    Debug.Print ws.Name
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault if statement is reached
diagnostic: provable-runtime-error
code: VBA_RUNTIME_OBJECT_VARIABLE_NOT_SET
```

Caution:

```text
Only mark provable if no intervening assignment may set the object.
```

---

# 2. Runtime Faults and Error Handling

## ERROR_FLOW_001 suppressed runtime fault: On Error Resume Next

```vba
Public Sub Demo()
    On Error Resume Next
    Debug.Print 1 / 0
    Debug.Print Err.Number
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime fault: suppressed by active On Error Resume Next
diagnostic: suppressed-runtime-fault
message: Division by zero would raise an error, but active On Error Resume Next suppresses it.
```

---

## ERROR_FLOW_002 active error handler: On Error GoTo label

```vba
Public Sub Demo()
    On Error GoTo Handler
    Debug.Print 1 / 0
    Exit Sub

Handler:
    Debug.Print Err.Number
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime fault: transfers to handler
diagnostic: deterministic fault with active handler
not procedure-fails
```

---

## ERROR_FLOW_003 error handling disabled before fault

```vba
Public Sub Demo()
    On Error Resume Next
    Debug.Print "before"
    On Error GoTo 0
    Debug.Print 1 / 0
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: deterministic runtime fault is no longer suppressed
diagnostic: provable-runtime-error
```

---

## ERROR_FLOW_004 clear current exception state

```vba
Public Sub Demo()
    On Error Resume Next
    Err.Raise 5
    On Error GoTo -1
End Sub
```

Expected:

```text
syntax: valid if accepted by VBE
compile/runtime: verify with oracle
diagnostic: none unless host rejects
```

Canary note:

```text
Record exact compile and runtime behavior.
```

---

## ERROR_FLOW_005 unreachable deterministic runtime fault

```vba
Public Sub Demo()
    Exit Sub
    Debug.Print 1 / 0
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: unreachable in straight-line flow
diagnostic: optional unreachable-code warning
do not report as active provable-runtime-error unless analyzer reports unreachable suppressed fault separately
```

---

# 3. Variant Coercion and Operator Ambiguity

## COERCE_001 valid but warning: plus used for string concatenation

```vba
Public Sub Demo()
    Dim value As String
    value = "string" + "string"
    Debug.Print value
End Sub
```

Expected:

```text
syntax: valid
compile: accepted by Excel VBE oracle
runtime: no error
observed value: stringstring
diagnostic: optional style-warning
code: VBA_STYLE_PLUS_USED_FOR_STRING_CONCAT
message: Prefer & for string concatenation. + is legal for string operands but ambiguous with mixed types.
```

---

## COERCE_002 valid: canonical ampersand concatenation

```vba
Public Sub Demo()
    Dim value As String
    value = "string" & "string"
    Debug.Print value
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: no error
diagnostic: none
```

---

## COERCE_003 mixed plus: nonnumeric string plus number

```vba
Public Sub Demo()
    Dim value As Variant
    value = "x" + 2
    Debug.Print value
End Sub
```

Expected:

```text
syntax: valid
compile: verify with VBE oracle
runtime: likely deterministic Type mismatch if execution reaches statement
diagnostic: provable-runtime-error only after oracle confirmation
```

---

## COERCE_004 mixed plus: numeric-looking string plus number

```vba
Public Sub Demo()
    Dim value As Variant
    value = "1" + 2
    Debug.Print value
End Sub
```

Expected:

```text
syntax: valid
compile: verify with VBE oracle
runtime: likely numeric coercion
diagnostic: optional coercion warning only
not a syntax error
```

---

## COERCE_005 ampersand with number

```vba
Public Sub Demo()
    Dim value As String
    value = "1" & 2
    Debug.Print value
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: no error
observed value likely: 12
diagnostic: none
```

---

## COERCE_006 Null plus string

```vba
Public Sub Demo()
    Dim v As Variant
    v = Null
    Debug.Print v + "x"
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: verify with VBE oracle
diagnostic: oracle-backed
```

---

## COERCE_007 Null ampersand string

```vba
Public Sub Demo()
    Dim v As Variant
    v = Null
    Debug.Print v & "x"
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: verify with VBE oracle
diagnostic: oracle-backed
```

---

## COERCE_008 Empty plus numeric-looking string

```vba
Public Sub Demo()
    Dim v As Variant
    v = Empty
    Debug.Print v + "1"
End Sub
```

Expected:

```text
syntax: valid
compile: valid
runtime: verify with VBE oracle
diagnostic: optional coercion warning based on observed behavior
```

---

# 4. For Each Control Variable Restrictions

## FOREACH_001 invalid or compile-error: scalar Long control variable

```vba
Public Sub Demo()
    Dim x As Long
    For Each x In Array(1, 2, 3)
        Debug.Print x
    Next x
End Sub
```

Expected:

```text
syntax: valid shape
compile: verify with VBE oracle
likely compile-error: For Each control variable must be Variant or Object
diagnostic: compile-error if confirmed
```

---

## FOREACH_002 valid: Variant control variable

```vba
Public Sub Demo()
    Dim x As Variant
    For Each x In Array(1, 2, 3)
        Debug.Print x
    Next x
End Sub
```

Expected:

```text
valid
```

---

## FOREACH_003 valid: object control variable

```vba
Public Sub Demo()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
```

Expected:

```text
valid if Excel symbols are loaded
```

---

## FOREACH_004 invalid: array control variable

```vba
Public Sub Demo()
    Dim values() As Variant
    For Each values In Array(1, 2, 3)
        Debug.Print values
    Next values
End Sub
```

Expected:

```text
syntax: parseable
compile: verify with VBE oracle
likely compile-error
```

---

# 5. Branching Into Structured Blocks

## BRANCH_BLOCK_001 branch into If block

```vba
Public Sub Demo()
    GoTo Inside

    If True Then
Inside:
        Debug.Print "inside"
    End If
End Sub
```

Expected:

```text
syntax: parseable
compile: oracle-required
diagnostic: invalid branch into block if VBE rejects
```

---

## BRANCH_BLOCK_002 branch into For block

```vba
Public Sub Demo()
    GoTo Inside

    Dim i As Long
    For i = 1 To 3
Inside:
        Debug.Print i
    Next i
End Sub
```

Expected:

```text
syntax: parseable
compile/runtime: oracle-required
diagnostic: invalid branch into block if VBE rejects or runtime behavior is unsafe
```

---

## BRANCH_BLOCK_003 branch into With block

```vba
Public Sub Demo()
    GoTo Inside

    With ThisWorkbook
Inside:
        Debug.Print .Name
    End With
End Sub
```

Expected:

```text
syntax: parseable
compile/runtime: oracle-required
diagnostic: invalid branch into With block if VBE rejects
```

---

## BRANCH_BLOCK_004 branch out of block

```vba
Public Sub Demo()
    If True Then
        GoTo Outside
    End If

Outside:
    Debug.Print "outside"
End Sub
```

Expected:

```text
syntax: valid
compile: verify with VBE oracle
usually valid branch out of block
```

---

# 6. Project References and Late Binding

## REF_001 external type requiring project reference

```vba
Public Sub Demo()
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
End Sub
```

Expected:

```text
syntax: valid
compile: depends on Microsoft Scripting Runtime reference
diagnostic if missing: reference-resolution-error
not syntax error
```

---

## REF_002 late-bound dictionary

```vba
Public Sub Demo()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
End Sub
```

Expected:

```text
syntax: valid
compile: valid without explicit Scripting Runtime reference
runtime: depends on host availability
diagnostic: optional late-binding warning only
```

---

## REF_003 unresolved external enum constant

```vba
Public Sub Demo()
    Dim mode As FileSystemObject
End Sub
```

Expected:

```text
syntax: valid
compile: depends on project references
diagnostic: unresolved external type if reference absent
```

---

## REF_004 ambiguous type name across references

```vba
Public Sub Demo()
    Dim app As Application
    Set app = Application
End Sub
```

Expected:

```text
syntax: valid
semantic resolution depends on project references and host object model
diagnostic only if ambiguity is proven
```

---

# 7. Host Event Signature Binding

## HOST_EVENT_SIG_001 valid worksheet change signature

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print Target.Address
End Sub
```

Expected:

```text
valid in worksheet module
host event signature accepted
```

---

## HOST_EVENT_SIG_002 wrong argument type

```vba
Private Sub Worksheet_Change(ByVal Target As Object)
    Debug.Print TypeName(Target)
End Sub
```

Expected:

```text
syntax: valid procedure
host binding: likely invalid event signature
diagnostic: host-binding-warning or compile-error after oracle confirmation
not syntax error
```

---

## HOST_EVENT_SIG_003 missing argument

```vba
Private Sub Worksheet_Change()
    Debug.Print "changed"
End Sub
```

Expected:

```text
syntax: valid procedure
host binding: wrong event signature
diagnostic: host-binding-warning or compile-error after oracle confirmation
```

---

## HOST_EVENT_SIG_004 wrong module for workbook event

```vba
Private Sub Workbook_Open()
    Debug.Print "opened"
End Sub
```

Expected with `moduleKind: standard`:

```text
syntax: valid
host-binding-warning: Workbook_Open will not bind outside ThisWorkbook module
not syntax error
```

Expected with `moduleKind: workbook`:

```text
valid
```

---

## HOST_EVENT_SIG_005 valid Application event via WithEvents

```vba
Option Explicit

Private WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print Wb.Name
End Sub
```

Expected:

```text
valid in class or userform module if WithEvents declaration exists
handler binds to App event source
```

---

# 8. Public API Visibility and Private Types

## API_VIS_001 public function returns private UDT

```vba
Option Explicit

Private Type Customer
    Id As Long
End Type

Public Function GetCustomer() As Customer
End Function
```

Expected:

```text
syntax: valid shape
compile: oracle-required
diagnostic: public member exposes private type if VBE rejects
```

---

## API_VIS_002 private function returns private UDT

```vba
Option Explicit

Private Type Customer
    Id As Long
End Type

Private Function GetCustomer() As Customer
End Function
```

Expected:

```text
valid if confirmed by VBE oracle
```

---

## API_VIS_003 public method parameter uses private UDT

```vba
Option Explicit

Private Type Customer
    Id As Long
End Type

Public Sub SaveCustomer(ByVal value As Customer)
End Sub
```

Expected:

```text
syntax: valid shape
compile: oracle-required
diagnostic: public member exposes private type if VBE rejects
```

---

## API_VIS_004 public UDT in standard module

```vba
Option Explicit

Public Type Customer
    Id As Long
End Type

Public Function GetCustomer() As Customer
End Function
```

Expected:

```text
valid in standard module if VBE confirms
```

---

# 9. WithEvents Type Restrictions

## WITHEVENTS_TYPE_001 invalid: WithEvents As Object

```vba
Option Explicit

Private WithEvents App As Object
```

Expected:

```text
syntax: parseable
compile: oracle-required
likely compile-error because WithEvents needs a specific event source type
```

---

## WITHEVENTS_TYPE_002 valid: WithEvents As Application

```vba
Option Explicit

Private WithEvents App As Application
```

Expected:

```text
valid in class or userform module
invalid in standard module
```

---

## WITHEVENTS_TYPE_003 invalid: WithEvents with late-bound CreateObject assignment only

```vba
Option Explicit

Private WithEvents App As Object

Private Sub Class_Initialize()
    Set App = CreateObject("Excel.Application")
End Sub
```

Expected:

```text
compile: oracle-required
diagnostic: WithEvents cannot use generic Object if VBE rejects
```

---

## WITHEVENTS_TYPE_004 valid: WithEvents assigned host Application

```vba
Option Explicit

Private WithEvents App As Application

Private Sub Class_Initialize()
    Set App = Application
End Sub
```

Expected:

```text
valid in class module if Excel symbols are loaded
```

---

# 10. Macro Discoverability vs Compile Validity

## MACRO_VIS_001 public Sub with no arguments

```vba
Public Sub MacroVisible()
    Debug.Print "visible macro"
End Sub
```

Expected:

```text
compile: valid
Excel macro dialog visibility: visible candidate
diagnostic: none
```

---

## MACRO_VIS_002 public Sub with required argument

```vba
Public Sub MacroWithArgs(ByVal x As Long)
    Debug.Print x
End Sub
```

Expected:

```text
compile: valid
Excel macro dialog visibility: not normal macro-dialog runnable
diagnostic: optional host-info, not error
```

---

## MACRO_VIS_003 private Sub

```vba
Private Sub PrivateMacro()
    Debug.Print "private"
End Sub
```

Expected:

```text
compile: valid
Excel macro dialog visibility: not public macro
diagnostic: optional host-info, not error
```

---

## MACRO_VIS_004 public Function

```vba
Public Function FunctionMacro() As Long
    FunctionMacro = 1
End Function
```

Expected:

```text
compile: valid
function may be callable from worksheet depending on context
macro dialog behavior differs from public Sub
diagnostic: optional host-info only
```

---

## MACRO_VIS_005 Option Private Module

```vba
Option Private Module

Public Sub MacroHiddenFromOtherProjects()
    Debug.Print "hidden externally"
End Sub
```

Expected:

```text
compile: valid
host visibility affected by Option Private Module
diagnostic: optional host-info only
```

---

# 11. Predeclared Class Instances

## PREDECLARED_001 UserForm default instance

```vba
Public Sub Demo()
    UserForm1.Caption = "Hello"
    UserForm1.Show
End Sub
```

Expected:

```text
valid if UserForm1 exists
UserForms commonly have predeclared/default instance behavior
project metadata required
```

---

## PREDECLARED_002 ordinary class used as default instance

```vba
Public Sub Demo()
    Class1.DoWork
End Sub
```

Expected:

```text
syntax: valid shape
compile: depends on class metadata
if Class1 is ordinary class without predeclared instance, diagnostic
if exported metadata VB_PredeclaredId = True, may be valid
```

---

## PREDECLARED_003 explicit class instance

```vba
Public Sub Demo()
    Dim obj As Class1
    Set obj = New Class1
    obj.DoWork
End Sub
```

Expected:

```text
valid if Class1 exists and exposes DoWork
```

---

## PREDECLARED_004 exported class metadata with predeclared instance

Exported `.cls` metadata:

```vba
Attribute VB_Name = "Class1"
Attribute VB_PredeclaredId = True
```

Consumer code:

```vba
Public Sub Demo()
    Class1.DoWork
End Sub
```

Expected:

```text
valid only if metadata confirms predeclared instance
without metadata, do not assume ordinary classes behave like UserForms
```

---

# 12. Conditional Compilation and Semantic Reachability

## COND_SEM_001 inactive compile branch with bad type

```vba
#Const USE_BAD_TYPE = False

Public Sub Demo()
#If USE_BAD_TYPE Then
    Dim d As Missing.ExternalType
#Else
    Debug.Print "ok"
#End If
End Sub
```

Expected:

```text
syntax: valid conditional compilation structure
active semantics: no unresolved type if branch inactive
optional inactive-branch diagnostic only if analyzer checks all branches separately
```

---

## COND_SEM_002 active compile branch with bad type

```vba
#Const USE_BAD_TYPE = True

Public Sub Demo()
#If USE_BAD_TYPE Then
    Dim d As Missing.ExternalType
#Else
    Debug.Print "ok"
#End If
End Sub
```

Expected:

```text
syntax: valid
active semantics: unresolved external type diagnostic
```

---

## COND_SEM_003 conditional constant affects runtime fault reachability

```vba
#Const DEBUG_MODE = False

Public Sub Demo()
#If DEBUG_MODE Then
    Debug.Print 1 / 0
#End If
End Sub
```

Expected:

```text
if DEBUG_MODE is False, deterministic runtime fault is in inactive branch
do not report as active provable-runtime-error
```

---

# 13. Realtime Behavior

## RT_SEM_001 incomplete expression should not trigger runtime analysis

```vba
Public Sub Demo()
    Debug.Print 1 /
End Sub
```

Expected in realtime mode:

```text
incomplete expression
do not emit division-by-zero runtime diagnostic yet
```

---

## RT_SEM_002 partial string plus expression

```vba
Public Sub Demo()
    Debug.Print "x" +
End Sub
```

Expected in realtime mode:

```text
incomplete expression
do not classify as coercion warning or runtime fault until right operand exists
```

---

## RT_SEM_003 partial For Each

```vba
Public Sub Demo()
    Dim x As Long
    For Each x In
End Sub
```

Expected in realtime mode:

```text
incomplete For Each statement
do not emit final control-variable diagnostic until loop expression is complete
```

---

# 14. VBE Canary Matrix

Run these through the oracle and record compile/runtime verdicts.

```text
RUNTIME_004        Left$("abcdef", -1)
RUNTIME_005        String$(-1, "x")
COERCE_003         "x" + 2
COERCE_004         "1" + 2
COERCE_006         Null + "x"
COERCE_007         Null & "x"
COERCE_008         Empty + "1"
FOREACH_001        Long control variable in For Each
FOREACH_004        array control variable in For Each
BRANCH_BLOCK_001   branch into If block
BRANCH_BLOCK_002   branch into For block
BRANCH_BLOCK_003   branch into With block
BRANCH_BLOCK_004   branch out of block
API_VIS_001        public function returns private UDT
API_VIS_003        public method parameter uses private UDT
WITHEVENTS_TYPE_001 WithEvents As Object
PREDECLARED_002    ordinary class used as default instance
```

Recommended canary record:

```jsonc
{
  "id": "COERCE_003",
  "host": "Excel",
  "hostVersion": "recorded-by-runner",
  "bitness": "recorded-by-runner",
  "moduleKind": "standard",
  "sourceMode": "live-vbe",
  "compileResult": "accept | reject",
  "runtimeResult": "not-run | no-error | runtime-error",
  "errorNumber": null,
  "errorMessageContains": null,
  "observedValue": null,
  "notes": ""
}
```

---

# 15. Integration Guidance

## Do Not Run This Too Early

This suite should run after:

```text
lexical parsing
statement parsing
declaration parsing
scope recovery
module-kind classification
project symbol indexing
conditional compilation evaluation
basic host symbol loading
```

If the parser is still unstable, this suite will produce misleading noise.

---

## Realtime Policy

Do not run deterministic runtime analysis aggressively on every keystroke.

Recommended modes:

```text
onType: syntax recovery and completion context only
onSave: syntax, compile-shape, semantic checks
onDemand: full semantic, host, reference, runtime-static checks
```

Runtime-static diagnostics should be precise, sparse, and high-confidence.

---

## Confidence Rules

Only emit `provable-runtime-error` when all of the following are true:

```text
1. The code is parseable.
2. The code would compile, or the compile result is already known.
3. The failing expression is reachable under the analyzer's current control-flow model.
4. The operands or object state are statically known.
5. Error handling state is accounted for.
6. The result does not depend on workbook state, locale, references, add-ins, user input, file system, active sheet, or external COM registration.
```

If any of those are not true, downgrade to:

```text
probable-runtime-warning
host-warning
reference-resolution-error
needs-verification
```

---

## The Main Boundary

This file exists to prevent XLIDE from confusing four different things:

```text
what parses
what compiles
what runs
what should be warned about
```

The oracle decides the first three. Analyzer policy decides the fourth.
