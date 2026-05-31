# Excel VBA Realtime Syntax Linting Test Corpus

**Purpose:** Provide an LLM agent with a structured test plan and snippet corpus for validating realtime syntax linting in a VS Code extension for Excel VBA.

**Target host:** Excel VBA.

**Primary grammar oracle:** Microsoft **[MS-VBAL]: VBA Language Specification**. The agent must verify grammar, lexical rules, statement forms, declaration rules, and ambiguous cases against the current Microsoft specification before changing parser or linter behavior.

Official reference landing page: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

Current published version observed during creation of this test plan: **MS-VBAL 2.4, 2025-05-20**. Recheck this before implementation.

---

## Instructions for the LLM Agent

You are helping implement and validate realtime syntax linting for Excel VBA inside a VS Code extension.

Do **not** treat this corpus as the final authority. Treat it as a practical coverage scaffold. For every grammar-sensitive behavior, verify against **MS-VBAL**. For Excel host object examples, distinguish between **VBA syntax validity** and **Excel object model semantic validity**.

The linter should support three diagnostic levels:

1. **Lexical / parse error**: cannot be valid VBA syntax.
2. **Semantic compile error**: syntactically shaped like VBA but invalid by declaration, scope, module type, or language rules.
3. **Host-aware warning**: likely invalid or suspicious in Excel, but not a pure VBA syntax failure.

Realtime behavior matters. While the user is typing, the linter should avoid noisy false positives for incomplete but plausible constructs. Prefer delayed diagnostics for open blocks, unfinished string literals on the current line, and partial member access such as `ThisWorkbook.`.

---

## Required Test Harness Shape

Create a data-driven test harness with this shape:

```jsonc
{
  "id": "CTRL_IF_001",
  "title": "Multiline If / ElseIf / Else block",
  "moduleKind": "standard",
  "expected": "valid",
  "diagnostics": [],
  "snippet": "..."
}
```

Use these allowed values:

```text
moduleKind: standard | class | worksheet | workbook | userform | any
expected: valid | invalid | incomplete | warning
phase: lexical | parser | declaration | scope | host | realtime
```

For every failing test, include:

```jsonc
{
  "code": "VBA_PARSE_MISSING_END_IF",
  "severity": "error",
  "line": 4,
  "messageContains": "End If"
}
```

Line numbers should be stable and 1-based.

---

## Realtime Linting Rules

The agent must implement realtime tolerance before strict final validation.

During active typing, the following should usually be `incomplete`, not `invalid`:

```vba
Sub Demo()
```

```vba
If x > 0 Then
```

```vba
For i = 1 To 10
```

```vba
With ThisWorkbook.Worksheets(1)
```

```vba
ThisWorkbook.
```

```vba
Range("A1").
```

```vba
Debug.Print "unfinished
```

The same snippets may become invalid when the user leaves the construct stale, saves the file, runs full lint, or the parser reaches EOF with no plausible continuation.

Recommended modes:

```text
onType: tolerant parser, low-noise diagnostics
onSave: strict parser and declaration validation
onDemand: full syntax + semantic + host-aware validation
```

---

# Coverage Matrix

## 1. Module-Level Boilerplate

### MOD_001 valid: Option statements

```vba
Option Explicit
Option Compare Binary
Option Base 1
Option Private Module

Public Const APP_NAME As String = "XLIDE"
Private m_Count As Long
```

Expected: valid in a standard module.

### MOD_002 invalid: Option statement after declaration

```vba
Private m_Count As Long
Option Explicit
```

Expected: invalid. Option statements must appear before declarations/procedures.

### MOD_003 valid: Attribute lines from exported modules

```vba
Attribute VB_Name = "Module1"
Attribute VB_Description = "Test module"
Option Explicit

Public Sub Demo()
End Sub
```

Expected: valid or ignored as metadata. The lexer must not choke on `Attribute` lines in exported `.bas`, `.cls`, or `.frm` text.

### MOD_004 warning: Missing Option Explicit

```vba
Public Sub Demo()
    x = 1
End Sub
```

Expected: warning, not syntax error, unless project policy requires `Option Explicit`.

---

## 2. Comments, Labels, Colons, and Line Separators

### LEX_001 valid: Apostrophe comments

```vba
Public Sub Demo()
    ' This is a comment
    Debug.Print 1 ' trailing comment
End Sub
```

### LEX_002 valid: Rem comment

```vba
Public Sub Demo()
    Rem This is a comment
    Debug.Print 1
End Sub
```

### LEX_003 valid: Colon-separated statements

```vba
Public Sub Demo()
    Dim x As Long: x = 1: Debug.Print x
End Sub
```

### LEX_004 valid: Line label and GoTo

```vba
Public Sub Demo()
StartHere:
    Debug.Print "start"
    GoTo Done
Done:
End Sub
```

### LEX_005 valid: Numeric line labels

```vba
Public Sub Demo()
10  Debug.Print "line 10"
20  GoTo 40
30  Debug.Print "skipped"
40  Debug.Print "done"
End Sub
```

### LEX_006 invalid: Bad label position

```vba
Public Sub Demo()
    Debug.Print 1
    MyLabel: Dim x As Long
End Sub
```

Expected: verify against MS-VBAL. Some parsers incorrectly mishandle labels combined with colon-separated statements.

---

## 3. Line Continuation

### CONT_001 valid: Explicit line continuation

```vba
Public Sub Demo()
    Debug.Print "hello " & _
                "world"
End Sub
```

### CONT_002 valid: Continued procedure call

```vba
Public Sub Demo()
    MsgBox _
        Prompt:="Hello", _
        Buttons:=vbInformation, _
        Title:="XLIDE"
End Sub
```

### CONT_003 invalid: Comment after underscore

```vba
Public Sub Demo()
    Debug.Print "hello" & _ ' comment
                "world"
End Sub
```

Expected: invalid. VBA line continuation has strict trailing character behavior.

### CONT_004 invalid: Missing whitespace before underscore

```vba
Public Sub Demo()
    Debug.Print "hello" &_
                "world"
End Sub
```

Expected: invalid or warning depending on exact lexical implementation. Verify against MS-VBAL.

### CONT_005 valid: Long Excel chain continuation

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Sheet1") _
        .Range("A1") _
        .Resize(10, 3) _
        .Value = 42
End Sub
```

Expected: valid. Member access can continue on subsequent lines after explicit continuation.

---

## 4. String, Date, and Numeric Literals

### LIT_001 valid: Escaped quote by doubling

```vba
Public Sub Demo()
    Debug.Print "He said ""hello""."
End Sub
```

### LIT_002 invalid: Unterminated string

```vba
Public Sub Demo()
    Debug.Print "hello
End Sub
```

Expected during typing: incomplete. Expected on save: invalid.

### LIT_003 valid: Date literals

```vba
Public Sub Demo()
    Dim d As Date
    d = #1/31/2026#
    Debug.Print d
End Sub
```

### LIT_004 valid: Hex and octal literals

```vba
Public Sub Demo()
    Debug.Print &HFF
    Debug.Print &O77
End Sub
```

### LIT_005 valid: Type-declaration suffixes

```vba
Public Sub Demo()
    Dim a%, b&, c!, d#, e@, f$
    a = 1%
    b = 2&
    c = 3!
    d = 4#
    e = 5@
    f = "text"
End Sub
```

### LIT_006 warning: Ambiguous date format

```vba
Public Sub Demo()
    Debug.Print #3/4/2026#
End Sub
```

Expected: syntactically valid, optional host/style warning only.

---

## 5. Declarations and Type Clauses

### DECL_001 valid: Dim with multiple declarators

```vba
Public Sub Demo()
    Dim a As Long, b As String, c As Variant
End Sub
```

### DECL_002 warning: Only final variable typed

```vba
Public Sub Demo()
    Dim a, b, c As Long
End Sub
```

Expected: syntactically valid. Warn that `a` and `b` are Variants.

### DECL_003 valid: Object variable and Set assignment

```vba
Public Sub Demo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Debug.Print ws.Name
End Sub
```

### DECL_004 invalid or semantic error: Object assignment without Set

```vba
Public Sub Demo()
    Dim ws As Worksheet
    ws = ThisWorkbook.Worksheets(1)
End Sub
```

Expected: syntactically shaped but semantic compile error in VBA object assignment.

### DECL_005 valid: New object declaration

```vba
Public Sub Demo()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
End Sub
```

### DECL_006 valid: Static local variable

```vba
Public Sub Demo()
    Static counter As Long
    counter = counter + 1
End Sub
```

### DECL_007 valid: Module-level constants

```vba
Option Explicit

Private Const MAX_ROWS As Long = 1000
Public Const DEFAULT_NAME As String = "Sheet1"
```

### DECL_008 invalid: Assignment in Dim statement

```vba
Public Sub Demo()
    Dim x As Long = 1
End Sub
```

Expected: invalid in VBA. Do not accidentally apply VB.NET grammar.

---

## 6. Procedures

### PROC_001 valid: Basic Sub

```vba
Public Sub Demo()
    Debug.Print "ok"
End Sub
```

### PROC_002 valid: Function with return assignment

```vba
Public Function Add(ByVal a As Long, ByVal b As Long) As Long
    Add = a + b
End Function
```

### PROC_003 valid: Optional parameter and default

```vba
Public Function Greeting(Optional ByVal name As String = "world") As String
    Greeting = "hello " & name
End Function
```

### PROC_004 valid: ParamArray

```vba
Public Function SumAll(ParamArray values() As Variant) As Double
    Dim i As Long
    For i = LBound(values) To UBound(values)
        SumAll = SumAll + CDbl(values(i))
    Next i
End Function
```

### PROC_005 invalid: Required parameter after Optional

```vba
Public Sub Demo(Optional ByVal x As Long = 1, ByVal y As Long)
End Sub
```

### PROC_006 invalid: ParamArray not last

```vba
Public Sub Demo(ParamArray values() As Variant, ByVal x As Long)
End Sub
```

### PROC_007 valid: ByRef default

```vba
Public Sub Increment(value As Long)
    value = value + 1
End Sub
```

Expected: valid. Optional warning that default parameter passing is ByRef.

### PROC_008 invalid: Missing End Sub

```vba
Public Sub Demo()
    Debug.Print "ok"
```

Expected during typing: incomplete. Expected on save: invalid.

---

## 7. Properties

### PROP_001 valid: Property Get / Let pair

```vba
Private m_Name As String

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = value
End Property
```

### PROP_002 valid: Object Property Set

```vba
Private m_Worksheet As Worksheet

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = m_Worksheet
End Property

Public Property Set TargetSheet(ByVal value As Worksheet)
    Set m_Worksheet = value
End Property
```

### PROP_003 invalid: Property Let using object assignment target

```vba
Private m_Worksheet As Worksheet

Public Property Let TargetSheet(ByVal value As Worksheet)
    Set m_Worksheet = value
End Property
```

Expected: likely semantic invalid. Verify exact rule against MS-VBAL and host compile behavior.

---

## 8. Control Flow: If, Select, Loops

### CTRL_IF_001 valid: Multiline If / ElseIf / Else

```vba
Public Sub Demo(ByVal x As Long)
    If x > 10 Then
        Debug.Print "large"
    ElseIf x > 0 Then
        Debug.Print "positive"
    Else
        Debug.Print "other"
    End If
End Sub
```

### CTRL_IF_002 valid: Single-line If

```vba
Public Sub Demo(ByVal x As Long)
    If x > 0 Then Debug.Print "positive" Else Debug.Print "not positive"
End Sub
```

### CTRL_IF_003 invalid: Missing Then

```vba
Public Sub Demo(ByVal x As Long)
    If x > 0
        Debug.Print x
    End If
End Sub
```

### CTRL_IF_004 invalid: Else without If

```vba
Public Sub Demo()
    Else
        Debug.Print "bad"
End Sub
```

### CTRL_SELECT_001 valid: Select Case forms

```vba
Public Sub Demo(ByVal x As Long)
    Select Case x
        Case 1
            Debug.Print "one"
        Case 2 To 5
            Debug.Print "range"
        Case Is > 5
            Debug.Print "large"
        Case Else
            Debug.Print "other"
    End Select
End Sub
```

### CTRL_SELECT_002 invalid: Missing End Select

```vba
Public Sub Demo(ByVal x As Long)
    Select Case x
        Case 1
            Debug.Print "one"
End Sub
```

### CTRL_FOR_001 valid: For / Next

```vba
Public Sub Demo()
    Dim i As Long
    For i = 1 To 10 Step 2
        Debug.Print i
    Next i
End Sub
```

### CTRL_FOR_002 valid: For Each

```vba
Public Sub Demo()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
```

### CTRL_FOR_003 invalid: Missing Next

```vba
Public Sub Demo()
    Dim i As Long
    For i = 1 To 10
        Debug.Print i
End Sub
```

### CTRL_DO_001 valid: Do While / Loop

```vba
Public Sub Demo()
    Dim i As Long
    Do While i < 10
        i = i + 1
    Loop
End Sub
```

### CTRL_DO_002 valid: Do / Loop Until

```vba
Public Sub Demo()
    Dim i As Long
    Do
        i = i + 1
    Loop Until i >= 10
End Sub
```

### CTRL_WHILE_001 valid: While / Wend

```vba
Public Sub Demo()
    Dim i As Long
    While i < 10
        i = i + 1
    Wend
End Sub
```

---

## 9. Error Handling

### ERR_001 valid: On Error GoTo label

```vba
Public Sub Demo()
    On Error GoTo CleanFail
    Debug.Print 1 / 0
    Exit Sub
CleanFail:
    Debug.Print Err.Number, Err.Description
End Sub
```

### ERR_002 valid: On Error Resume Next

```vba
Public Sub Demo()
    On Error Resume Next
    Debug.Print ThisWorkbook.Worksheets("Missing").Name
    On Error GoTo 0
End Sub
```

### ERR_003 valid: Resume forms

```vba
Public Sub Demo()
    On Error GoTo Handler
    Err.Raise 5
    Exit Sub
Handler:
    Resume Next
End Sub
```

### ERR_004 valid legacy: GoSub / Return

```vba
Public Sub Demo()
    GoSub Worker
    Exit Sub
Worker:
    Debug.Print "worker"
    Return
End Sub
```

Expected: valid legacy syntax. Optional style warning only.

---

## 10. Expressions and Operators

### EXPR_001 valid: Arithmetic and precedence

```vba
Public Sub Demo()
    Debug.Print 1 + 2 * 3 ^ 2
End Sub
```

### EXPR_002 valid: String concatenation

```vba
Public Sub Demo()
    Debug.Print "A" & "B"
End Sub
```

### EXPR_003 valid: Comparison and logical operators

```vba
Public Sub Demo(ByVal a As Boolean, ByVal b As Boolean)
    Debug.Print Not a Or b And a Xor b
End Sub
```

### EXPR_004 valid: Like operator

```vba
Public Sub Demo()
    Debug.Print "ABC123" Like "[A-Z]*#"
End Sub
```

### EXPR_005 valid: Is operator

```vba
Public Sub Demo()
    Dim ws As Worksheet
    Set ws = Nothing
    Debug.Print ws Is Nothing
End Sub
```

### EXPR_006 valid: TypeOf Is

```vba
Public Sub Demo(ByVal obj As Object)
    If TypeOf obj Is Worksheet Then
        Debug.Print "worksheet"
    End If
End Sub
```

### EXPR_007 valid: Integer division and Mod

```vba
Public Sub Demo()
    Debug.Print 7 \ 2
    Debug.Print 7 Mod 2
End Sub
```

### EXPR_008 valid: Eqv and Imp

```vba
Public Sub Demo(ByVal a As Boolean, ByVal b As Boolean)
    Debug.Print a Eqv b
    Debug.Print a Imp b
End Sub
```

---

## 11. Arrays

### ARR_001 valid: Static array

```vba
Public Sub Demo()
    Dim values(1 To 10) As Long
    values(1) = 42
End Sub
```

### ARR_002 valid: Dynamic array and ReDim

```vba
Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10)
    ReDim Preserve values(1 To 20)
End Sub
```

### ARR_003 valid: Variant array from Array function

```vba
Public Sub Demo()
    Dim values As Variant
    values = Array("a", "b", "c")
    Debug.Print values(0)
End Sub
```

### ARR_004 invalid: Preserve changing lower dimension count

```vba
Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10, 1 To 2)
    ReDim Preserve values(1 To 20, 1 To 2)
End Sub
```

Expected: syntactically valid but semantic compile/runtime issue depending exact rule. Verify and classify carefully.

---

## 12. User-Defined Types and Enums

### TYPE_001 valid: Private Type

```vba
Private Type Person
    Id As Long
    Name As String
End Type

Public Sub Demo()
    Dim p As Person
    p.Id = 1
    p.Name = "Ada"
End Sub
```

### TYPE_002 valid: Enum

```vba
Public Enum JobStatus
    JobPending = 0
    JobRunning = 1
    JobDone = 2
End Enum

Public Sub Demo()
    Dim status As JobStatus
    status = JobRunning
End Sub
```

### TYPE_003 invalid: Type inside procedure

```vba
Public Sub Demo()
    Type Person
        Id As Long
    End Type
End Sub
```

Expected: invalid. Type declarations are module-level.

---

## 13. With Blocks and Member Access

### WITH_001 valid: Basic With

```vba
Public Sub Demo()
    With ThisWorkbook.Worksheets("Sheet1").Range("A1")
        .Value = "Hello"
        .Font.Bold = True
    End With
End Sub
```

### WITH_002 invalid: Dot member outside With or chain

```vba
Public Sub Demo()
    .Value = 123
End Sub
```

Expected: invalid or semantic error depending parser architecture. A leading dot statement requires a With context.

### WITH_003 incomplete: Member access after dot

```vba
Public Sub Demo()
    ThisWorkbook.
End Sub
```

Expected during typing: incomplete, not hard invalid until strict mode.

### WITH_004 valid: Nested With

```vba
Public Sub Demo()
    With ThisWorkbook.Worksheets(1)
        With .Range("A1")
            .Value = 1
        End With
    End With
End Sub
```

---

## 14. Excel Host Object Patterns

These are mainly syntax and linter-noise tests. The linter must not require complete Excel type inference before recognizing common valid chains.

### XL_001 valid: Workbook / worksheet / range chain

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Data").Range("A1").Value2 = "ok"
End Sub
```

### XL_002 valid: Range Resize assignment

```vba
Public Sub Demo()
    Dim data(1 To 2, 1 To 2) As Variant
    data(1, 1) = "A"
    data(1, 2) = "B"
    data(2, 1) = 1
    data(2, 2) = 2

    ThisWorkbook.Worksheets("Data").Range("A1").Resize(2, 2).Value = data
End Sub
```

### XL_003 valid: Named arguments

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Data").Range("A1:B10").Sort _
        Key1:=ThisWorkbook.Worksheets("Data").Range("A1"), _
        Order1:=xlAscending, _
        Header:=xlYes
End Sub
```

### XL_004 valid: Worksheet event in sheet module

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A:A")) Is Nothing Then
        Debug.Print Target.Address
    End If
End Sub
```

Expected: valid in worksheet module. Standard module should not treat it as syntax error, but may warn as host-context mismatch.

### XL_005 valid: Workbook event in ThisWorkbook module

```vba
Private Sub Workbook_Open()
    MsgBox "Opened"
End Sub
```

Expected: valid in workbook module. Standard module may warn as host-context mismatch.

### XL_006 warning: Unqualified Range

```vba
Public Sub Demo()
    Range("A1").Value = 1
End Sub
```

Expected: syntactically valid. Optional host-aware warning: unqualified `Range` depends on active sheet.

### XL_007 valid: Application.WorksheetFunction

```vba
Public Sub Demo()
    Debug.Print Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(1).Range("A1:A10"))
End Sub
```

### XL_008 valid: Evaluate with formula string

```vba
Public Sub Demo()
    Debug.Print Application.Evaluate("SUM(1,2,3)")
End Sub
```

Expected: valid. Do not parse Excel formula strings as VBA.

---

## 15. Preprocessor Directives and Conditional Compilation

### PP_001 valid: VBA7 / Win64 guards

```vba
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
```

### PP_002 valid: Conditional compilation inside procedure

```vba
Public Sub Demo()
#If DEBUG_MODE Then
    Debug.Print "debug"
#Else
    Debug.Print "release"
#End If
End Sub
```

### PP_003 invalid: Missing #End If

```vba
#If VBA7 Then
Public Sub Demo()
End Sub
```

Expected during typing: incomplete. Expected on save: invalid.

### PP_004 valid: #Const

```vba
#Const DEBUG_MODE = True

Public Sub Demo()
#If DEBUG_MODE Then
    Debug.Print "debug"
#End If
End Sub
```

---

## 16. Declare Statements and 64-bit Compatibility

### API_001 valid: PtrSafe Declare

```vba
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
```

### API_002 valid: LongPtr usage

```vba
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
#End If
```

### API_003 warning: Declare without PtrSafe under VBA7

```vba
Private Declare Function GetTickCount Lib "kernel32" () As Long
```

Expected: warning in modern Office/VBA7 context, not pure syntax error if compatibility mode is configurable.

---

## 17. Classes, Events, Implements, and WithEvents

### CLS_001 valid class: Private state and public method

```vba
Option Explicit

Private m_Name As String

Public Sub Initialize(ByVal name As String)
    m_Name = name
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property
```

### CLS_002 valid class: Event and RaiseEvent

```vba
Option Explicit

Public Event Changed(ByVal propertyName As String)

Public Property Let Name(ByVal value As String)
    RaiseEvent Changed("Name")
End Property
```

Expected: valid in class module. Standard module should classify carefully.

### CLS_003 valid class: WithEvents field

```vba
Option Explicit

Private WithEvents m_App As Application

Private Sub m_App_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print Wb.Name
End Sub
```

Expected: valid in class module.

### CLS_004 valid class: Implements

```vba
Option Explicit

Implements IDisposableLike

Private Sub IDisposableLike_Dispose()
    Debug.Print "dispose"
End Sub
```

Expected: syntactically valid in a class module if interface exists. Semantic resolution of interface may be external to syntax linter.

### CLS_005 invalid standard module: WithEvents in standard module

```vba
Option Explicit

Private WithEvents m_App As Application
```

Expected: semantic/module-context error if moduleKind is standard.

---

## 18. UserForm Patterns

### FORM_001 valid: UserForm event

```vba
Private Sub UserForm_Initialize()
    Me.Caption = "Demo"
End Sub
```

Expected: valid in userform module. Standard module should not hard syntax-fail, but may host-context warn.

### FORM_002 valid: Control event

```vba
Private Sub CommandButton1_Click()
    MsgBox "clicked"
End Sub
```

Expected: valid in userform module.

---

## 19. Ambiguous and Dangerous Parser Traps

### TRAP_001 valid: Keyword-like identifiers in legal contexts

```vba
Public Sub Demo()
    Dim [Date] As Date
    Dim [Name] As String
    [Name] = "Ada"
    Debug.Print [Name], [Date]
End Sub
```

Expected: bracketed identifiers valid. The lexer must not treat them as keywords.

### TRAP_002 valid: Default member call without Call keyword

```vba
Public Sub Demo()
    MsgBox "hello", vbInformation, "Title"
End Sub
```

### TRAP_003 valid: Call keyword requires parentheses

```vba
Public Sub Demo()
    Call MsgBox("hello", vbInformation, "Title")
End Sub
```

### TRAP_004 invalid: Call keyword without parentheses

```vba
Public Sub Demo()
    Call MsgBox "hello"
End Sub
```

### TRAP_005 valid: Let keyword optional/legacy

```vba
Public Sub Demo()
    Let x = 1
End Sub
```

Expected: syntactically valid. With Option Explicit this becomes undeclared variable error.

### TRAP_006 valid: Empty statement lines

```vba
Public Sub Demo()


    Debug.Print "ok"

End Sub
```

### TRAP_007 valid: Colon after If single-line body

```vba
Public Sub Demo(ByVal x As Long)
    If x > 0 Then Debug.Print "a": Debug.Print "b"
End Sub
```

Expected: verify exact association rules against MS-VBAL.

---

## 20. Invalid Corpus: High-Value Negative Tests

### BAD_001 missing block closer

```vba
Public Sub Demo()
    If True Then
        Debug.Print "x"
End Sub
```

Expected: invalid, missing `End If`.

### BAD_002 mismatched block closer

```vba
Public Sub Demo()
    For i = 1 To 10
        Debug.Print i
    End If
End Sub
```

Expected: invalid, `End If` without `If`; missing `Next`.

### BAD_003 duplicate procedure end

```vba
Public Sub Demo()
End Sub
End Sub
```

Expected: invalid.

### BAD_004 procedure inside procedure

```vba
Public Sub Outer()
    Public Sub Inner()
    End Sub
End Sub
```

Expected: invalid.

### BAD_005 invalid assignment target

```vba
Public Sub Demo()
    1 = x
End Sub
```

Expected: invalid.

### BAD_006 invalid chained comparison assumption

```vba
Public Sub Demo()
    If 1 < 2 < 3 Then Debug.Print "bad idea"
End Sub
```

Expected: syntactically valid in VBA expression grammar, but warning. It does not mean mathematical chained comparison.

### BAD_007 malformed named argument

```vba
Public Sub Demo()
    MsgBox Prompt = "Hello"
End Sub
```

Expected: parse according to expression rules; likely not intended named argument. Warn or error depending grammar position.

### BAD_008 unterminated block comment misconception

```vba
Public Sub Demo()
    /* not a VBA comment */
End Sub
```

Expected: invalid. VBA has no C-style block comments.

### BAD_009 VB.NET syntax not valid VBA

```vba
Public Sub Demo()
    Dim x As Integer = 1
    Console.WriteLine(x)
End Sub
```

Expected: invalid / host warning. Do not accept VB.NET initialization syntax.

### BAD_010 JavaScript/Python syntax contamination

```vba
Public Sub Demo()
    If x == 1 Then
        Debug.Print("x")
    End If
End Sub
```

Expected: invalid because `==` is not VBA equality.

---

# Incremental Typing Test Sequences

The realtime parser must be tested as a sequence, not only as complete files.

## RT_001 Sub block typing

Steps:

```text
1. "Sub Demo"
2. "Sub Demo()"
3. "Sub Demo()\n"
4. "Sub Demo()\n    Debug.Print 1"
5. "Sub Demo()\n    Debug.Print 1\nEnd Sub"
```

Expected:

```text
1. incomplete, not invalid
2. incomplete, not invalid
3. incomplete, missing End Sub suppressed or low severity
4. incomplete, missing End Sub suppressed or low severity
5. valid
```

## RT_002 Dot member typing

Steps:

```text
1. "ThisWorkbook"
2. "ThisWorkbook."
3. "ThisWorkbook.Worksheets"
4. "ThisWorkbook.Worksheets("
5. "ThisWorkbook.Worksheets(1).Range(\"A1\").Value = 1"
```

Expected:

```text
1. valid expression fragment when in statement context may be incomplete
2. incomplete, trigger completion list, no hard error
3. valid/incomplete depending expression context
4. incomplete, missing close paren suppressed while typing
5. valid
```

## RT_003 String typing

Steps:

```text
1. "Debug.Print \""
2. "Debug.Print \"hello"
3. "Debug.Print \"hello\""
```

Expected:

```text
1. incomplete
2. incomplete while current line active; invalid on save if still unterminated
3. valid
```

## RT_004 If block typing

Steps:

```text
1. "If x > 0"
2. "If x > 0 Then"
3. "If x > 0 Then\n    Debug.Print x"
4. "If x > 0 Then\n    Debug.Print x\nEnd If"
```

Expected:

```text
1. incomplete or invalid depending cursor position; avoid noisy error while active
2. incomplete
3. incomplete
4. valid
```

---

# Diagnostics the Agent Should Implement

Use stable diagnostic codes. Suggested minimum set:

```text
VBA_LEX_UNTERMINATED_STRING
VBA_LEX_INVALID_LINE_CONTINUATION
VBA_LEX_UNKNOWN_TOKEN
VBA_PARSE_EXPECTED_THEN
VBA_PARSE_EXPECTED_END_SUB
VBA_PARSE_EXPECTED_END_FUNCTION
VBA_PARSE_EXPECTED_END_IF
VBA_PARSE_EXPECTED_END_SELECT
VBA_PARSE_EXPECTED_NEXT
VBA_PARSE_UNEXPECTED_END_IF
VBA_PARSE_UNEXPECTED_NEXT
VBA_PARSE_PROCEDURE_NESTED
VBA_PARSE_INVALID_DECLARATION_CONTEXT
VBA_SEM_OPTION_EXPLICIT_UNDECLARED_IDENTIFIER
VBA_SEM_OBJECT_ASSIGNMENT_REQUIRES_SET
VBA_SEM_DUPLICATE_DECLARATION
VBA_SEM_REQUIRED_PARAM_AFTER_OPTIONAL
VBA_SEM_PARAMARRAY_NOT_LAST
VBA_SEM_WITHOUT_WITH_CONTEXT
VBA_HOST_UNQUALIFIED_RANGE
VBA_HOST_EVENT_CONTEXT_MISMATCH
VBA_STYLE_IMPLICIT_VARIANT
VBA_STYLE_LEGACY_GOSUB
VBA_STYLE_DECLARE_PTRSAFE_REQUIRED
```

---

# Suggested Test File Layout

```text
tests/
  fixtures/
    valid/
      standard_module.bas
      excel_patterns.bas
      class_module.cls
      worksheet_module.cls
      userform_module.frm
    invalid/
      missing_end_if.bas
      bad_line_continuation.bas
      vbnet_syntax.bas
    realtime/
      incremental_sub.json
      incremental_member_access.json
  expected/
    valid.snap.json
    invalid.snap.json
    realtime.snap.json
```

---

# Agent Implementation Priorities

## Phase 1: Fast lexical/parser coverage

Implement enough parsing to correctly handle:

```text
comments
strings
date literals
numeric literals
line continuations
colon-separated statements
labels
procedure declarations
block matching
If / Select / For / Do / While
With blocks
basic expressions
```

This phase should catch the majority of live syntax errors without requiring full type inference.

## Phase 2: Declaration and scope validation

Add:

```text
Option Explicit checks
local/module scope symbols
duplicate declaration checks
procedure parameter validation
object assignment requires Set
invalid declaration context
module-kind-sensitive checks
```

## Phase 3: Excel host-aware diagnostics

Add warnings for:

```text
unqualified Range / Cells / Rows / Columns
worksheet events in non-worksheet modules
workbook events in non-workbook modules
UserForm events in non-userform modules
potentially missing PtrSafe in VBA7 mode
suspicious default member usage
```

## Phase 4: Completion-aware parse recovery

For VS Code experience, implement parse recovery around:

```text
trailing dot
unfinished call argument list
unfinished string on active line
unfinished block on active line
partially typed keyword
unfinished line continuation
```

Do not punish the user for being halfway through a correct sentence.

---

# Quality Gate Before Shipping

The extension should pass all of these before public release:

1. No hard errors for valid Excel macro patterns.
2. No VB.NET syntax accepted as valid VBA.
3. Missing block terminators are caught on save/full lint.
4. Realtime member access like `ThisWorkbook.` triggers completion without red-squiggle spam.
5. `Dim a, b, c As Long` is not misreported as syntax invalid.
6. `Set` rules are handled for object assignment.
7. `Call` syntax is distinguished from normal argument-list syntax.
8. `Attribute VB_Name = "..."` does not poison exported-module parsing.
9. Conditional compilation blocks are parsed independently enough to avoid false global corruption.
10. Excel-specific event procedure names are recognized as valid procedure declarations.

---

# Agent Warnings

Avoid these common mistakes:

```text
Do not parse VBA as VB.NET.
Do not parse Excel formula strings as VBA.
Do not require Excel object model resolution for pure syntax validity.
Do not report incomplete active-line code as hard invalid too early.
Do not assume every identifier followed by parentheses is a function call; arrays and default members exist.
Do not assume every assignment without Set is valid; object assignment is special.
Do not break exported modules because of Attribute lines.
Do not ignore colon-separated statements.
Do not ignore labels and numeric line numbers.
Do not collapse all invalid states into one generic parser error.
```

---

# Optional Fuzzing Seeds

Use these as mutation seeds:

```vba
Public Sub FuzzSeed()
    Dim ws As Worksheet, i As Long, values() As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    ReDim values(1 To 10, 1 To 3)

    For i = LBound(values, 1) To UBound(values, 1)
        If i Mod 2 = 0 Then
            ws.Cells(i, 1).Value = "even"
        Else
            ws.Cells(i, 1).Value = "odd"
        End If
    Next i
End Sub
```

Mutation targets:

```text
remove End If
remove Next
swap End If and Next
delete closing quote
insert bad line continuation
replace = with ==
insert VB.NET Dim initializer
move Option Explicit below procedure
add trailing dot
add colon-separated statement after label
change Set assignment to normal assignment
```

---

# Final Note

The goal is not merely to catch errors. The goal is to make VBA feel native in VS Code without betraying how VBA actually works. A good realtime linter should feel calm while typing and strict when it matters.
