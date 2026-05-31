# Excel VBA Realtime Linting Addendum: Additional Edge-Case Snippets

**Purpose:** Extend the primary Excel VBA realtime linting corpus with high-value edge cases that commonly break parsers, completion engines, and syntax highlighters.

**Audience:** LLM agent implementing or validating a VS Code extension for Excel VBA.

**Non-negotiable rule:** Verify grammar-sensitive behavior against Microsoft **MS-VBAL** before changing parser behavior. Do not infer VBA rules from VB.NET, VBScript, JavaScript, Python, or generic BASIC.

Official MS-VBAL reference: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

---

## How to Use This Addendum

Add these cases to the same fixture/test harness used by the main corpus.

Each test should include:

```jsonc
{
  "id": "EDGE_001",
  "title": "Short description",
  "moduleKind": "standard",
  "expected": "valid",
  "phase": "parser",
  "diagnostics": [],
  "snippet": "..."
}
```

Allowed values:

```text
moduleKind: standard | class | worksheet | workbook | userform | any
expected: valid | invalid | incomplete | warning
phase: lexical | parser | declaration | scope | host | realtime
```

Classification guidance:

```text
valid      = valid VBA syntax and acceptable for the declared module kind
invalid    = should produce an error in strict/full lint mode
incomplete = should not hard-error while user is actively typing
warning    = syntactically valid, but suspicious, host-specific, or style-risky
```

---

# A. Declarations, Options, and Module Header Edge Cases

## OPT_005 valid: Option Compare Text

```vba
Option Explicit
Option Compare Text

Public Sub Demo()
    Debug.Print "abc" = "ABC"
End Sub
```

Expected: valid.

## OPT_006 valid: Option Base 0

```vba
Option Explicit
Option Base 0

Public Sub Demo()
    Dim values(3) As Long
    Debug.Print LBound(values), UBound(values)
End Sub
```

Expected: valid.

## OPT_007 invalid or host-specific: Option Compare Database in Excel

```vba
Option Explicit
Option Compare Database

Public Sub Demo()
End Sub
```

Expected: for Excel VBA, likely invalid or host-context warning. Verify exact host behavior. Do not blindly accept Access-specific patterns as Excel-valid.

## OPT_008 invalid: Option Base after procedure

```vba
Public Sub Demo()
End Sub

Option Base 1
```

Expected: invalid. Option statements belong in the declaration section.

## OPT_009 warning or invalid: duplicate Option Explicit

```vba
Option Explicit
Option Explicit

Public Sub Demo()
End Sub
```

Expected: verify against host compile behavior. At minimum, parser must not crash.

## ATTR_001 valid exported module: module attributes before body

```vba
Attribute VB_Name = "Module1"
Attribute VB_Description = "Parser test"
Option Explicit

Public Sub Demo()
End Sub
```

Expected: valid when parsing exported module text. The linter may ignore attributes after recording metadata.

## ATTR_002 valid exported member attribute placement

```vba
Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
Attribute Demo.VB_Description = "Demo macro"
    Debug.Print "ok"
End Sub
```

Expected: valid for exported-module parsing if MS-VBAL/export grammar allows this placement. A naive parser often rejects `Attribute` inside a procedure body; the agent must distinguish exported metadata from executable source.

## ATTR_003 invalid live editor source: user-typed Attribute in body

```vba
Public Sub Demo()
    Attribute Demo.VB_Description = "Bad here"
End Sub
```

Expected: invalid in normal editor mode unless parsing exported-file metadata mode is enabled.

## DEF_001 valid legacy: DefLng type range

```vba
Option Explicit
DefLng A-Z

Public Sub Demo()
    Dim x
    x = 1
    Debug.Print TypeName(x)
End Sub
```

Expected: syntactically valid legacy VBA. Optional style warning. If `Option Explicit` is enabled, `Dim x` is still declared.

## DEF_002 valid legacy: multiple Def type ranges

```vba
DefInt I-N
DefStr S

Public Sub Demo()
    Dim IndexValue
    Dim StatusText
End Sub
```

Expected: valid. Optional style warning.

## DEF_003 invalid: malformed Def range

```vba
DefLng Z-A

Public Sub Demo()
End Sub
```

Expected: verify exact rule. Parser must at least detect malformed declaration if grammar disallows descending range.

---

# B. Procedure Declaration and Modifier Edge Cases

## PROC_009 valid: Static procedure

```vba
Public Static Sub Demo()
    Static counter As Long
    counter = counter + 1
    Debug.Print counter
End Sub
```

Expected: valid if modifier order is accepted. Verify exact allowed modifier order.

## PROC_010 invalid: duplicate access modifiers

```vba
Public Private Sub Demo()
End Sub
```

Expected: invalid.

## PROC_011 valid: Function name with type-declaration suffix

```vba
Public Function GetName$()
    GetName = "XLIDE"
End Function
```

Expected: valid legacy syntax. Optional style warning.

## PROC_012 invalid: suffix plus As clause conflict

```vba
Public Function GetName$() As String
    GetName = "XLIDE"
End Function
```

Expected: verify exact compile behavior. Likely invalid because the function type is declared twice.

## PROC_013 valid: Exit forms inside matching procedures

```vba
Public Sub DemoSub()
    Exit Sub
End Sub

Public Function DemoFunction() As Long
    Exit Function
    DemoFunction = 1
End Function

Public Property Get DemoProperty() As Long
    Exit Property
    DemoProperty = 1
End Property
```

Expected: valid.

## PROC_014 invalid: Exit Function inside Sub

```vba
Public Sub Demo()
    Exit Function
End Sub
```

Expected: semantic/compile error.

## PROC_015 valid: Empty procedure body

```vba
Public Sub Demo()
End Sub
```

Expected: valid.

## PROC_016 invalid: Procedure declaration missing parentheses

```vba
Public Sub Demo
End Sub
```

Expected: verify exact VBA behavior. If accepted by host, classify valid; if not, parser error. Do not guess.

## PROC_017 valid: Optional Variant without explicit default

```vba
Public Sub Demo(Optional value As Variant)
    If IsMissing(value) Then Debug.Print "missing"
End Sub
```

Expected: valid.

## PROC_018 invalid: IsMissing on non-Variant optional parameter

```vba
Public Sub Demo(Optional ByVal value As Long = 0)
    If IsMissing(value) Then Debug.Print "missing"
End Sub
```

Expected: syntactically valid but semantic warning/error candidate. `IsMissing` is meaningful for optional Variant parameters.

---

# C. Argument Lists, Calls, Parentheses, and Named Arguments

## CALL_001 valid: Sub call without Call and without parentheses

```vba
Public Sub Demo()
    MsgBox "hello", vbInformation, "Title"
End Sub
```

Expected: valid.

## CALL_002 valid: Call keyword with parentheses

```vba
Public Sub Demo()
    Call MsgBox("hello", vbInformation, "Title")
End Sub
```

Expected: valid.

## CALL_003 invalid: Call keyword without parentheses

```vba
Public Sub Demo()
    Call MsgBox "hello", vbInformation, "Title"
End Sub
```

Expected: invalid.

## CALL_004 parser trap: parenthesized argument changes meaning

```vba
Public Sub Demo()
    PrintValue (1)
End Sub

Private Sub PrintValue(ByRef x As Long)
    Debug.Print x
End Sub
```

Expected: syntactically valid. Optional semantic/style warning: parenthesized argument may force expression evaluation and can affect ByRef behavior.

## CALL_005 invalid: Multi-argument Sub call with parentheses and no Call

```vba
Public Sub Demo()
    MsgBox ("hello", vbInformation, "Title")
End Sub
```

Expected: invalid parse or semantic error. Do not accept JavaScript-style call syntax.

## CALL_006 valid: Function call in expression with parentheses

```vba
Public Sub Demo()
    Dim result As VbMsgBoxResult
    result = MsgBox("hello", vbYesNo, "Question")
End Sub
```

Expected: valid.

## CALL_007 valid: Named arguments with omitted optional arguments

```vba
Public Sub Demo()
    MsgBox Prompt:="hello", Title:="XLIDE"
End Sub
```

Expected: valid.

## CALL_008 valid: Empty positional argument slot

```vba
Public Sub Demo()
    Workbooks.Open Filename:="C:\Temp\Book1.xlsx", ReadOnly:=True
End Sub
```

Expected: valid. Host-aware warnings may apply for path existence, but not syntax.

## CALL_009 invalid: Named argument operator typo

```vba
Public Sub Demo()
    MsgBox Prompt = "hello"
End Sub
```

Expected: syntactically this may parse as an expression passed positionally, not as a named argument. The linter should warn that `:=` was probably intended.

---

# D. Identifier, Keyword, and Type Suffix Traps

## ID_001 valid: Bracketed keyword identifiers

```vba
Public Sub Demo()
    Dim [Sub] As String
    Dim [Function] As Long
    [Sub] = "not a keyword here"
    [Function] = 42
End Sub
```

Expected: valid. Lexer must preserve bracketed identifiers.

## ID_002 valid or host-specific: Unicode identifier

```vba
Public Sub Demo()
    Dim café As String
    café = "ok"
    Debug.Print café
End Sub
```

Expected: verify against MS-VBAL lexical rules and actual exported encoding behavior. Parser must be Unicode-aware where supported.

## ID_003 invalid: illegal identifier character

```vba
Public Sub Demo()
    Dim user-name As String
End Sub
```

Expected: invalid. This should not be parsed as a valid identifier.

## ID_004 valid: type-declaration characters on variables

```vba
Public Sub Demo()
    Dim total&, name$, price@, ratio#, flag%
    total = 1
    name = "XLIDE"
    price = 12.34@
    ratio = 0.5#
    flag = 1%
End Sub
```

Expected: valid legacy syntax.

## ID_005 invalid: type suffix plus conflicting As clause

```vba
Public Sub Demo()
    Dim name$ As String
End Sub
```

Expected: verify exact behavior. Likely invalid because the type is declared twice.

## ID_006 valid: Identifier beginning with underscore is not valid, but member can contain underscore

```vba
Public Sub Demo()
    Dim first_name As String
    first_name = "Ada"
End Sub
```

Expected: valid.

## ID_007 invalid: leading underscore local variable

```vba
Public Sub Demo()
    Dim _name As String
End Sub
```

Expected: verify exact identifier grammar. Many parsers over-accept this.

---

# E. Operators and Expression Edge Cases

## EXPR_009 valid: AddressOf callback argument

```vba
Public Sub Demo()
    Debug.Print ObjPtr(AddressOf Callback)
End Sub

Private Sub Callback()
End Sub
```

Expected: verify exact allowed contexts for `AddressOf`. If `ObjPtr(AddressOf ...)` is invalid in host compile, replace with a known API callback fixture. Parser must still recognize `AddressOf` as VBA syntax.

## EXPR_010 valid: Nothing assignment to object variable

```vba
Public Sub Demo()
    Dim ws As Worksheet
    Set ws = Nothing
End Sub
```

Expected: valid.

## EXPR_011 invalid: Set used with value type

```vba
Public Sub Demo()
    Dim x As Long
    Set x = 1
End Sub
```

Expected: semantic error.

## EXPR_012 valid: Object comparison with Is

```vba
Public Sub Demo()
    Dim a As Worksheet
    Dim b As Worksheet
    Set a = ThisWorkbook.Worksheets(1)
    Set b = ThisWorkbook.Worksheets(1)
    Debug.Print a Is b
End Sub
```

Expected: valid.

## EXPR_013 invalid: Is used with non-object values

```vba
Public Sub Demo()
    Debug.Print 1 Is 1
End Sub
```

Expected: semantic error.

## EXPR_014 valid: Like pattern with escaped bracket

```vba
Public Sub Demo()
    Debug.Print "A[1]" Like "A[[]#]"
End Sub
```

Expected: valid.

## EXPR_015 valid: chained member call with default property risk

```vba
Public Sub Demo()
    Debug.Print ThisWorkbook.Worksheets(1).Cells(1, 1)
End Sub
```

Expected: syntactically valid. Optional warning: implicit default member `.Value` may be intended.

---

# F. Block and Statement Separator Edge Cases

## SEP_001 valid: Multiple statements after Then

```vba
Public Sub Demo(ByVal x As Long)
    If x > 0 Then Debug.Print "positive": x = x + 1: Debug.Print x
End Sub
```

Expected: valid. Parser must associate colon-separated statements with the single-line If correctly.

## SEP_002 valid: Empty statements from repeated colons

```vba
Public Sub Demo()
    Debug.Print "a":::Debug.Print "b"
End Sub
```

Expected: verify exact behavior. Parser must not crash on empty statements.

## SEP_003 valid: Label followed by statement on same line

```vba
Public Sub Demo()
StartHere: Debug.Print "start"
End Sub
```

Expected: valid if grammar allows. Verify.

## SEP_004 valid: On Error label with colon

```vba
Public Sub Demo()
    On Error GoTo Handler
    Err.Raise 5
    Exit Sub
Handler:
    Debug.Print Err.Number
End Sub
```

Expected: valid.

## SEP_005 invalid: End If used for single-line If

```vba
Public Sub Demo()
    If True Then Debug.Print "x"
    End If
End Sub
```

Expected: invalid. Single-line If does not open a multiline If block.

## SEP_006 valid: End statement, not End Sub

```vba
Public Sub Demo()
    End
End Sub
```

Expected: valid but style warning. `End` abruptly terminates execution.

## SEP_007 valid: Stop statement

```vba
Public Sub Demo()
    Stop
End Sub
```

Expected: valid but optional debug/style warning.

---

# G. Select Case Edge Cases

## CASE_001 valid: Multiple Case values

```vba
Public Sub Demo(ByVal x As Long)
    Select Case x
        Case 1, 3, 5
            Debug.Print "odd small"
        Case 2, 4, 6
            Debug.Print "even small"
    End Select
End Sub
```

Expected: valid.

## CASE_002 valid: Case Is comparison

```vba
Public Sub Demo(ByVal x As Long)
    Select Case x
        Case Is < 0
            Debug.Print "negative"
        Case Is = 0
            Debug.Print "zero"
        Case Is > 0
            Debug.Print "positive"
    End Select
End Sub
```

Expected: valid.

## CASE_003 valid: Case range and comma mix

```vba
Public Sub Demo(ByVal x As Long)
    Select Case x
        Case 1 To 10, 20 To 30, 99
            Debug.Print "matched"
        Case Else
            Debug.Print "other"
    End Select
End Sub
```

Expected: valid.

## CASE_004 invalid: Case outside Select

```vba
Public Sub Demo()
    Case 1
        Debug.Print "bad"
End Sub
```

Expected: invalid.

---

# H. Loop Edge Cases

## LOOP_001 valid: For Next without variable after Next

```vba
Public Sub Demo()
    Dim i As Long
    For i = 1 To 3
        Debug.Print i
    Next
End Sub
```

Expected: valid.

## LOOP_002 warning or invalid: mismatched Next variable

```vba
Public Sub Demo()
    Dim i As Long, j As Long
    For i = 1 To 3
        For j = 1 To 3
            Debug.Print i, j
        Next i
    Next j
End Sub
```

Expected: invalid or semantic warning depending host behavior. Parser must track loop stack for high-quality diagnostics.

## LOOP_003 valid: Exit For / Exit Do

```vba
Public Sub Demo()
    Dim i As Long
    For i = 1 To 10
        If i = 5 Then Exit For
    Next i

    Do
        Exit Do
    Loop
End Sub
```

Expected: valid.

## LOOP_004 invalid: Exit For outside For

```vba
Public Sub Demo()
    Exit For
End Sub
```

Expected: semantic/compile error.

## LOOP_005 valid: For Each Variant over Array

```vba
Public Sub Demo()
    Dim item As Variant
    For Each item In Array("a", "b", "c")
        Debug.Print item
    Next item
End Sub
```

Expected: valid.

## LOOP_006 invalid: For Each control variable as scalar non-Variant over collection-like values

```vba
Public Sub Demo()
    Dim item As Long
    For Each item In Array(1, 2, 3)
        Debug.Print item
    Next item
End Sub
```

Expected: verify host compile behavior. Often For Each control variable must be Variant or Object. Good semantic test.

---

# I. Arrays, Erase, ReDim, and Bounds

## ARR_005 valid: Erase dynamic array

```vba
Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10)
    Erase values
End Sub
```

Expected: valid.

## ARR_006 valid: Erase fixed array

```vba
Public Sub Demo()
    Dim values(1 To 10) As Long
    Erase values
End Sub
```

Expected: valid.

## ARR_007 invalid: ReDim fixed array

```vba
Public Sub Demo()
    Dim values(1 To 10) As Long
    ReDim values(1 To 20)
End Sub
```

Expected: semantic compile error.

## ARR_008 valid: multidimensional ReDim Preserve changing last dimension

```vba
Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10, 1 To 2)
    ReDim Preserve values(1 To 10, 1 To 3)
End Sub
```

Expected: valid.

## ARR_009 invalid: ReDim Preserve changing non-last dimension

```vba
Public Sub Demo()
    Dim values() As Long
    ReDim values(1 To 10, 1 To 2)
    ReDim Preserve values(1 To 20, 1 To 2)
End Sub
```

Expected: semantic/runtime issue. Classify according to compile behavior.

## ARR_010 valid: ParamArray consumed as Variant array

```vba
Public Sub Demo()
    PrintAll 1, "two", #1/1/2026#
End Sub

Private Sub PrintAll(ParamArray values() As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Debug.Print values(i)
    Next i
End Sub
```

Expected: valid.

---

# J. UDT and Enum Edge Cases

## UDT_001 valid: Fixed-length string member

```vba
Private Type RecordHeader
    Name As String * 20
    Id As Long
End Type
```

Expected: valid.

## UDT_002 valid: Fixed-size array member

```vba
Private Type Matrix3x3
    Values(1 To 3, 1 To 3) As Double
End Type
```

Expected: valid.

## UDT_003 valid: Dynamic array member

```vba
Private Type Packet
    Bytes() As Byte
End Type
```

Expected: valid if MS-VBAL permits dynamic array members in UDT. Verify.

## UDT_004 invalid: Object member in public UDT

```vba
Public Type BadRecord
    Sheet As Worksheet
End Type
```

Expected: verify exact rule. This is a high-value semantic case because public UDTs have restrictions across project boundaries.

## UDT_005 valid: Enum implicit values

```vba
Public Enum ColorCode
    ColorRed
    ColorGreen
    ColorBlue
End Enum
```

Expected: valid.

## UDT_006 valid: Enum explicit negative and hex values

```vba
Public Enum Flags
    FlagNone = 0
    FlagRead = &H1
    FlagWrite = &H2
    FlagAll = -1
End Enum
```

Expected: valid.

## UDT_007 invalid: Duplicate UDT member

```vba
Private Type Person
    Id As Long
    Id As String
End Type
```

Expected: semantic error.

---

# K. Property Procedure Edge Cases

## PROP_004 valid: Indexed property Get / Let

```vba
Private m_Items(1 To 10) As String

Public Property Get Item(ByVal index As Long) As String
    Item = m_Items(index)
End Property

Public Property Let Item(ByVal index As Long, ByVal value As String)
    m_Items(index) = value
End Property
```

Expected: valid.

## PROP_005 invalid: Property Let missing value parameter

```vba
Public Property Let Name()
End Property
```

Expected: invalid.

## PROP_006 valid: Object property Get requires Set return assignment

```vba
Private m_Sheet As Worksheet

Public Property Get Sheet() As Worksheet
    Set Sheet = m_Sheet
End Property
```

Expected: valid.

## PROP_007 invalid: Object property Get without Set return assignment

```vba
Private m_Sheet As Worksheet

Public Property Get Sheet() As Worksheet
    Sheet = m_Sheet
End Property
```

Expected: semantic error.

## PROP_008 valid: Property Set object setter

```vba
Private m_Sheet As Worksheet

Public Property Set Sheet(ByVal value As Worksheet)
    Set m_Sheet = value
End Property
```

Expected: valid.

## PROP_009 invalid: Property Set with value type

```vba
Private m_Name As String

Public Property Set Name(ByVal value As String)
    m_Name = value
End Property
```

Expected: invalid or semantic error. Value properties use Property Let.

---

# L. Class, Interface, Event, and WithEvents Edge Cases

## CLS_006 valid class: Class_Initialize and Class_Terminate

```vba
Option Explicit

Private Sub Class_Initialize()
    Debug.Print "init"
End Sub

Private Sub Class_Terminate()
    Debug.Print "terminate"
End Sub
```

Expected: valid in class module.

## CLS_007 valid class: RaiseEvent with declared Event

```vba
Option Explicit

Public Event Changed(ByVal Name As String)

Public Sub Touch()
    RaiseEvent Changed("Touch")
End Sub
```

Expected: valid in class module.

## CLS_008 invalid: RaiseEvent undeclared

```vba
Option Explicit

Public Sub Touch()
    RaiseEvent Changed("Touch")
End Sub
```

Expected: semantic error.

## CLS_009 invalid standard module: Event declaration in standard module

```vba
Public Event Changed()
```

Expected: module-kind semantic error if Event is class-only. Verify exact host behavior.

## CLS_010 invalid: WithEvents array

```vba
Option Explicit

Private WithEvents Apps(1 To 2) As Application
```

Expected: invalid. `WithEvents` variables cannot be arrays.

## CLS_011 invalid: WithEvents As New

```vba
Option Explicit

Private WithEvents App As New Application
```

Expected: invalid. `WithEvents` and `New` are not compatible.

## CLS_012 valid class: Implements in declaration section

```vba
Option Explicit
Implements IDisposableLike

Private Sub IDisposableLike_Dispose()
End Sub
```

Expected: syntactically valid in class module if interface exists.

## CLS_013 invalid: Implements after procedure

```vba
Option Explicit

Public Sub Demo()
End Sub

Implements IDisposableLike
```

Expected: invalid. `Implements` belongs in the declaration section.

---

# M. Excel Object Model and Host-Aware Edge Cases

## XL_009 valid: ThisWorkbook, ActiveWorkbook, and Workbooks distinction

```vba
Public Sub Demo()
    Debug.Print ThisWorkbook.Name
    Debug.Print ActiveWorkbook.Name
    Debug.Print Workbooks.Count
End Sub
```

Expected: valid. Optional host-aware warning if code uses `ActiveWorkbook` where `ThisWorkbook` is likely intended.

## XL_010 warning: Unqualified Cells

```vba
Public Sub Demo()
    Cells(1, 1).Value = "active sheet dependent"
End Sub
```

Expected: syntactically valid. Host-aware warning.

## XL_011 warning: Unqualified Rows and Columns

```vba
Public Sub Demo()
    Rows(1).Hidden = True
    Columns("A").Hidden = True
End Sub
```

Expected: syntactically valid. Host-aware warning.

## XL_012 valid: Fully qualified Cells through Worksheet

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Data").Cells(1, 1).Value = "safe"
End Sub
```

Expected: valid.

## XL_013 valid: Range object assigned with Set

```vba
Public Sub Demo()
    Dim target As Range
    Set target = ThisWorkbook.Worksheets("Data").Range("A1")
    target.Value = 1
End Sub
```

Expected: valid.

## XL_014 invalid: Range object assignment without Set

```vba
Public Sub Demo()
    Dim target As Range
    target = ThisWorkbook.Worksheets("Data").Range("A1")
End Sub
```

Expected: semantic error.

## XL_015 valid: default worksheet code names

```vba
Public Sub Demo()
    Sheet1.Range("A1").Value = "code name"
End Sub
```

Expected: syntactically valid. Host symbol resolution depends on project. Do not hard-error unknown sheet code names in syntax phase.

## XL_016 valid: Me in worksheet module

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Me.Range("A1").Value = Target.Address
End Sub
```

Expected: valid in worksheet module.

## XL_017 warning: Me in standard module

```vba
Public Sub Demo()
    Me.Range("A1").Value = 1
End Sub
```

Expected: syntax valid shape, but module-kind semantic error/warning in standard module.

## XL_018 valid: Application.Run macro call

```vba
Public Sub Demo()
    Application.Run "OtherMacro", 1, "two"
End Sub
```

Expected: valid.

## XL_019 valid: Formula string is not VBA

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Data").Range("A1").Formula = "=SUM(B1:B10)"
End Sub
```

Expected: valid. Do not parse formula strings as VBA.

## XL_020 warning: Formula string appears malformed but should not be syntax error

```vba
Public Sub Demo()
    Range("A1").Formula = "=SUM(B1:B10"
End Sub
```

Expected: VBA syntax valid. Optional host-aware Excel formula warning only if formula linting is explicitly enabled.

---

# N. Conditional Compilation Edge Cases

## PP_005 valid: complex #If expression

```vba
#Const DEBUG_MODE = True
#Const TARGET_X64 = False

#If DEBUG_MODE And Not TARGET_X64 Then
Public Sub Demo()
    Debug.Print "debug x86-ish"
End Sub
#End If
```

Expected: valid.

## PP_006 mode-sensitive: inactive branch with invalid code

```vba
#Const USE_BAD_CODE = False

#If USE_BAD_CODE Then
Public Sub Broken()
    If True Then
#Else
Public Sub Working()
    Debug.Print "ok"
End Sub
#End If
```

Expected: active-branch compile may be valid, whole-file strict parser may detect invalid inactive branch. The linter must have a configurable policy: `parseAllBranches` vs `activeBranchOnly`.

## PP_007 invalid: #Else without #If

```vba
#Else
Public Sub Demo()
End Sub
```

Expected: invalid.

## PP_008 invalid: duplicate #Else

```vba
#If VBA7 Then
Public Sub Demo()
End Sub
#Else
Public Sub Demo2()
End Sub
#Else
Public Sub Demo3()
End Sub
#End If
```

Expected: invalid.

## PP_009 valid: conditional Declare with PtrSafe and LongPtr

```vba
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
#End If
```

Expected: valid.

---

# O. Declare, PtrSafe, LongPtr, LongLong, and External APIs

## API_004 valid: Declare with Alias

```vba
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
#End If
```

Expected: valid.

## API_005 valid: Declare Sub

```vba
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
```

Expected: valid.

## API_006 warning: Long used where LongPtr likely required

```vba
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
#End If
```

Expected: syntactically valid, host/platform warning.

## API_007 invalid or warning: LongLong outside 64-bit-only context

```vba
Private value As LongLong
```

Expected: mode-sensitive. Verify Excel/VBA7/Win64 behavior.

## API_008 valid: Any in Declare parameter

```vba
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, _
        Source As Any, _
        ByVal Length As LongPtr)
#End If
```

Expected: valid if `Any` is accepted in Declare contexts. Verify.

---

# P. File I/O Statements and `#` Token Edge Cases

## FILE_001 valid: Open For Output

```vba
Public Sub Demo()
    Dim n As Integer
    n = FreeFile
    Open "C:\Temp\x.txt" For Output As #n
    Print #n, "hello"
    Close #n
End Sub
```

Expected: valid syntax. Optional host warning about absolute path.

## FILE_002 valid: Line Input

```vba
Public Sub Demo()
    Dim n As Integer, text As String
    n = FreeFile
    Open "C:\Temp\x.txt" For Input As #n
    Line Input #n, text
    Close #n
End Sub
```

Expected: valid.

## FILE_003 valid: Binary access Put/Get

```vba
Public Sub Demo()
    Dim n As Integer
    Dim b As Byte
    n = FreeFile
    Open "C:\Temp\x.bin" For Binary As #n
    b = 42
    Put #n, 1, b
    Get #n, 1, b
    Close #n
End Sub
```

Expected: valid.

## FILE_004 valid: Write # statement

```vba
Public Sub Demo()
    Dim n As Integer
    n = FreeFile
    Open "C:\Temp\x.csv" For Output As #n
    Write #n, "a", 1, #1/1/2026#
    Close #n
End Sub
```

Expected: valid. Lexer must not confuse `#1/1/2026#` date literal with file-number `#n` token.

## FILE_005 invalid: Missing # in file number

```vba
Public Sub Demo()
    Open "C:\Temp\x.txt" For Output As 1
End Sub
```

Expected: verify exact grammar. Likely invalid.

---

# Q. Error Handling and Legacy Branching Edge Cases

## ERR_005 valid: On Error GoTo 0

```vba
Public Sub Demo()
    On Error Resume Next
    Debug.Print 1 / 0
    On Error GoTo 0
End Sub
```

Expected: valid.

## ERR_006 mode-sensitive: On Error GoTo -1

```vba
Public Sub Demo()
    On Error GoTo -1
End Sub
```

Expected: verify exact VBA host behavior. Do not assume VB.NET rule.

## ERR_007 valid: Resume label

```vba
Public Sub Demo()
    On Error GoTo Handler
    Err.Raise 5
    Exit Sub
Handler:
    Resume CleanExit
CleanExit:
End Sub
```

Expected: valid.

## ERR_008 invalid: Resume outside error handler context

```vba
Public Sub Demo()
    Resume Next
End Sub
```

Expected: syntactically valid shape but semantic/runtime-invalid usage. Classify carefully.

## ERR_009 valid: On expression GoTo

```vba
Public Sub Demo(ByVal choice As Long)
    On choice GoTo One, Two, Three
    Exit Sub
One:
    Debug.Print "one"
    Exit Sub
Two:
    Debug.Print "two"
    Exit Sub
Three:
    Debug.Print "three"
End Sub
```

Expected: valid legacy syntax. Optional style warning.

## ERR_010 valid: On expression GoSub

```vba
Public Sub Demo(ByVal choice As Long)
    On choice GoSub One, Two
    Exit Sub
One:
    Debug.Print "one"
    Return
Two:
    Debug.Print "two"
    Return
End Sub
```

Expected: valid legacy syntax. Optional style warning.

---

# R. Realtime Incremental Recovery Cases

## RT_005 incomplete: Partially typed access modifier

Steps:

```text
1. "Pub"
2. "Public"
3. "Public Sub"
4. "Public Sub Demo"
5. "Public Sub Demo()"
```

Expected:

```text
1. incomplete, no hard unknown-token error while active
2. incomplete declaration
3. incomplete declaration
4. incomplete or valid depending host grammar; verify
5. incomplete procedure body until End Sub
```

## RT_006 incomplete: Named argument typing

Steps:

```text
1. "MsgBox Prompt"
2. "MsgBox Prompt:"
3. "MsgBox Prompt:="
4. "MsgBox Prompt:=\"hello\""
```

Expected:

```text
1. incomplete/warning only while active
2. incomplete
3. incomplete
4. valid statement if in procedure context
```

## RT_007 incomplete: Line continuation typing

Steps:

```text
1. "Debug.Print \"a\" &"
2. "Debug.Print \"a\" & _"
3. "Debug.Print \"a\" & _\n"
4. "Debug.Print \"a\" & _\n    \"b\""
```

Expected:

```text
1. incomplete expression
2. incomplete continuation
3. incomplete continuation target
4. valid
```

## RT_008 incomplete: Block closer recovery

Steps:

```text
1. "If x Then\n    For i = 1 To 10\n"
2. "If x Then\n    For i = 1 To 10\n    Next i"
3. "If x Then\n    For i = 1 To 10\n    Next i\nEnd If"
```

Expected:

```text
1. incomplete, suppress cascading errors
2. incomplete, missing End If only
3. valid if wrapped inside procedure context
```

## RT_009 incomplete: Excel chain typing

Steps:

```text
1. "ThisWorkbook"
2. "ThisWorkbook."
3. "ThisWorkbook.Worksheets("
4. "ThisWorkbook.Worksheets(\"Data\")"
5. "ThisWorkbook.Worksheets(\"Data\").Range("
6. "ThisWorkbook.Worksheets(\"Data\").Range(\"A1\")."
7. "ThisWorkbook.Worksheets(\"Data\").Range(\"A1\").Value = 1"
```

Expected: no hard squiggle on steps 2, 3, 5, or 6 while active. These should drive completion, not punish the user.

---

# S. Syntax Highlighter / Tokenizer Torture Tests

## TOK_001 valid: apostrophe inside string is not comment

```vba
Public Sub Demo()
    Debug.Print "it's not a comment"
End Sub
```

Expected: valid. Syntax highlighter must not start a comment at apostrophe inside string.

## TOK_002 valid: double quote escape inside string

```vba
Public Sub Demo()
    Debug.Print "a ""quoted"" word"
End Sub
```

Expected: valid.

## TOK_003 invalid: C-style comment token

```vba
Public Sub Demo()
    Debug.Print 1 /* bad */
End Sub
```

Expected: invalid.

## TOK_004 valid: Rem after colon

```vba
Public Sub Demo()
    Debug.Print "before": Rem this is a comment
    Debug.Print "after"
End Sub
```

Expected: valid if grammar allows `Rem` as a statement after colon. Verify.

## TOK_005 trap: Rem inside identifier should not be comment

```vba
Public Sub Demo()
    Dim RememberMe As Long
    RememberMe = 1
End Sub
```

Expected: valid. Lexer must not treat `Rem` prefix inside identifier as comment.

## TOK_006 trap: date literal vs preprocessor hash

```vba
#Const DEBUG_MODE = True

Public Sub Demo()
    Debug.Print #5/30/2026#
End Sub
```

Expected: valid. Lexer must distinguish preprocessor directives from date literals by context.

---

# T. Suggested Additional Diagnostic Codes

Add these if the extension supports granular diagnostics:

```text
VBA_PARSE_INVALID_OPTION_POSITION
VBA_PARSE_INVALID_CALL_PARENS
VBA_PARSE_EXPECTED_PROPERTY_VALUE_PARAMETER
VBA_PARSE_CASE_WITHOUT_SELECT
VBA_PARSE_ELSE_WITHOUT_IF
VBA_PARSE_DUPLICATE_CONDITIONAL_ELSE
VBA_PARSE_PREPROCESSOR_MISSING_END_IF
VBA_SEM_EXIT_WRONG_PROCEDURE_KIND
VBA_SEM_SET_REQUIRED_FOR_OBJECT
VBA_SEM_SET_INVALID_FOR_VALUE
VBA_SEM_IS_REQUIRES_OBJECT
VBA_SEM_REDECLARE_TYPE_SUFFIX_AND_AS_CLAUSE
VBA_SEM_REDIM_FIXED_ARRAY
VBA_SEM_REDIM_PRESERVE_ONLY_LAST_DIMENSION
VBA_SEM_WITH_EVENTS_INVALID_CONTEXT
VBA_SEM_RAISEEVENT_UNDECLARED
VBA_SEM_IMPLEMENT_POSITION
VBA_HOST_ME_INVALID_IN_STANDARD_MODULE
VBA_HOST_EXCEL_UNQUALIFIED_CELLS
VBA_HOST_EXCEL_FORMULA_STRING_NOT_VBA
VBA_STYLE_LEGACY_DEF_TYPE
VBA_STYLE_IMPLICIT_DEFAULT_MEMBER
VBA_STYLE_ACTIVEWORKBOOK_AMBIGUITY
```

---

# U. Highest-Value New Cases to Prioritize

Prioritize these before the rest:

```text
1. Attribute lines in exported modules, including member attributes.
2. Call syntax with and without parentheses.
3. Single-line If with colon-separated statements.
4. Conditional compilation inactive branches.
5. Property Get/Let/Set object-vs-value rules.
6. WithEvents restrictions in class modules.
7. File I/O statements using # tokens.
8. Type-declaration suffixes on identifiers and functions.
9. Excel formula strings that must not be parsed as VBA.
10. Realtime recovery for named arguments and member chains.
```

---

# Final Agent Note

The extension should not only know what is wrong. It should know **when not to complain yet**. In VS Code, a linter that is technically right too early feels broken. VBA needs a tolerant foreground parser and a stricter background validator.
