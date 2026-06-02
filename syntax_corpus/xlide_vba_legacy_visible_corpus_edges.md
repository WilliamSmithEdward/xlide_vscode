# XLIDE VBA Legacy Visible Corpus Edges

## Purpose

This document defines one final recommended edge-case corpus file for XLIDE's realtime VBA analysis suite.

The goal is to cover **weird but VBE-visible VBA**: syntax and language forms that real users can see or type inside the VBA IDE, but that are easy for a modern parser to miss.

Recommended file name:

```text
18_legacy_visible_vba_edges.md
```

Optional follow-up file:

```text
19_negative_recovery_fuzz_seed.md
```

These files are not meant to replace the core corpus. They are hardening layers.

---

# 18. `legacy_visible_vba_edges.md`

## Coverage Goal

Test old, unusual, or easily misparsed VBA constructs that are visible in the VBA editor.

This file should focus on:

```text
line numbers
numeric error targets
DefType statements
Let / Set / LSet / RSet
Mid assignment statement
file I/O syntax
Optional / ByRef / ParamArray rules
Option Compare
AddressOf
date literal and # ambiguity
colon-separated statements
labels mixed with statements
```

The purpose is not to make the analyzer noisy. The purpose is to prevent valid legacy VBA from being falsely rejected.

---

## 1. Line Numbers and Numeric Error Targets

Old VBA allows numeric line labels.

### Valid: numeric line labels with `On Error GoTo`

```vba
Public Sub Test()
10  On Error GoTo 90
20  Debug.Print 1 / 0
30  Exit Sub
90  Debug.Print Err.Number
End Sub
```

Expected:

```text
valid VBA
numeric line labels recognized
On Error GoTo numeric target accepted
no parser confusion from leading numbers
```

### Valid: named label

```vba
Public Sub Test()
StartHere:
    On Error GoTo CleanFail
    Exit Sub

CleanFail:
    Debug.Print Err.Description
End Sub
```

Expected:

```text
valid VBA
named labels recognized
label target resolved if project/symbol pass exists
```

### Invalid or warning: missing label target

```vba
Public Sub Test()
    On Error GoTo MissingLabel
End Sub
```

Expected:

```text
syntax: valid
semantic/project diagnostic: unresolved label target
not parser error
```

---

## 2. `DefType` Statements

`DefInt`, `DefLng`, `DefStr`, and related statements are module-level declarations that affect implicit typing by first letter.

### Valid: module-level `DefLng`

```vba
DefLng A-Z

Public Function CountRows()
    CountRows = 42
End Function
```

Expected:

```text
valid VBA
DefLng accepted only at module declaration level
implicit return type may be Long depending on symbol/type pass
```

### Valid: multiple ranges

```vba
DefStr A-C
DefLng D-Z

Public Sub Test()
    Dim apple
    Dim count
End Sub
```

Expected:

```text
valid VBA
apple default type affected by DefStr
count default type affected by DefStr or DefLng depending on first letter
```

### Invalid: `DefLng` inside procedure

```vba
Public Sub Test()
    DefLng A-Z
End Sub
```

Expected:

```text
diagnostic: DefType statement invalid inside procedure
not a lexer failure
```

---

## 3. `Let`, Omitted `Let`, `Set`, `LSet`, and `RSet`

VBA assignment has several forms.

### Valid: explicit and omitted `Let`

```vba
Public Sub Test()
    Dim x As Long
    Let x = 1
    x = 2
End Sub
```

Expected:

```text
valid VBA
explicit Let accepted
omitted Let accepted
```

### Valid: fixed-length string with `LSet` and `RSet`

```vba
Public Sub Test()
    Dim s As String * 10
    LSet s = "abc"
    RSet s = "xyz"
End Sub
```

Expected:

```text
valid VBA
LSet and RSet recognized as statements
fixed-length string syntax accepted
```

### Invalid: `Let` with object assignment

```vba
Public Sub Test()
    Dim ws As Worksheet
    Let ws = ThisWorkbook.Worksheets(1)
End Sub
```

Expected:

```text
semantic diagnostic: object assignment requires Set
not syntax error
```

### Invalid: `Set` with scalar

```vba
Public Sub Test()
    Dim x As Long
    Set x = 123
End Sub
```

Expected:

```text
semantic diagnostic: Set requires object reference
not parser crash
```

---

## 4. `Mid` Assignment Statement

`Mid` and `Mid$` can appear on the left side of assignment.

This is easy to misparse as a function call used as an l-value.

### Valid: `Mid$` statement

```vba
Public Sub Test()
    Dim s As String
    s = "abcdef"
    Mid$(s, 2, 3) = "XYZ"
End Sub
```

Expected:

```text
valid VBA
Mid$ assignment statement recognized
not treated as invalid function-call assignment
```

### Valid: `Mid` statement without length

```vba
Public Sub Test()
    Dim s As String
    s = "abcdef"
    Mid(s, 2) = "zz"
End Sub
```

Expected:

```text
valid VBA
optional length omitted
```

### Invalid: non-variable target

```vba
Public Sub Test()
    Mid$("abcdef", 2, 3) = "XYZ"
End Sub
```

Expected:

```text
semantic diagnostic: target must be mutable string variable
not raw parse failure
```

---

## 5. File I/O Grammar

VBA file I/O uses syntax that does not look like normal procedure calls.

### Valid: output file

```vba
Public Sub Test()
    Open "C:\temp\out.txt" For Output As #1
    Print #1, "hello"
    Close #1
End Sub
```

Expected:

```text
valid VBA
Open statement recognized
file number syntax recognized
Print # recognized
Close # recognized
```

### Valid: input file with `Line Input #`

```vba
Public Sub Test()
    Dim s As String
    Open "C:\temp\in.txt" For Input As #1
    Line Input #1, s
    Close #1
End Sub
```

Expected:

```text
valid VBA
Line Input # recognized as file input statement
not confused with InputBox or normal function call
```

### Valid: free file number

```vba
Public Sub Test()
    Dim f As Integer
    f = FreeFile
    Open "C:\temp\out.txt" For Append As #f
    Print #f, "hello"
    Close #f
End Sub
```

Expected:

```text
valid VBA
file number may be variable after #
```

### Invalid: malformed Open

```vba
Public Sub Test()
    Open "C:\temp\out.txt" Output #1
End Sub
```

Expected:

```text
parse diagnostic: expected For / As structure
range on malformed Open statement
```

---

## 6. Optional, ByRef Default, and ParamArray

VBA defaults procedure arguments to `ByRef`.

Optional arguments and `ParamArray` have ordering restrictions.

### Valid: implicit ByRef

```vba
Public Sub TestArg(x As Long)
    x = x + 1
End Sub
```

Expected:

```text
valid VBA
x is ByRef by default
optional hint possible, not required
```

### Valid: optional trailing argument

```vba
Public Sub TestOptional(ByVal a As Long, Optional ByVal b As Long = 10)
End Sub
```

Expected:

```text
valid VBA
Optional argument has default value
```

### Invalid: required argument after optional

```vba
Public Sub Bad(Optional ByVal a As Long, ByVal b As Long)
End Sub
```

Expected:

```text
declaration diagnostic: required argument cannot follow Optional argument
```

### Invalid: Optional with ParamArray

```vba
Public Sub Bad(Optional ByVal a As Long = 1, ParamArray rest())
End Sub
```

Expected:

```text
declaration diagnostic: Optional cannot be combined with ParamArray in this shape
```

### Valid: ParamArray last

```vba
Public Sub LogMany(ParamArray values())
End Sub
```

Expected:

```text
valid VBA
ParamArray recognized
must be last argument
```

### Invalid: ParamArray not last

```vba
Public Sub Bad(ParamArray values(), ByVal x As Long)
End Sub
```

Expected:

```text
declaration diagnostic: ParamArray must be last
```

---

## 7. `Option Compare`

`Option Compare` is module-level and must appear before procedures.

### Valid: text compare

```vba
Option Compare Text
Option Explicit

Public Sub Test()
    Debug.Print "a" = "A"
End Sub
```

Expected:

```text
valid VBA
Option Compare accepted in module header
```

### Valid: binary compare

```vba
Option Compare Binary
Option Explicit

Public Sub Test()
    Debug.Print "a" = "A"
End Sub
```

Expected:

```text
valid VBA
```

### Host-sensitive: database compare

```vba
Option Compare Database

Public Sub Test()
End Sub
```

Expected:

```text
syntax may be valid in VBA family
Excel host may warn or reject depending on target policy
classify as host-sensitive, not generic lexer error
```

### Invalid placement

```vba
Public Sub Test()
End Sub

Option Compare Text
```

Expected:

```text
diagnostic: Option Compare must appear before procedures
not parser crash
```

---

## 8. `AddressOf`

`AddressOf` is used for passing procedure addresses, commonly to Windows API callbacks.

### Valid: callback pattern

```vba
Private Declare PtrSafe Function EnumWindows Lib "user32" ( _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr) As Long

Public Sub Test()
    EnumWindows AddressOf EnumProc, 0
End Sub

Private Function EnumProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
    EnumProc = 1
End Function
```

Expected:

```text
valid VBA in VBA7/PtrSafe context
AddressOf recognized as operator
LongPtr recognized
line continuations recognized
```

### Invalid: AddressOf non-procedure target

```vba
Public Sub Test()
    Dim x As Long
    Debug.Print AddressOf x
End Sub
```

Expected:

```text
semantic diagnostic: AddressOf requires procedure target
not lexer error
```

---

## 9. Date Literals and `#` Ambiguity

The `#` token appears in multiple unrelated VBA contexts.

It can mean:

```text
date literal delimiter
file number marker
Double type-declaration suffix
compiler/file syntax context
```

### Valid: date literal

```vba
Public Sub Test()
    Dim d As Date
    d = #1/2/2026#
End Sub
```

Expected:

```text
valid VBA
date literal tokenized as one literal
```

### Valid: file number and date literal together

```vba
Public Sub Test()
    Open "C:\temp\x.txt" For Output As #1
    Print #1, #1/2/2026#
    Close #1
End Sub
```

Expected:

```text
valid VBA
first # after As/Print is file-number context
second #...# is date literal
no lexer confusion
```

### Valid: Double type suffix

```vba
Public Sub Test()
    Dim ratio#
    ratio# = 1.5
End Sub
```

Expected:

```text
valid VBA
# after identifier is Double type suffix
not date literal
```

---

## 10. Colon Statement Separators with Labels and Comments

Colon-separated statements are common in old VBA and are easy to misparse.

### Valid: multiple statements on one line

```vba
Public Sub Test()
    a = 1: b = 2: Debug.Print a + b
End Sub
```

Expected:

```text
valid VBA
colon separates statements
range mapping still accurate
```

### Valid: labels and statements on same line

```vba
Public Sub Test()
StartHere: Debug.Print "start": GoTo Done
Done: Debug.Print "done"
End Sub
```

Expected:

```text
valid VBA
label recognized before colon
statements after label recognized
```

### Valid: comment after colon

```vba
Public Sub Test()
    a = 1: ' comment after colon
    b = 2
End Sub
```

Expected:

```text
valid VBA
comment consumes rest of physical line
b = 2 remains next line
```

### Bad recovery case

```vba
Public Sub Test()
    a = : b = 2: Debug.Print b
End Sub
```

Expected:

```text
diagnostic on incomplete assignment after a =
parser should recover at next colon
b = 2 should not be incorrectly marked invalid
```

---

# 19. `negative_recovery_fuzz_seed.md`

## Purpose

This optional file seeds parser recovery tests with broken code.

Do not expect perfect diagnostics for every case.

Expected assertions should focus on editor stability.

## Required Assertions

```text
does not crash
does not hang
does not flood diagnostics
does not consume the rest of the module unnecessarily
recovers by next statement where possible
recovers by next procedure when statement recovery fails
```

## Seed Snippets

### Broken If

```vba
Public Sub Test()
    If Then
    Debug.Print "after"
End Sub
```

Expected:

```text
diagnostic on If
Debug.Print should be parsed if recovery succeeds
```

### Broken For

```vba
Public Sub Test()
    For i =
    Debug.Print i
Next i
End Sub
```

Expected:

```text
diagnostic on For initializer
no crash
```

### Broken assignment

```vba
Public Sub Test()
    x =
    y = 2
End Sub
```

Expected:

```text
diagnostic on x =
y = 2 should parse cleanly
```

### Broken call

```vba
Public Sub Test()
    MsgBox (
    Debug.Print "after"
End Sub
```

Expected:

```text
diagnostic on incomplete call
parser should recover by next statement or next procedure
```

### Broken string

```vba
Public Sub Test()
    Debug.Print "unterminated
    Debug.Print "after"
End Sub
```

Expected:

```text
unterminated string diagnostic
do not mark entire rest of module as string
```

### Broken block nesting

```vba
Public Sub Test()
    If x Then
        For i = 1 To 10
    End If
Next i
End Sub
```

Expected:

```text
block mismatch diagnostics
no infinite recovery loop
```

---

# Diagnostic Classification Reminder

Every case should eventually declare its diagnostic lane:

```text
syntax error
compile-shape error
semantic/project error
Excel host warning
style warning
realtime recovery state
valid, no diagnostic
```

This prevents XLIDE from becoming noisy.

Examples:

```text
Range("A1").Value = 1
```

Should be:

```text
valid Excel VBA
optional host warning
not syntax error
```

```text
Debug.Print "unterminated
```

Should be:

```text
syntax/realtime recovery diagnostic
```

```text
Private Sub Worksheet_Change(ByVal Target As Range)
```

Should be:

```text
valid syntax
module-context-sensitive diagnostic if in wrong module kind
```

---

# Stop Rule

After adding these two files, stop expanding the hand-written corpus unless real bugs reveal gaps.

The next quality jump should come from:

```text
executable fixture format
range markers
stable diagnostic codes
realtime vs strict expected behavior
VS Code diagnostic integration tests
optional Excel COM / VBE oracle testing
negative recovery fuzzing
```

At this point, the bottleneck is no longer imagination. It is deterministic execution and comparison.
