# XLIDE VBA Realtime Analysis Final Corpus Addendum

**Recommended file name:** `xlide_vba_realtime_analysis_final_corpus_addendum.md`

**Purpose:** Add a final, narrow hardening layer to the existing XLIDE Excel VBA realtime analysis corpus.

This addendum does **not** attempt to add another broad syntax suite. The existing corpus already covers the main parser, declaration, limits, realtime recovery, module-kind, and legacy-visible cases.

This file focuses on remaining high-value gaps:

1. Excel host syntax traps.
2. Legacy control-transfer forms.
3. Completion-context fixtures.
4. VBE COM canary verdict metadata.
5. UserForm designer-backed implicit symbols.
6. Keyword and symbol casing behavior.

**Non-negotiable rule:** Verify grammar-sensitive behavior against Microsoft **MS-VBAL**. Verify host/compiler behavior against the live VBE/Excel compiler where practical. Do not infer VBA behavior from VB.NET, VBScript, TypeScript, Python, or generic BASIC.

MS-VBAL landing page: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

---

## Suggested Placement

```text
corpus/
  vbe-visible/
    20_excel_host_syntax_traps.md
    21_legacy_control_transfer.md
    22_completion_contexts.md
    23_vbe_canary_verdicts.md
    24_userform_designer_symbols.md
    25_keyword_and_symbol_casing.md
```

These files should be treated as final hardening layers after the core parser and realtime recovery suites are passing.

---

# 20. `excel_host_syntax_traps.md`

## Goal

Cover Excel/VBA-visible syntax that is easy to misparse because it does not look like ordinary VBA procedure-call or member-access syntax.

These should not become broad Excel object-model semantic tests. They are parser and classifier traps.

---

## EXCEL_SYNTAX_001 valid: Excel bracket evaluate shorthand

```vba
Public Sub Demo()
    [A1].Value = 42
    Debug.Print [SUM(1,2,3)]
End Sub
```

Expected:

```text
valid Excel VBA
bracketed Excel evaluate shorthand accepted
do not confuse with bracketed identifier syntax only
```

Notes:

```text
[A1] in Excel VBA is not merely the same category as [Sub] or [Date].
The lexer/parser should support both bracketed identifiers and Excel evaluate shorthand.
Semantic resolution may be host-specific.
```

---

## EXCEL_SYNTAX_002 valid: bracketed keyword identifier

```vba
Public Sub Demo()
    Dim [Date] As Date
    [Date] = Now
    Debug.Print [Date]
End Sub
```

Expected:

```text
valid VBA
bracketed identifier preserved
not classified as Excel range shorthand
```

---

## EXCEL_SYNTAX_003 valid: bang operator with simple member

```vba
Public Sub Demo(ByVal rs As Object)
    Debug.Print rs!CustomerID
End Sub
```

Expected:

```text
valid syntax
semantics are object-dependent
do not reject ! as invalid punctuation
```

---

## EXCEL_SYNTAX_004 valid: bang operator with bracketed member

```vba
Public Sub Demo(ByVal rs As Object)
    Debug.Print rs![Customer ID]
End Sub
```

Expected:

```text
valid syntax
bracketed field/member after ! accepted
semantic resolution deferred
```

---

## EXCEL_SYNTAX_005 valid: TypeOf expression

```vba
Public Sub Demo(ByVal obj As Object)
    If TypeOf obj Is Worksheet Then
        Debug.Print "worksheet"
    End If
End Sub
```

Expected:

```text
valid VBA
TypeOf ... Is ... recognized as a distinct expression form
not treated as normal binary Is between two values
```

---

## EXCEL_SYNTAX_006 invalid: TypeOf without object expression

```vba
Public Sub Demo()
    If TypeOf Is Worksheet Then
        Debug.Print "bad"
    End If
End Sub
```

Expected:

```text
parse or compile-shape diagnostic
diagnostic should localize around malformed TypeOf expression
no cascade across entire If block
```

---

## EXCEL_SYNTAX_007 warning: unqualified Range in standard module

```vba
Public Sub Demo()
    Range("A1").Value = 42
End Sub
```

Expected:

```text
valid VBA syntax
host-aware warning optional
unqualified Range implicitly binds through active sheet/application context
not syntax error
```

---

## EXCEL_SYNTAX_008 valid: qualified Range in worksheet module using Me

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Range("A1").Value = Target.Address
End Sub
```

Expected:

```text
valid in worksheet module
Me resolves to worksheet module instance
```

---

# 21. `legacy_control_transfer.md`

## Goal

Cover old but VBE-visible control-transfer forms that a modern parser may miss.

These are not style recommendations. The analyzer may warn, but it should not falsely reject valid legacy VBA.

---

## LEGACY_TRANSFER_001 valid: GoSub and Return

```vba
Public Sub Demo()
    GoSub Work
    Exit Sub

Work:
    Debug.Print "work"
    Return
End Sub
```

Expected:

```text
valid VBA if confirmed by VBE canary
GoSub target label recognized
Return statement recognized
optional style warning only
```

Canary requirement:

```text
Run through Excel VBE compile canary before hardcoding final verdict.
```

---

## LEGACY_TRANSFER_002 invalid or semantic: Return without GoSub context

```vba
Public Sub Demo()
    Return
End Sub
```

Expected:

```text
verify with VBE canary
if compile-invalid, classify as compile-error
not lexer failure
```

---

## LEGACY_TRANSFER_003 valid: On expression GoTo numeric labels

```vba
Public Sub Demo(ByVal n As Long)
    On n GoTo 100, 200
    Exit Sub

100
    Debug.Print "first"
    Exit Sub

200
    Debug.Print "second"
End Sub
```

Expected:

```text
valid VBA if confirmed
numeric line labels recognized
On expression GoTo list recognized
```

---

## LEGACY_TRANSFER_004 valid: On expression GoSub labels

```vba
Public Sub Demo(ByVal n As Long)
    On n GoSub First, Second
    Exit Sub

First:
    Debug.Print "first"
    Return

Second:
    Debug.Print "second"
    Return
End Sub
```

Expected:

```text
valid VBA if confirmed
On expression GoSub recognized
label list recognized
Return recognized
```

---

## LEGACY_TRANSFER_005 valid: Resume Next in error handler

```vba
Public Sub Demo()
    On Error GoTo Handler
    Err.Raise 5
    Exit Sub

Handler:
    Resume Next
End Sub
```

Expected:

```text
valid VBA
Resume Next recognized
```

---

## LEGACY_TRANSFER_006 valid: Resume label

```vba
Public Sub Demo()
    On Error GoTo Handler
    Err.Raise 5

ContinueHere:
    Debug.Print "continued"
    Exit Sub

Handler:
    Resume ContinueHere
End Sub
```

Expected:

```text
valid VBA
Resume target label resolved during semantic/project pass
missing target should be semantic diagnostic, not syntax failure
```

---

## LEGACY_TRANSFER_007 valid: On Error GoTo 0

```vba
Public Sub Demo()
    On Error Resume Next
    Debug.Print 1 / 0
    On Error GoTo 0
End Sub
```

Expected:

```text
valid VBA
On Error GoTo 0 recognized as disabling active error handler
```

---

## LEGACY_TRANSFER_008 valid: On Error GoTo -1

```vba
Public Sub Demo()
    On Error Resume Next
    Err.Raise 5
    On Error GoTo -1
End Sub
```

Expected:

```text
verify host behavior
commonly used to clear current exception state
if accepted by Excel VBE, classify valid
```

---

# 22. `completion_contexts.md`

## Goal

Test parser state used by IntelliSense and member completion.

This is not only a analysis concern. XLIDE's realtime parser should preserve enough context to answer:

```text
What receiver is before the dot?
What module is this code in?
Are we inside a With block?
Are we inside an argument list?
Are we inside a named-argument position?
```

Use marker syntax:

```text
<|>
```

The marker should be removed before parsing and used as the completion trigger position.

---

## Fixture Shape

```jsonc
{
  "id": "COMP_001",
  "title": "Worksheet variable member completion",
  "moduleKind": "standard",
  "trigger": ".",
  "expected": "completion-context",
  "expectedReceiverText": "ws",
  "expectedReceiverType": "Worksheet",
  "diagnostics": [],
  "snippet": "..."
}
```

---

## COMP_001 worksheet variable member completion

```vba
Public Sub Demo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.<|>
End Sub
```

Expected:

```text
completion context produced
receiver text: ws
receiver type: Worksheet, if Excel symbols are loaded
no syntax error from trailing dot in realtime mode
```

---

## COMP_002 ThisWorkbook member completion

```vba
Public Sub Demo()
    ThisWorkbook.<|>
End Sub
```

Expected:

```text
completion context produced
receiver text: ThisWorkbook
receiver type: Workbook, if Excel symbols are loaded
trailing dot is incomplete in realtime mode, not fatal parse failure
```

---

## COMP_003 Application.WorksheetFunction completion

```vba
Public Sub Demo()
    Application.WorksheetFunction.<|>
End Sub
```

Expected:

```text
completion context produced
receiver chain preserved
receiver type: WorksheetFunction, if Excel symbols are loaded
```

---

## COMP_004 Range call result completion

```vba
Public Sub Demo()
    Range("A1").<|>
End Sub
```

Expected:

```text
completion context produced
receiver expression: Range("A1")
receiver type: Range, if Excel symbols are loaded
host warning optional for unqualified Range
```

---

## COMP_005 Me completion in worksheet module

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Me.<|>
End Sub
```

Expected:

```text
moduleKind: worksheet
receiver text: Me
receiver type: Worksheet/module instance
valid completion context
```

---

## COMP_006 Me completion in standard module

```vba
Public Sub Demo()
    Me.<|>
End Sub
```

Expected:

```text
moduleKind: standard
syntax may parse
semantic diagnostic: Me invalid or unavailable in standard module
completion should not crash
```

---

## COMP_007 leading dot inside With

```vba
Public Sub Demo()
    With ThisWorkbook.Worksheets(1)
        .<|>
    End With
End Sub
```

Expected:

```text
completion context produced
receiver comes from active With expression
receiver type: Worksheet, if Excel symbols are loaded
```

---

## COMP_008 nested With leading dot

```vba
Public Sub Demo()
    With ThisWorkbook
        With .Worksheets(1)
            .<|>
        End With
    End With
End Sub
```

Expected:

```text
completion context produced
receiver comes from innermost With
outer With still tracked for resolving .Worksheets(1)
```

---

## COMP_009 named argument completion

```vba
Public Sub Demo()
    Workbooks.Open <|>
End Sub
```

Expected:

```text
completion context: argument list / call statement
target procedure/member: Workbooks.Open, if Excel symbols are loaded
suggest named arguments if symbol metadata is available
```

---

## COMP_010 incomplete call argument

```vba
Public Sub Demo()
    MsgBox Prompt:=<|>
End Sub
```

Expected:

```text
completion context: expression after named argument assignment
realtime diagnostic: incomplete expression only
no cascade into End Sub
```

---

# 23. `vbe_canary_verdicts.md`

## Goal

Attach live Excel VBE compile verdicts to ambiguous grammar/behavior cases.

This is a metadata pattern, not a syntax section. The goal is to prevent uncertain corpus cases from remaining permanently vague.

---

## Recommended Canary Metadata

```jsonc
{
  "id": "CANARY_001",
  "title": "Procedure declaration missing parentheses",
  "moduleKind": "standard",
  "sourceMode": "live-vbe",
  "expectedBySpec": "needs-verification",
  "vbeCanary": {
    "host": "Excel",
    "hostVersion": "recorded-by-test-runner",
    "bitness": "recorded-by-test-runner",
    "moduleKind": "standard",
    "compileResult": "accept | reject",
    "errorMessageContains": "optional stable substring",
    "notes": "manual or automated observation"
  },
  "snippet": "Public Sub Demo\nEnd Sub"
}
```

---

## CANARY_001 procedure declaration missing parentheses

```vba
Public Sub Demo
End Sub
```

Expected:

```text
needs VBE canary verdict
do not guess
record actual live compiler result
```

---

## CANARY_002 duplicate Option Explicit

```vba
Option Explicit
Option Explicit

Public Sub Demo()
End Sub
```

Expected:

```text
needs VBE canary verdict
record whether accepted, rejected, or ignored
```

---

## CANARY_003 suffix plus As clause on function

```vba
Public Function GetName$() As String
    GetName = "XLIDE"
End Function
```

Expected:

```text
needs VBE canary verdict
likely duplicate type declaration, but verify
```

---

## CANARY_004 suffix plus As clause on variable

```vba
Public Sub Demo()
    Dim name$ As String
End Sub
```

Expected:

```text
needs VBE canary verdict
likely duplicate type declaration, but verify
```

---

## CANARY_005 descending DefType range

```vba
DefLng Z-A

Public Sub Demo()
End Sub
```

Expected:

```text
needs VBE canary verdict
record exact result
```

---

## CANARY_006 user-typed Attribute in live editor body

```vba
Public Sub Demo()
    Attribute Demo.VB_Description = "Bad here"
End Sub
```

Expected:

```text
sourceMode: live-vbe
likely invalid
verify because exported metadata mode differs from live editor mode
```

---

## CANARY_007 Attribute in exported member metadata position

```vba
Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
Attribute Demo.VB_Description = "Demo macro"
    Debug.Print "ok"
End Sub
```

Expected:

```text
sourceMode: exported-file
verify import/export behavior
not necessarily same as live editor-visible source
```

---

## Canary Policy

```text
If MS-VBAL and VBE canary disagree:
1. Record both.
2. Prefer VBE canary for realtime Excel-host analysis UX.
3. Keep MS-VBAL note for spec tracking.
4. Do not silently weaken the parser without a recorded reason.
```

---

# 24. `userform_designer_symbols.md`

## Goal

Test symbols that are not declared in the code pane but are valid because they are designer-backed UserForm controls or host-generated module members.

These should not be syntax errors. They are symbol-resolution cases.

---

## Fixture Metadata

```jsonc
{
  "moduleKind": "userform",
  "moduleName": "UserForm1",
  "designerSymbols": [
    {
      "name": "CommandButton1",
      "type": "MSForms.CommandButton"
    }
  ]
}
```

---

## FORM_SYMBOL_001 valid: known UserForm control symbol

```vba
Private Sub CommandButton1_Click()
    CommandButton1.Caption = "OK"
End Sub
```

Expected:

```text
valid if designerSymbols contains CommandButton1
CommandButton1 resolved as implicit member/control
no undeclared variable diagnostic
```

---

## FORM_SYMBOL_002 warning: unknown control symbol

```vba
Private Sub CommandButton2_Click()
    CommandButton2.Caption = "OK"
End Sub
```

Expected:

```text
if designer metadata unavailable: avoid hard error or downgrade
if designer metadata available and symbol absent: unresolved control/member diagnostic
not syntax error
```

---

## FORM_SYMBOL_003 valid: Me-qualified known control

```vba
Private Sub CommandButton1_Click()
    Me.CommandButton1.Caption = "OK"
End Sub
```

Expected:

```text
valid with designer symbol metadata
Me resolves to UserForm instance
CommandButton1 resolved as member
```

---

## FORM_SYMBOL_004 valid: UserForm Initialize event

```vba
Private Sub UserForm_Initialize()
    Me.Caption = "Ready"
End Sub
```

Expected:

```text
valid in userform module
host/module-context warning if found in standard module
```

---

## FORM_SYMBOL_005 ambiguous: code-only parse without designer metadata

```vba
Private Sub TextBox1_Change()
    TextBox1.Text = UCase$(TextBox1.Text)
End Sub
```

Expected:

```text
syntax valid
if designer metadata absent, do not mark TextBox1 as syntax error
optional unresolved-symbol warning should be suppressible
```

---

# 25. `keyword_and_symbol_casing.md`

## Goal

Test capitalization normalization without corrupting user-defined symbols, string literals, comments, bracketed identifiers, or host/member names.

This matters for XLIDE because VBA users expect canonical keyword casing, but casing should not be blindly applied everywhere.

---

## CASING_001 canonical keyword casing

Input:

```vba
public sub demo()
    dim customerID as long
    customerID = 1
    debug.print customerID
end sub
```

Expected normalized output:

```vba
Public Sub demo()
    Dim customerID As Long
    customerID = 1
    Debug.Print customerID
End Sub
```

Expected:

```text
keywords normalized
user-defined symbol casing preserved according to declaration/symbol policy
member access casing may require symbol metadata
```

---

## CASING_002 symbol declaration casing propagates to references

Input:

```vba
Public Sub Demo()
    Dim CustomerID As Long
    customerid = 1
    Debug.Print CUSTOMERID
End Sub
```

Expected normalized output:

```vba
Public Sub Demo()
    Dim CustomerID As Long
    CustomerID = 1
    Debug.Print CustomerID
End Sub
```

Expected:

```text
if symbol table available, references adopt declaration casing
if no symbol table, only keyword casing should be changed
```

---

## CASING_003 do not change string literals or comments

Input:

```vba
public sub demo()
    debug.print "public sub should not change inside string"
    ' private function should not change inside comment
end sub
```

Expected normalized output:

```vba
Public Sub demo()
    Debug.Print "public sub should not change inside string"
    ' private function should not change inside comment
End Sub
```

Expected:

```text
string contents preserved exactly
comments preserved exactly unless explicit formatter option says otherwise
```

---

## CASING_004 bracketed identifiers preserved

Input:

```vba
Public Sub Demo()
    Dim [sub] As String
    [sub] = "ok"
End Sub
```

Expected:

```text
bracketed identifier text preserved or normalized only by explicit symbol policy
never convert [sub] into [Sub] merely because Sub is a keyword
```

---

## CASING_005 type-declaration suffix preserved

Input:

```vba
public sub demo()
    dim total&, name$, price@
    total = 1
    name = "x"
    price = 12.34@
end sub
```

Expected normalized output:

```vba
Public Sub demo()
    Dim total&, name$, price@
    total = 1
    name = "x"
    price = 12.34@
End Sub
```

Expected:

```text
keyword casing corrected
legacy type-declaration suffixes preserved
numeric literal suffix preserved
```

---

## CASING_006 host member casing with symbol metadata

Input:

```vba
Public Sub Demo()
    thisworkbook.worksheets(1).range("A1").value = 42
End Sub
```

Expected normalized output if Excel symbol metadata is loaded:

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets(1).Range("A1").Value = 42
End Sub
```

Expected normalized output if Excel symbol metadata is not loaded:

```text
implementation-defined
should at minimum not corrupt tokens
may leave host member casing unchanged
```

---

# Integration Guidance

## Keep These Separate From Core Parser Tests

These are late-stage hardening tests. Do not use them to block basic parser development before the main corpus passes.

Recommended validation order:

```text
1. lexical fixtures
2. parser fixtures
3. realtime incomplete fixtures
4. declaration/module-kind fixtures
5. host-aware semantic fixtures
6. legacy visible fixtures
7. completion-context fixtures
8. casing/normalization fixtures
9. VBE canary verdict reconciliation
```

---

## Diagnostic Classification Reminder

Use these classes consistently:

```text
syntax-error       cannot parse as valid VBA
compile-error      parseable, but VBE compiler rejects
semantic-warning   suspicious or symbol-dependent
host-warning       Excel-specific risk, not pure VBA grammar
style-warning      valid but discouraged
incomplete         valid partial state during active typing
completion-context parser state for IntelliSense, not a diagnostic
needs-verification must be resolved by MS-VBAL and/or VBE canary
```

---

## What Not To Add

Avoid expanding this addendum into:

```text
more basic If/For/Select syntax
more ordinary Dim/Sub/Function cases
large random invalid-code fuzzing
hidden metadata diagnostics shown to users
VB.NET-inspired conveniences
formatting rules not tied to VBA semantics
```

The corpus is already strong enough on ordinary syntax. The remaining value is in preventing **mode confusion**:

```text
live VBE source vs exported module text
syntax validity vs Excel host semantics
visible code vs hidden designer metadata
parse error vs realtime incomplete state
parser state vs completion state
keyword casing vs symbol casing
MS-VBAL spec result vs live Excel compiler result
```

That is the final button-up layer.
