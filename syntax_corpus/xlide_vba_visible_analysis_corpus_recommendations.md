# XLIDE VBA Visible Analysis Corpus Recommendations

## Purpose

This document defines the recommended remaining corpus files for testing realtime syntax analysis in XLIDE, focused on what a user actually sees and edits in the VBA IDE.

The core rule is simple:

> If the user cannot see it or edit it in XLIDE, do not produce user-facing diagnostics for it.

Hidden/exported metadata should still be tested, but only at the internal IO and round-trip layer.

---

## Recommended Corpus Structure

```text
corpus/
  vbe-visible/
    01_module_header.md
    02_declarations.md
    03_procedures.md
    04_control_flow.md
    05_strings_comments_literals.md
    06_line_continuation.md
    07_properties.md
    08_events.md
    09_error_handling_labels.md
    10_arrays_redim.md
    11_conditional_compilation.md
    12_limits_boundaries.md
    13_realtime_incomplete_states.md
    14_diagnostic_ranges.md
    15_scope_and_shadowing.md
    16_module_context_events.md
    17_excel_host_common_patterns.md

  internal-io-only/
    bas_attributes.md
    cls_attributes.md
    frm_metadata.md
    roundtrip_preservation.md
```

The `vbe-visible` folder is the user-facing analysis corpus.

The `internal-io-only` folder is not for realtime user diagnostics. It exists to make sure XLIDE can safely import, normalize, preserve, and export VBA modules without corrupting metadata.

---

# 1. `vbe_realtime_incomplete_states.md`

## Goal

Test what happens while the user is still typing.

Realtime analysis should not behave like a final compiler pass. It should tolerate partial thoughts, avoid diagnostic floods, and recover cleanly as soon as the user finishes the construct.

## Required Assertions

Each case should verify:

```text
does not crash
does not hang
does not flood diagnostics
diagnostic severity is appropriate for realtime mode
parser recovers by the next valid statement
parser recovers by the next procedure when needed
```

## Example Cases

### Incomplete member access

```vba
Public Sub Test()
    ThisWorkbook.
End Sub
```

Expected realtime behavior:

```text
soft incomplete diagnostic or no diagnostic
no full-procedure parse failure
no cascade errors
```

Expected strict behavior:

```text
VBA_REALTIME_INCOMPLETE_MEMBER_ACCESS or equivalent strict parse error
```

### Incomplete object assignment

```vba
Public Sub Test()
    Dim ws As Worksheet
    Set ws =
End Sub
```

Expected realtime behavior:

```text
soft incomplete expression diagnostic
no cascade into End Sub
```

### Incomplete If statement

```vba
Public Sub Test()
    If x Then
End Sub
```

Expected realtime behavior:

```text
incomplete block diagnostic
should not create unrelated errors on End Sub
```

### Incomplete For statement

```vba
Public Sub Test()
    For i =
End Sub
```

Expected realtime behavior:

```text
incomplete For initializer diagnostic
parser recovers by End Sub
```

### Incomplete function call

```vba
Public Sub Test()
    Range("A1"
End Sub
```

Expected realtime behavior:

```text
missing close parenthesis diagnostic
diagnostic range should be on call/paren area
```

### Unterminated string

```vba
Public Sub Test()
    Debug.Print "unterminated
End Sub
```

Expected realtime behavior:

```text
unterminated string diagnostic
diagnostic should not consume the whole rest of file
```

---

# 2. `vbe_diagnostic_ranges.md`

## Goal

Test squiggle placement.

A diagnostic can be logically correct but feel bad if the underline lands on the wrong token. This corpus should verify the diagnostic range, not only the diagnostic code.

## Required Assertions

Each case should verify:

```text
diagnostic code
diagnostic severity
diagnostic range
no excessive full-line squiggles when a precise token can be marked
```

Use marker syntax in fixtures if possible:

```text
<|bad token|>
```

The marker should be stripped before analysis, then used to compare expected ranges.

## Example Cases

### Bad line continuation

```vba
Public Sub Test()
    Debug.Print "hello" & <|_ ' invalid trailing comment|>
End Sub
```

Expected:

```text
code: VBA_PARSE_BAD_LINE_CONTINUATION
range: continuation token and/or trailing tokens
```

### Invalid identifier start

```vba
Public Sub Test()
    Dim <|1bad|> As Long
End Sub
```

Expected:

```text
code: VBA_DECL_INVALID_IDENTIFIER
range: 1bad
```

### Incomplete return type

```vba
Public Function Foo() As <| |>
End Function
```

Expected:

```text
code: VBA_PARSE_EXPECTED_TYPE_NAME
range: after As
```

### Bad Call syntax

```vba
Public Sub Test()
    <|Call MsgBox "Hello"|>
End Sub
```

Expected:

```text
code: VBA_PARSE_BAD_CALL_SYNTAX
range: call statement or missing parentheses area
```

### Invalid Set assignment

```vba
Public Sub Test()
    Dim x As Long
    <|Set x = 123|>
End Sub
```

Expected:

```text
code: VBA_SEMANTIC_SET_REQUIRES_OBJECT
range: Set statement
```

---

# 3. `vbe_scope_and_shadowing.md`

## Goal

Test names that are legal, confusing, ambiguous, or context-sensitive.

This file prevents the analyzer from treating every suspicious name as invalid syntax.

## Classification Rules

Each case should classify the result as one of:

```text
syntax error
compile-shape error
semantic/project error
Excel host warning
style warning
valid but suspicious
valid and no diagnostic
```

## Example Cases

### Shadowing Excel object member name

```vba
Public Sub Test()
    Dim Range As String
    Range = "hello"
End Sub
```

Expected:

```text
valid VBA
possible style warning only
not syntax error
```

### Shadowing built-in Date

```vba
Public Sub Test()
    Dim Date As Date
    Date = Now
End Sub
```

Expected:

```text
valid or compile-warning-equivalent depending on rule design
not parser failure
```

### Procedure name reused as local variable

```vba
Public Sub Test()
    Dim Test As Long
End Sub
```

Expected:

```text
project/scope diagnostic if applicable
not syntax error
```

### Duplicate private members in same module

```vba
Private Sub Foo()
End Sub

Private Function Foo() As Long
End Function
```

Expected:

```text
duplicate member diagnostic
not parser error
```

### Variant default in multi-declaration

```vba
Public Sub Test()
    Dim a, b As Long
End Sub
```

Expected:

```text
valid VBA
optional warning: a is Variant, b is Long
```

### Type suffix names

```vba
Public Sub Test()
    Dim count&
    Dim label$
    Dim amount@
    Dim ratio#
End Sub
```

Expected:

```text
valid VBA
no invalid-character diagnostic for suffixes
```

---

# 4. `vbe_module_context_events.md`

## Goal

Test event procedures that are syntactically valid but only correct in specific module contexts.

Events should not be judged by raw syntax alone.

## Fixture Metadata

Each test case should include:

```text
moduleKind: standard | worksheet | workbook | class | form
```

Optionally include:

```text
host: excel
moduleName: Sheet1 | ThisWorkbook | UserForm1 | Class1
```

## Example Cases

### Worksheet event in worksheet module

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
End Sub
```

Expected with `moduleKind: worksheet`:

```text
valid event procedure
```

Expected with `moduleKind: standard`:

```text
context diagnostic
not syntax error
```

### Workbook event in ThisWorkbook module

```vba
Private Sub Workbook_Open()
End Sub
```

Expected with `moduleKind: workbook`:

```text
valid event procedure
```

Expected with `moduleKind: standard`:

```text
context diagnostic
not syntax error
```

### UserForm control event

```vba
Private Sub CommandButton1_Click()
End Sub
```

Expected with `moduleKind: form` and known control:

```text
valid event procedure
```

Expected with unknown control ownership:

```text
unknown event target warning or no diagnostic
do not mark as syntax error
```

### Class event source pattern

```vba
Private WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
End Sub
```

Expected:

```text
valid in class-like module context
requires symbol/event-source awareness for full validation
```

---

# 5. `vbe_excel_host_common_patterns.md`

## Goal

Test common Excel VBA patterns users actually write.

The goal is to avoid noisy diagnostics on valid Excel VBA while still allowing optional host-aware warnings.

## Classification Rules

Each case should distinguish:

```text
valid VBA
valid Excel VBA
optional host warning
style warning
not syntax error
```

## Example Cases

### Unqualified Range

```vba
Public Sub Test()
    Range("A1").Value = 1
End Sub
```

Expected:

```text
valid Excel VBA
optional warning: unqualified Range
not syntax error
```

### Unqualified Cells

```vba
Public Sub Test()
    Cells(1, 1).Value = "x"
End Sub
```

Expected:

```text
valid Excel VBA
optional warning: unqualified Cells
not syntax error
```

### Rows deletion

```vba
Public Sub Test()
    Rows("1:1").Delete
End Sub
```

Expected:

```text
valid Excel VBA
optional warning if unqualified object rule enabled
```

### Qualified worksheet access

```vba
Public Sub Test()
    Sheets("Sheet1").Range("A1").Value = 123
End Sub
```

Expected:

```text
valid Excel VBA
optional warning: Sheets collection may refer to chart sheets too
not syntax error
```

### Fully qualified workbook and worksheet access

```vba
Public Sub Test()
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = 123
End Sub
```

Expected:

```text
valid Excel VBA
no unqualified Range warning
```

### Formula string

```vba
Public Sub Test()
    ActiveSheet.Range("A1").Formula = "=SUM(B1:B10)"
End Sub
```

Expected:

```text
valid Excel VBA
do not parse formula string as VBA
optional formula-length host warning only if enabled
```

### FormulaR1C1 string

```vba
Public Sub Test()
    ActiveSheet.Range("A2").FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
End Sub
```

Expected:

```text
valid Excel VBA
formula content is not VBA source
```

---

# 6. `internal-io-only/roundtrip_preservation.md`

## Goal

Test hidden/exported content that XLIDE may read or write but should not show to the user as editable VBA source.

This corpus is not for realtime user diagnostics.

## Rule

```text
If it is hidden from the user, do not analyze it as visible source.
If XLIDE writes it back, test that it is preserved or normalized intentionally.
```

## Example Hidden Content

```vba
Attribute VB_Name = "Module1"
Attribute VB_Description = "..."
Attribute VB_ProcData.VB_Invoke_Func = " \n14"
```

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
```

## Expected IO Behavior

```text
metadata is not shown in visible editor buffer
metadata does not produce user diagnostics
metadata is preserved on export when required
metadata is intentionally regenerated only when documented
```

---

# Fixture Metadata Recommendation

Use machine-readable fixture metadata even if the corpus is documented in Markdown.

Recommended fields:

```json
{
  "id": "string-unterminated-realtime",
  "title": "Unterminated string should not consume whole module",
  "moduleKind": "standard",
  "host": "excel",
  "mode": "realtime",
  "category": "realtime-recovery",
  "vbeVisible": true,
  "vbeCompileEquivalent": true,
  "code": "Public Sub Test()\n    Debug.Print \"unterminated\nEnd Sub",
  "expected": [
    {
      "code": "VBA_PARSE_UNTERMINATED_STRING",
      "severity": "error",
      "rangeMarker": 0
    }
  ],
  "assertions": [
    "doesNotCrash",
    "doesNotHang",
    "noDiagnosticFlood",
    "recoversByNextProcedure"
  ]
}
```

## Required Diagnostic Fields

Each expected diagnostic should eventually include:

```text
code
severity
category
range
message pattern, optional
vbeCompileEquivalent
hostSpecific
requiresProjectContext
```

Avoid relying on exact message text early. Diagnostic codes and ranges are more stable.

---

# Recommended Build Order

## Step 1

Add `13_realtime_incomplete_states.md`.

This gives the highest editor-quality payoff.

## Step 2

Add `14_diagnostic_ranges.md`.

This makes diagnostics feel precise instead of noisy.

## Step 3

Add `17_excel_host_common_patterns.md`.

This protects XLIDE from false-positive warnings on normal Excel VBA.

## Step 4

Add `16_module_context_events.md`.

This makes XLIDE feel aware of Excel modules and event contexts.

## Step 5

Add `15_scope_and_shadowing.md`.

This improves semantic quality and prepares the engine for project-level analysis.

## Step 6

Keep `internal-io-only` tests separate.

This protects round-tripping without polluting user-facing analysis behavior.

---

# Stop Rule

Do not keep adding random snippets indefinitely.

After these files exist, the next quality jump should come from:

```text
golden test harness
diagnostic range markers
realtime vs strict mode comparisons
negative parser recovery fuzzing
optional Excel COM / VBE oracle testing
```

The corpus is no longer the bottleneck once it covers user-visible syntax, realtime incomplete states, diagnostic precision, host-aware Excel patterns, and module-context-sensitive events.

At that point, the bottleneck becomes deterministic test execution and comparison against actual editor behavior.
