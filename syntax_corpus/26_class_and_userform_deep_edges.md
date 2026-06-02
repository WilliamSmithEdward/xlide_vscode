# XLIDE VBA Corpus Addendum: Class and UserForm Deep Edges

**Recommended file name:** `26_class_and_userform_deep_edges.md`

**Purpose:** Add focused class-module and UserForm-module coverage to the XLIDE Excel VBA realtime analysis corpus.

The existing corpus covers module-kind awareness, core syntax, realtime recovery, legacy syntax, limits, host-aware warnings, and some basic class/UserForm examples. This file adds the deeper object-module cases that usually break static analysis, IntelliSense, symbol resolution, and false-positive control.

**Target host:** Excel VBA.

**Primary grammar oracle:** Microsoft **MS-VBAL**.

**Behavior oracle:** live Excel VBE compile/runtime canary.

**Rule:** Do not infer class or UserForm behavior from VB.NET, VBScript, TypeScript, Python, or generic BASIC.

---

## Diagnostic Classes

Use the same diagnostic classes as the rest of the corpus.

```text
syntax-error              cannot parse as valid VBA
compile-error             parseable, but VBE compiler rejects it
semantic-warning          suspicious or symbol-dependent
host-warning              Excel-specific risk, not pure VBA grammar
style-warning             valid but discouraged
incomplete                valid partial state during active typing
completion-context        parser state for IntelliSense, not a diagnostic
needs-verification        must be resolved by MS-VBAL and/or VBE canary
designer-symbol-required  valid only when .frm designer metadata provides the symbol
```

---

## Fixture Shape

```jsonc
{
  "id": "CLASS_LIFE_001",
  "title": "Class Initialize event",
  "moduleKind": "class",
  "moduleName": "Class1",
  "expected": "valid",
  "phase": "parser",
  "diagnostics": [],
  "snippet": "..."
}
```

For UserForm designer-backed symbols:

```jsonc
{
  "id": "FORM_SYMBOL_001",
  "title": "Known UserForm control symbol",
  "moduleKind": "userform",
  "moduleName": "UserForm1",
  "designerSymbols": [
    {
      "name": "CommandButton1",
      "type": "MSForms.CommandButton"
    }
  ],
  "expected": "valid",
  "phase": "scope",
  "diagnostics": [],
  "snippet": "..."
}
```

---

# 1. Class Lifecycle Procedures

## CLASS_LIFE_001 valid: Class Initialize

```vba
Option Explicit

Private Sub Class_Initialize()
    Debug.Print "created"
End Sub
```

Expected:

```text
valid in class module
host/module-context warning if found in standard module
```

---

## CLASS_LIFE_002 valid: Class Terminate

```vba
Option Explicit

Private Sub Class_Terminate()
    Debug.Print "destroyed"
End Sub
```

Expected:

```text
valid in class module
host/module-context warning if found in standard module
```

---

## CLASS_LIFE_003 warning: Class Initialize with wrong signature

```vba
Option Explicit

Private Sub Class_Initialize(ByVal value As Long)
End Sub
```

Expected:

```text
syntax: valid procedure
class event binding: invalid signature
diagnostic: semantic-warning or compile-error after VBE canary
```

Canary note:

```text
Verify exact Excel VBE compile behavior.
```

---

## CLASS_LIFE_004 warning: Class Terminate with wrong visibility

```vba
Option Explicit

Public Sub Class_Terminate()
End Sub
```

Expected:

```text
syntax: valid procedure
class event binding: suspicious or invalid
verify with VBE canary
```

---

# 2. UserForm Lifecycle Procedures

## FORM_LIFE_001 valid: UserForm Initialize

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Ready"
End Sub
```

Expected:

```text
valid in userform module
context warning if found outside userform module
```

---

## FORM_LIFE_002 valid: UserForm Activate

```vba
Option Explicit

Private Sub UserForm_Activate()
    Debug.Print Me.Caption
End Sub
```

Expected:

```text
valid in userform module
```

---

## FORM_LIFE_003 valid: UserForm QueryClose

```vba
Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Debug.Print "closed from X button"
    End If
End Sub
```

Expected:

```text
valid in userform module
signature-sensitive event procedure
```

---

## FORM_LIFE_004 valid: UserForm Terminate

```vba
Option Explicit

Private Sub UserForm_Terminate()
    Debug.Print "form terminated"
End Sub
```

Expected:

```text
valid in userform module
```

---

## FORM_LIFE_005 warning: UserForm event in standard module

```vba
Option Explicit

Private Sub UserForm_Initialize()
End Sub
```

Expected with `moduleKind: standard`:

```text
syntax: valid
host/module-context warning
not syntax error
```

Expected with `moduleKind: userform`:

```text
valid
```

---

# 3. `Me` Semantics Across Module Kinds

## ME_001 valid: Me in class module

```vba
Option Explicit

Private mName As String

Public Property Get Name() As String
    Name = Me.GetName()
End Property

Private Function GetName() As String
    GetName = mName
End Function
```

Expected:

```text
valid in class module
Me resolves to current class instance
```

---

## ME_002 valid: Me in UserForm module

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "XLIDE"
End Sub
```

Expected:

```text
valid in userform module
Me resolves to UserForm instance
```

---

## ME_003 valid: Me in worksheet module

```vba
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Range("A1").Value = Target.Address
End Sub
```

Expected:

```text
valid in worksheet module
Me resolves to worksheet instance
```

---

## ME_004 invalid: Me in standard module

```vba
Option Explicit

Public Sub Demo()
    Debug.Print Me.Name
End Sub
```

Expected:

```text
syntax: parseable
compile-error or semantic diagnostic
Me is not valid in standard module
```

---

# 4. Friend Visibility

## FRIEND_001 valid: Friend procedure in class module

```vba
Option Explicit

Friend Sub InternalOnly()
    Debug.Print "friend"
End Sub
```

Expected:

```text
valid in class module
```

---

## FRIEND_002 valid: Friend procedure in UserForm module

```vba
Option Explicit

Friend Sub InternalOnly()
    Debug.Print "friend"
End Sub
```

Expected:

```text
valid in userform module if confirmed by VBE canary
```

---

## FRIEND_003 invalid: Friend procedure in standard module

```vba
Option Explicit

Friend Sub InternalOnly()
    Debug.Print "friend"
End Sub
```

Expected with `moduleKind: standard`:

```text
compile-error
Friend is only valid in object modules
```

---

## FRIEND_004 invalid: Friend variable declaration

```vba
Option Explicit

Friend mValue As Long
```

Expected:

```text
compile-error
Friend modifies procedures, not variables
verify exact diagnostic with VBE canary
```

---

# 5. Public and Private Members in Object Modules

## OBJECT_MEMBER_001 valid: Public class method

```vba
Option Explicit

Public Sub Save()
    Debug.Print "save"
End Sub
```

Expected:

```text
valid in class module
public member exposed on class interface
```

---

## OBJECT_MEMBER_002 valid: Private class state

```vba
Option Explicit

Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property
```

Expected:

```text
valid in class module
```

---

## OBJECT_MEMBER_003 warning: Public field in class module

```vba
Option Explicit

Public Value As Long
```

Expected:

```text
syntax: valid
optional style-warning
public fields expose mutable state directly
```

---

## OBJECT_MEMBER_004 valid: Private UserForm state

```vba
Option Explicit

Private mIsLoaded As Boolean

Private Sub UserForm_Initialize()
    mIsLoaded = True
End Sub
```

Expected:

```text
valid in userform module
```

---

# 6. Events and RaiseEvent

## EVENT_001 valid: Event declaration in class module

```vba
Option Explicit

Public Event Saved(ByVal fileName As String)

Public Sub Save(ByVal fileName As String)
    RaiseEvent Saved(fileName)
End Sub
```

Expected:

```text
valid in class module
Event declaration accepted
RaiseEvent accepted
```

---

## EVENT_002 invalid: Event declaration inside procedure

```vba
Option Explicit

Public Sub Demo()
    Public Event Saved()
End Sub
```

Expected:

```text
compile-error
Event declarations are module-level only
```

---

## EVENT_003 invalid: RaiseEvent for undeclared event

```vba
Option Explicit

Public Sub Demo()
    RaiseEvent Saved()
End Sub
```

Expected:

```text
syntax: valid shape
semantic/project diagnostic: event not declared in this module
verify VBE compile behavior
```

---

## EVENT_004 invalid or module-specific: Event in standard module

```vba
Option Explicit

Public Event Saved()
```

Expected with `moduleKind: standard`:

```text
verify with VBE canary
likely invalid outside object module
```

Expected with `moduleKind: class`:

```text
valid
```

---

## EVENT_005 valid: Private event declaration

```vba
Option Explicit

Private Event Changed()

Private Sub NotifyChanged()
    RaiseEvent Changed
End Sub
```

Expected:

```text
verify with VBE canary
if accepted, valid in class module
```

---

# 7. WithEvents Declarations

## WITHEVENTS_001 valid: WithEvents object variable in class module

```vba
Option Explicit

Private WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print Wb.Name
End Sub
```

Expected:

```text
valid in class module
handler name binds to WithEvents variable name plus event name
```

---

## WITHEVENTS_002 valid: WithEvents in UserForm module

```vba
Option Explicit

Private WithEvents App As Application

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Debug.Print Wb.Name
End Sub
```

Expected:

```text
valid in userform module if VBE accepts
```

---

## WITHEVENTS_003 invalid: WithEvents in standard module

```vba
Option Explicit

Private WithEvents App As Application
```

Expected with `moduleKind: standard`:

```text
compile-error
WithEvents not valid in standard module
```

---

## WITHEVENTS_004 invalid: WithEvents local variable

```vba
Option Explicit

Public Sub Demo()
    Dim WithEvents App As Application
End Sub
```

Expected:

```text
compile-error
WithEvents is not valid for local variables
```

---

## WITHEVENTS_005 invalid: WithEvents As New

```vba
Option Explicit

Private WithEvents App As New Application
```

Expected:

```text
compile-error
WithEvents variables cannot use As New
verify exact diagnostic with VBE canary
```

---

## WITHEVENTS_006 invalid: WithEvents array

```vba
Option Explicit

Private WithEvents Apps(1 To 3) As Application
```

Expected:

```text
compile-error
WithEvents variable cannot be an array
verify exact diagnostic with VBE canary
```

---

## WITHEVENTS_007 warning: handler for unknown WithEvents source

```vba
Option Explicit

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print Wb.Name
End Sub
```

Expected:

```text
syntax: valid procedure
semantic warning if no WithEvents App variable exists
not syntax error
```

---

# 8. Implements

## IMPLEMENTS_001 valid: Implements statement in class module

```vba
Option Explicit

Implements IFoo

Private Sub IFoo_DoWork()
    Debug.Print "work"
End Sub
```

Expected:

```text
valid syntax in class module
full semantic validation requires project-wide IFoo interface definition
```

---

## IMPLEMENTS_002 invalid: Implements inside procedure

```vba
Option Explicit

Public Sub Demo()
    Implements IFoo
End Sub
```

Expected:

```text
compile-error
Implements is module-level only
```

---

## IMPLEMENTS_003 invalid: Implements in standard module

```vba
Option Explicit

Implements IFoo
```

Expected with `moduleKind: standard`:

```text
compile-error
Implements only valid in class/object modules
```

---

## IMPLEMENTS_004 project diagnostic: missing interface member implementation

Interface class module `IFoo`:

```vba
Option Explicit

Public Sub DoWork()
End Sub
```

Implementation class module:

```vba
Option Explicit

Implements IFoo
```

Expected:

```text
syntax: valid
project semantic diagnostic: IFoo_DoWork implementation missing
requires project-wide symbol analysis
```

---

## IMPLEMENTS_005 project diagnostic: signature mismatch

Interface class module `IFoo`:

```vba
Option Explicit

Public Sub DoWork(ByVal value As Long)
End Sub
```

Implementation class module:

```vba
Option Explicit

Implements IFoo

Private Sub IFoo_DoWork(ByVal value As String)
End Sub
```

Expected:

```text
syntax: valid
project semantic diagnostic: implemented member signature mismatch
verify exact VBE compile behavior
```

---

# 9. Property Procedure Consistency

## PROP_CLASS_001 valid: Get and Let pair

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

Expected:

```text
valid in class module
```

---

## PROP_CLASS_002 valid: Object Get and Set pair

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

Expected:

```text
valid in class module
object property uses Set for assignment
```

---

## PROP_CLASS_003 invalid: Object property with Let

```vba
Option Explicit

Private mSheet As Worksheet

Public Property Let Sheet(ByVal value As Worksheet)
    Set mSheet = value
End Property
```

Expected:

```text
compile-error or semantic diagnostic
object property assignment should use Property Set
verify exact VBE behavior
```

---

## PROP_CLASS_004 invalid: Scalar property with Set

```vba
Option Explicit

Private mName As String

Public Property Set Name(ByVal value As String)
    mName = value
End Property
```

Expected:

```text
compile-error or semantic diagnostic
scalar property assignment should use Property Let
verify exact VBE behavior
```

---

## PROP_CLASS_005 invalid: Get and Let type mismatch

```vba
Option Explicit

Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal value As String)
    mValue = CLng(value)
End Property
```

Expected:

```text
compile-error
property procedures for same property are inconsistent
```

---

## PROP_CLASS_006 invalid: Access mismatch across property procedures

```vba
Option Explicit

Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Private Property Let Name(ByVal value As String)
    mName = value
End Property
```

Expected:

```text
verify with VBE canary
likely inconsistent property procedure definitions
```

---

## PROP_CLASS_007 valid: Write-only property

```vba
Option Explicit

Private mSecret As String

Public Property Let Secret(ByVal value As String)
    mSecret = value
End Property
```

Expected:

```text
valid if confirmed
optional style-warning for write-only property
```

---

# 10. UserForm Designer-Backed Symbols

## FORM_SYMBOL_001 valid: known UserForm control symbol

Designer metadata:

```jsonc
{
  "designerSymbols": [
    { "name": "CommandButton1", "type": "MSForms.CommandButton" }
  ]
}
```

Code:

```vba
Option Explicit

Private Sub CommandButton1_Click()
    CommandButton1.Caption = "OK"
End Sub
```

Expected:

```text
valid when designerSymbols contains CommandButton1
CommandButton1 resolved as implicit control member
no undeclared variable diagnostic
```

---

## FORM_SYMBOL_002 valid: Me-qualified known control

Designer metadata:

```jsonc
{
  "designerSymbols": [
    { "name": "TextBox1", "type": "MSForms.TextBox" }
  ]
}
```

Code:

```vba
Option Explicit

Private Sub TextBox1_Change()
    Me.TextBox1.Text = UCase$(Me.TextBox1.Text)
End Sub
```

Expected:

```text
valid when designerSymbols contains TextBox1
Me resolves to UserForm instance
TextBox1 resolves as designer-backed member
```

---

## FORM_SYMBOL_003 warning: unknown control symbol with designer metadata present

Designer metadata:

```jsonc
{
  "designerSymbols": [
    { "name": "CommandButton1", "type": "MSForms.CommandButton" }
  ]
}
```

Code:

```vba
Option Explicit

Private Sub CommandButton2_Click()
    CommandButton2.Caption = "OK"
End Sub
```

Expected:

```text
syntax: valid
semantic diagnostic: CommandButton2 unresolved if designer metadata is trusted
not syntax error
```

---

## FORM_SYMBOL_004 tolerant: no designer metadata loaded

Code:

```vba
Option Explicit

Private Sub TextBox1_Change()
    TextBox1.Text = UCase$(TextBox1.Text)
End Sub
```

Expected:

```text
syntax: valid
if designer metadata unavailable, avoid hard unresolved-symbol error
optional unknown-context warning only
```

---

## FORM_SYMBOL_005 valid: control event with exact designer-backed control

Designer metadata:

```jsonc
{
  "designerSymbols": [
    { "name": "ListBox1", "type": "MSForms.ListBox" }
  ]
}
```

Code:

```vba
Option Explicit

Private Sub ListBox1_Click()
    Debug.Print ListBox1.ListIndex
End Sub
```

Expected:

```text
valid with designer metadata
event procedure recognized as control event
```

---

## FORM_SYMBOL_006 warning: plausible control event without known control

```vba
Option Explicit

Private Sub UnknownControl_Click()
    Debug.Print "clicked"
End Sub
```

Expected:

```text
syntax: valid
if designer metadata says UnknownControl does not exist, semantic diagnostic
if metadata unavailable, unknown-control warning or no diagnostic
```

---

# 11. UserForm Show, Load, Hide, and Unload Patterns

## FORM_USE_001 valid: Show form default instance

```vba
Public Sub Demo()
    UserForm1.Show
End Sub
```

Expected:

```text
valid if UserForm1 module exists
project symbol resolution required
```

---

## FORM_USE_002 valid: Load default instance

```vba
Public Sub Demo()
    Load UserForm1
    UserForm1.Show
End Sub
```

Expected:

```text
valid if UserForm1 module exists
Load statement recognized
```

---

## FORM_USE_003 valid: Unload form instance

```vba
Public Sub Demo()
    Unload UserForm1
End Sub
```

Expected:

```text
valid if UserForm1 module exists
Unload statement recognized
```

---

## FORM_USE_004 valid: Me.Hide inside UserForm

```vba
Private Sub CommandButton1_Click()
    Me.Hide
End Sub
```

Expected:

```text
valid in userform module
requires designer metadata only for CommandButton1 event binding, not for Me.Hide
```

---

## FORM_USE_005 valid: explicit UserForm instance

```vba
Public Sub Demo()
    Dim frm As UserForm1
    Set frm = New UserForm1
    frm.Show
    Unload frm
End Sub
```

Expected:

```text
valid if UserForm1 type exists
project symbol resolution required
```

---

## FORM_USE_006 warning: unknown form default instance

```vba
Public Sub Demo()
    MissingForm.Show
End Sub
```

Expected:

```text
syntax: valid
project semantic diagnostic if MissingForm is not a known module/type/symbol
not parser error
```

---

# 12. Default Instances and Class Instancing

## DEFAULT_INSTANCE_001 valid: UserForm default instance

```vba
Public Sub Demo()
    UserForm1.Caption = "Default"
    UserForm1.Show
End Sub
```

Expected:

```text
valid if UserForm1 exists
UserForm default instance behavior recognized
```

---

## DEFAULT_INSTANCE_002 invalid or warning: ordinary class default instance assumption

```vba
Public Sub Demo()
    Class1.DoWork
End Sub
```

Expected:

```text
if Class1 is a normal class module, this is not equivalent to UserForm default instance behavior
diagnostic depends on project metadata and instancing rules
verify VBE behavior
```

---

## DEFAULT_INSTANCE_003 valid: explicit class instance

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

# 13. Exported `.cls` and `.frm` Metadata Separation

## EXPORT_META_001 valid internal IO: class attributes

Exported `.cls` text:

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Demo()
End Sub
```

Expected:

```text
valid in internal-io/exported-source mode
not all lines are VBE-visible code
metadata must be preserved for roundtrip
do not show user-facing diagnostics for hidden metadata in live editor mode
```

---

## EXPORT_META_002 valid internal IO: UserForm metadata

Exported `.frm` text:

```vba
VERSION 5.00
Begin VB.UserForm UserForm1
   Caption         =   "Demo"
End
Attribute VB_Name = "UserForm1"
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Ready"
End Sub
```

Expected:

```text
valid in internal-io/exported-source mode
designer metadata parsed separately from code pane text
designer symbols extracted where possible
metadata preserved for roundtrip
```

---

## EXPORT_META_003 invalid live editor: pasted VERSION block

Live editor text:

```vba
VERSION 5.00
Begin VB.UserForm UserForm1
End
```

Expected:

```text
invalid in live VBE-visible code mode
valid only in exported .frm parsing mode
```

---

## EXPORT_META_004 valid exported member attribute

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
valid only in exported-module metadata mode if VBE import/export confirms it
invalid or hidden in live editor mode
do not conflate source modes
```

---

# 14. Realtime Incomplete Object-Module States

## RT_CLASS_001 incomplete: partial property

```vba
Option Explicit

Public Property Get Name() As String
```

Expected in realtime mode:

```text
incomplete
missing End Property should not cascade into whole file if user is actively typing
```

Expected on save/full analysis:

```text
invalid
missing End Property
```

---

## RT_CLASS_002 incomplete: partial event declaration

```vba
Option Explicit

Public Event Changed(
```

Expected in realtime mode:

```text
incomplete parameter list
no diagnostic flood
```

---

## RT_CLASS_003 incomplete: partial Implements

```vba
Option Explicit

Implements
```

Expected in realtime mode:

```text
incomplete
expected interface name
```

---

## RT_FORM_001 incomplete: partial control handler body

```vba
Option Explicit

Private Sub CommandButton1_Click()
    Me.
End Sub
```

Expected in realtime mode:

```text
incomplete member access
completion context should still be produced
no full procedure failure
```

---

## RT_FORM_002 incomplete: partial Load statement

```vba
Public Sub Demo()
    Load
End Sub
```

Expected in realtime mode:

```text
incomplete
expected object/form expression after Load
```

Expected on save/full analysis:

```text
invalid or compile-error
```

---

# 15. Completion Contexts for Class and UserForm Modules

Use marker syntax:

```text
<|>
```

The marker should be stripped before parsing and used as the completion request position.

---

## COMP_CLASS_001 Me completion in class

```vba
Option Explicit

Private mValue As Long

Public Sub Demo()
    Me.<|>
End Sub
```

Expected:

```text
moduleKind: class
completion context produced
receiver: Me
receiver type: current class
```

---

## COMP_CLASS_002 private member completion in class

```vba
Option Explicit

Private mValue As Long

Public Sub Demo()
    Debug.Print Me.<|>
End Sub
```

Expected:

```text
completion includes visible members from current class
private members are visible from inside same class
```

---

## COMP_FORM_001 Me completion in UserForm

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.<|>
End Sub
```

Expected:

```text
moduleKind: userform
completion context produced
receiver: Me
receiver type: UserForm1/current form
include designer-backed controls if metadata loaded
```

---

## COMP_FORM_002 designer control completion

Designer metadata:

```jsonc
{
  "designerSymbols": [
    { "name": "CommandButton1", "type": "MSForms.CommandButton" }
  ]
}
```

Code:

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.CommandButton1.<|>
End Sub
```

Expected:

```text
completion context produced
receiver: Me.CommandButton1
receiver type: MSForms.CommandButton
```

---

## COMP_FORM_003 unknown designer control completion

Code:

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.UnknownControl.<|>
End Sub
```

Expected:

```text
syntax: incomplete member completion context
if designer metadata loaded and control absent, unresolved member warning
completion should not crash
```

---

# 16. Suggested VBE Canary Matrix

Run the following cases through the live Excel VBE canary and record verdicts.

```text
CLASS_LIFE_003     Class_Initialize with parameters
CLASS_LIFE_004     Public Class_Terminate
FRIEND_002         Friend in UserForm
FRIEND_004         Friend variable
EVENT_004          Event in standard module
EVENT_005          Private Event
WITHEVENTS_002     WithEvents in UserForm
WITHEVENTS_005     WithEvents As New
WITHEVENTS_006     WithEvents array
PROP_CLASS_003     Object property with Let
PROP_CLASS_004     Scalar property with Set
PROP_CLASS_006     Property access mismatch
DEFAULT_INSTANCE_002 ordinary class default instance assumption
EXPORT_META_004    exported member attribute placement
```

Recommended canary record:

```jsonc
{
  "id": "WITHEVENTS_005",
  "host": "Excel",
  "hostVersion": "recorded-by-runner",
  "bitness": "recorded-by-runner",
  "moduleKind": "class",
  "sourceMode": "live-vbe",
  "compileResult": "accept | reject",
  "runtimeResult": "not-run | no-error | runtime-error",
  "errorNumber": null,
  "errorMessageContains": null,
  "notes": ""
}
```

---

# 17. Integration Guidance

## Do Not Treat These as Generic Parser Tests

These cases require module context.

The same code may be:

```text
valid in class module
valid in userform module
valid in worksheet module
warning in standard module
invalid in live editor mode
valid in exported-source mode
```

Do not flatten that into one global valid/invalid result.

---

## Required Analyzer State

To handle this file well, the analyzer needs:

```text
moduleKind
moduleName
sourceMode: live-vbe | exported-source | internal-io
project symbol table
known class modules
known UserForm modules
known worksheet/workbook modules
designer symbols for UserForms
WithEvents declarations
Implements relationships
property procedure groups
active realtime parse state
```

---

## False Positive Rule

If designer metadata or project metadata is unavailable, downgrade diagnostics.

Example:

```vba
Private Sub CommandButton1_Click()
    CommandButton1.Caption = "OK"
End Sub
```

Without `.frm` designer metadata, this should not become a hard error.

Use:

```text
unknown-context warning
```

not:

```text
undeclared variable compile-error
```

unless project metadata proves the control does not exist.

---

## Most Important Boundary

Class/UserForm analysis is not just syntax.

It is the intersection of:

```text
VBA grammar
module kind
source mode
project symbols
designer metadata
VBE compile behavior
runtime behavior
```

That is why this file belongs late in the corpus, after the base parser is already stable.
