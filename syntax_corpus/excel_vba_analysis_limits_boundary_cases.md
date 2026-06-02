# Excel VBA Realtime Analysis Addendum: Limits, Boundaries, Names, and Continuations

**Purpose:** Add hard-boundary fixtures for Excel VBA realtime analysis in VS Code.

**Audience:** LLM agent implementing or validating the XLIDE-style Excel VBA parser, lexer, semantic checker, and host-aware diagnostics.

**Non-negotiable rule:** Verify all grammar-sensitive behavior against Microsoft **MS-VBAL** and, where relevant, Microsoft Office VBA language docs. Do not infer these limits from VB.NET, VBScript, TypeScript, Python, or generic BASIC.

Official references to verify before implementation:

- MS-VBAL landing page: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74
- Office VBA line length: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/line-too-long
- Office VBA naming rules: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules
- Office VBA string data type: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/string-data-type
- MS-VBAL modules: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/4599fae2-3f41-4e70-968e-2398741f446b
- Office VBA arrays: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/too-many-dimensions
- Office VBA too many arguments: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/too-many-arguments
- Excel specifications and limits: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3

---

## Core distinction for the agent

The analyzer must keep these categories separate:

1. **VBA source text limits**: physical source line length, logical line continuation count, malformed `_` continuation.
2. **VBA runtime value limits**: `String` variables can be much larger than 255 characters.
3. **VBA identifier limits**: variable, procedure, argument, constant, and module names have separate rules.
4. **Excel host limits**: formula length, worksheet function argument count, cell content, etc. These are host-aware warnings, not pure parser errors.

The most common bug to avoid: treating a long runtime string as if VBA has a 255-character `String` limit. That is wrong. Some Excel APIs and UI surfaces have smaller limits, but a VBA variable-length `String` does not.

---

# A. Line continuation coverage

## CONT_LIMIT_001 valid: normal explicit continuation

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello " & _
                "world"
End Sub
```

Expected: valid.

Reason: `_` continuation must be parsed before statement grammar. The newline after `_` does not terminate the logical statement.

---

## CONT_LIMIT_002 valid: procedure call with named arguments across lines

```vba
Option Explicit

Public Sub Demo()
    MsgBox _
        Prompt:="Hello", _
        Buttons:=vbInformation, _
        Title:="XLIDE"
End Sub
```

Expected: valid.

Realtime behavior: while typing after `Prompt:=`, classify as `incomplete`, not invalid.

---

## CONT_LIMIT_003 invalid: missing whitespace before underscore

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello" &_
                "world"
End Sub
```

Expected: invalid.

Diagnostic suggestion: `VBA_LINE_CONTINUATION_REQUIRES_PRECEDING_SPACE`.

---

## CONT_LIMIT_004 invalid: text after continuation underscore

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello" & _ ' bad trailing comment
                "world"
End Sub
```

Expected: invalid in VBA.

Important: do not import VB.NET's newer permissive behavior here. Verify exact VBA behavior.

---

## CONT_LIMIT_005 invalid: continuation inside a string literal

```vba
Option Explicit

Public Sub Demo()
    Dim s As String
    s = "hello _
world"
End Sub
```

Expected: invalid.

Reason: `_` inside a string literal is just text. It cannot continue the literal. The first physical line has an unterminated string.

---

## CONT_LIMIT_006 valid: continued member access chain

```vba
Option Explicit

Public Sub Demo()
    ThisWorkbook.Worksheets("Sheet1") _
        .Range("A1") _
        .Resize(10, 3) _
        .Value = "ok"
End Sub
```

Expected: valid.

Reason: this is critical for Excel object model code. The parser must bind the continued leading-dot member access to the previous expression.

---

## CONT_LIMIT_007 incomplete: user has just typed trailing continuation

```vba
Option Explicit

Public Sub Demo()
    Debug.Print "hello" & _
```

Expected: incomplete in realtime mode, invalid on save or full analysis.

Diagnostic mode:

```text
onType: incomplete, no red squiggle yet
onSave: error, expected continuation expression
```

---

## CONT_LIMIT_008 generated boundary: maximum continuation count

Microsoft's Office VBA docs state that up to 25 physical lines can be joined with line-continuation characters, meaning 24 consecutive line-continuation characters.

Implement generated fixtures instead of hand-writing these.

```ts
function makeContinuationExpression(continuations: number): string {
  const lines: string[] = [];
  lines.push("Option Explicit");
  lines.push("");
  lines.push("Public Sub Demo()");
  lines.push("    Dim s As String");
  lines.push("    s = \"0\" & _");

  for (let i = 1; i < continuations; i++) {
    lines.push(`        \"${i}\" & _`);
  }

  lines.push(`        \"${continuations}\"`);
  lines.push("End Sub");
  return lines.join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "CONT_LIMIT_008A",
    "title": "24 continuation characters, 25 physical lines",
    "continuations": 24,
    "expected": "valid",
    "phase": "parser"
  },
  {
    "id": "CONT_LIMIT_008B",
    "title": "25 continuation characters, 26 physical lines",
    "continuations": 25,
    "expected": "invalid",
    "phase": "parser",
    "diagnostics": [
      {
        "code": "VBA_TOO_MANY_LINE_CONTINUATIONS",
        "severity": "error"
      }
    ]
  }
]
```

---

# B. Physical and logical line length limits

## LINE_LIMIT_001 generated boundary: 1023-character physical line

Microsoft's Office VBA docs state that a physical line of Visual Basic code can contain up to 1023 characters.

Do not hand-write the line. Generate it.

```ts
function makePhysicalLineLengthCase(length: number): string {
  const prefix = "Public Sub Demo(): Debug.Print \"";
  const suffix = "\": End Sub";
  const fillLength = length - prefix.length - suffix.length;

  if (fillLength < 0) {
    throw new Error("Requested length is too short for fixture shape");
  }

  return prefix + "x".repeat(fillLength) + suffix;
}
```

Required generated tests:

```jsonc
[
  {
    "id": "LINE_LIMIT_001A",
    "title": "Physical line exactly 1023 characters",
    "physicalLineLength": 1023,
    "expected": "valid",
    "phase": "lexical"
  },
  {
    "id": "LINE_LIMIT_001B",
    "title": "Physical line 1024 characters",
    "physicalLineLength": 1024,
    "expected": "invalid",
    "phase": "lexical",
    "diagnostics": [
      {
        "code": "VBA_LINE_TOO_LONG",
        "severity": "error"
      }
    ]
  }
]
```

---

## LINE_LIMIT_002 generated boundary: logical line length through continuation

Microsoft's Office VBA docs state that 25 physical lines can be joined to form one logical line, and that such a logical line could potentially contain 10,230 characters.

Generated tests:

```jsonc
[
  {
    "id": "LINE_LIMIT_002A",
    "title": "Maximum legal logical line via continuations",
    "expected": "valid",
    "phase": "lexical"
  },
  {
    "id": "LINE_LIMIT_002B",
    "title": "Logical line exceeds continuation count",
    "expected": "invalid",
    "phase": "lexical",
    "diagnostics": [
      {
        "code": "VBA_TOO_MANY_LINE_CONTINUATIONS",
        "severity": "error"
      }
    ]
  }
]
```

Implementation note: physical line length and continuation count are separate diagnostics. Do not collapse both into generic parse failure.

---

# C. String size and string-literal coverage

## STRING_LIMIT_001 valid: variable-length string larger than 255 characters

```vba
Option Explicit

Public Sub Demo()
    Dim s As String
    s = String$(300, "x")
    Debug.Print Len(s)
End Sub
```

Expected: valid.

Important: no diagnostic saying `String` is limited to 255 characters.

---

## STRING_LIMIT_002 valid: concatenated literal larger than 255 characters

```vba
Option Explicit

Public Sub Demo()
    Dim s As String
    s = String$(200, "a") & String$(200, "b")
    Debug.Print Len(s)
End Sub
```

Expected: valid.

Reason: runtime string length is not the same thing as source physical line length.

---

## STRING_LIMIT_003 invalid or warning: giant literal on one physical line over 1023 characters

Generate a single physical line longer than 1023 characters:

```vba
Public Sub Demo(): Dim s As String: s = "xxxxx...": End Sub
```

Expected: invalid due to source line length, not due to runtime string length.

Diagnostic suggestion: `VBA_LINE_TOO_LONG`, not `VBA_STRING_TOO_LONG`.

---

## STRING_LIMIT_004 valid: fixed-length string in standard module

```vba
Option Explicit

Private s As String * 10

Public Sub Demo()
    s = "abcdef"
    Debug.Print Len(s)
End Sub
```

Expected: valid in a standard module.

---

## STRING_LIMIT_005 invalid: fixed-length string of zero

```vba
Option Explicit

Public Sub Demo()
    Dim s As String * 0
End Sub
```

Expected: invalid.

Reason: fixed-length strings must have a positive length. Verify exact diagnostic against MS-VBAL and Office VBA behavior.

---

## STRING_LIMIT_006 generated boundary: fixed-length string size

MS-VBAL gives a fixed-length string range of 1 to 65,526 characters. Office VBA docs describe this as approximately 64K.

Required generated tests:

```jsonc
[
  {
    "id": "STRING_LIMIT_006A",
    "title": "Fixed-length string length 1",
    "snippet": "Public Sub Demo(): Dim s As String * 1: End Sub",
    "expected": "valid"
  },
  {
    "id": "STRING_LIMIT_006B",
    "title": "Fixed-length string length 65526",
    "snippet": "Public Sub Demo(): Dim s As String * 65526: End Sub",
    "expected": "valid"
  },
  {
    "id": "STRING_LIMIT_006C",
    "title": "Fixed-length string length 65527",
    "snippet": "Public Sub Demo(): Dim s As String * 65527: End Sub",
    "expected": "invalid"
  }
]
```

Caution: if host behavior differs, record the host result explicitly and do not weaken the parser silently.

---

## STRING_LIMIT_007 invalid in class module: Public fixed-length string

```vba
Option Explicit

Public NameBuffer As String * 32
```

Fixture metadata:

```jsonc
{
  "id": "STRING_LIMIT_007",
  "moduleKind": "class",
  "expected": "invalid",
  "phase": "declaration",
  "diagnostics": [
    {
      "code": "VBA_PUBLIC_FIXED_LENGTH_STRING_IN_OBJECT_MODULE",
      "severity": "error"
    }
  ]
}
```

Expected: invalid in class/object module. Validity may differ in standard modules.

---

# D. Identifier, variable, function, and Sub naming limits

## NAME_LIMIT_001 valid: 255-character variable name

Generate this. Do not hand-write it.

```ts
function makeIdentifier(length: number): string {
  if (length < 1) throw new Error("length must be positive");
  return "A" + "x".repeat(length - 1);
}

function makeVariableNameLengthCase(length: number): string {
  const name = makeIdentifier(length);
  return [
    "Option Explicit",
    "",
    "Public Sub Demo()",
    `    Dim ${name} As Long`,
    `    ${name} = 1`,
    "End Sub"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "NAME_LIMIT_001A",
    "title": "Variable name exactly 255 characters",
    "identifierLength": 255,
    "expected": "valid",
    "phase": "lexical"
  },
  {
    "id": "NAME_LIMIT_001B",
    "title": "Variable name 256 characters",
    "identifierLength": 256,
    "expected": "invalid",
    "phase": "lexical",
    "diagnostics": [
      {
        "code": "VBA_IDENTIFIER_TOO_LONG",
        "severity": "error"
      }
    ]
  }
]
```

---

## NAME_LIMIT_002 valid: 255-character Sub name

```ts
function makeSubNameLengthCase(length: number): string {
  const name = makeIdentifier(length);
  return [
    "Option Explicit",
    "",
    `Public Sub ${name}()`,
    "End Sub"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "NAME_LIMIT_002A",
    "title": "Sub name exactly 255 characters",
    "identifierLength": 255,
    "expected": "valid"
  },
  {
    "id": "NAME_LIMIT_002B",
    "title": "Sub name 256 characters",
    "identifierLength": 256,
    "expected": "invalid"
  }
]
```

---

## NAME_LIMIT_003 valid: 255-character Function name

```ts
function makeFunctionNameLengthCase(length: number): string {
  const name = makeIdentifier(length);
  return [
    "Option Explicit",
    "",
    `Public Function ${name}() As Long`,
    `    ${name} = 1`,
    "End Function"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "NAME_LIMIT_003A",
    "title": "Function name exactly 255 characters",
    "identifierLength": 255,
    "expected": "valid"
  },
  {
    "id": "NAME_LIMIT_003B",
    "title": "Function name 256 characters",
    "identifierLength": 256,
    "expected": "invalid"
  }
]
```

---

## NAME_RULE_001 invalid: identifier starts with digit

```vba
Option Explicit

Public Sub Demo()
    Dim 1value As Long
End Sub
```

Expected: invalid.

Diagnostic suggestion: `VBA_IDENTIFIER_MUST_START_WITH_LETTER`.

---

## NAME_RULE_002 invalid: identifier starts with underscore

```vba
Option Explicit

Public Sub Demo()
    Dim _value As Long
End Sub
```

Expected: invalid for VBA naming rules.

Important: VB.NET allows some forms that VBA does not. Do not import VB.NET identifier rules.

---

## NAME_RULE_003 invalid: forbidden punctuation inside identifier

```vba
Option Explicit

Public Sub Demo()
    Dim bad.name As Long
End Sub
```

Expected: invalid.

Other forbidden name characters to test: space, `.`, `!`, `@`, `&`, `$`, `#`.

---

## NAME_RULE_004 valid: type-declaration suffix is not part of the base name

```vba
Option Explicit

Public Sub Demo()
    Dim count&
    Dim label$
    Dim price@
    Dim ratio#

    count = 1
    label = "ok"
    price = 12.34
    ratio = 0.5
End Sub
```

Expected: valid.

Reason: type-declaration characters can follow an identifier, but they are not ordinary identifier characters.

---

## NAME_RULE_005 warning: shadowing intrinsic function name

```vba
Option Explicit

Public Sub Demo()
    Dim Left As Long
    Left = 1
    Debug.Print VBA.Left$("abc", 1)
End Sub
```

Expected: syntactically valid, warning.

Diagnostic suggestion: `VBA_NAME_SHADOWS_INTRINSIC`.

Important: shadowing is not the same thing as a parse error.

---

## NAME_RULE_006 invalid: duplicate names in same scope, case-insensitive

```vba
Option Explicit

Public Sub Demo()
    Dim value As Long
    Dim VALUE As Long
End Sub
```

Expected: invalid or semantic error.

Reason: VBA is not case-sensitive for identity, although it preserves declared capitalization.

---

## NAME_RULE_007 valid or warning: same name in different scopes

```vba
Option Explicit

Private value As Long

Public Sub Demo()
    Dim value As Long
    value = 1
End Sub
```

Expected: valid, optional shadowing warning.

---

# E. Module name limits

## MODULE_LIMIT_001 generated boundary: VB_Name length

MS-VBAL states that a module name has a maximum length of 31 characters.

```ts
function makeModuleNameLengthCase(length: number): string {
  const name = "M" + "x".repeat(length - 1);
  return [
    `Attribute VB_Name = "${name}"`,
    "Option Explicit",
    "",
    "Public Sub Demo()",
    "End Sub"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "MODULE_LIMIT_001A",
    "title": "Module name exactly 31 characters",
    "moduleNameLength": 31,
    "expected": "valid",
    "phase": "declaration"
  },
  {
    "id": "MODULE_LIMIT_001B",
    "title": "Module name 32 characters",
    "moduleNameLength": 32,
    "expected": "invalid",
    "phase": "declaration",
    "diagnostics": [
      {
        "code": "VBA_MODULE_NAME_TOO_LONG",
        "severity": "error"
      }
    ]
  }
]
```

---

# F. Procedure argument limits

## ARG_LIMIT_001 generated boundary: 60 procedure arguments

Office VBA docs state that a procedure can have only 60 arguments.

```ts
function makeSubWithArgumentCount(count: number): string {
  const args = Array.from({ length: count }, (_, i) => `ByVal a${i + 1} As Long`).join(", ");
  return [
    "Option Explicit",
    "",
    `Public Sub Demo(${args})`,
    "End Sub"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "ARG_LIMIT_001A",
    "title": "Procedure with 60 arguments",
    "argumentCount": 60,
    "expected": "valid",
    "phase": "declaration"
  },
  {
    "id": "ARG_LIMIT_001B",
    "title": "Procedure with 61 arguments",
    "argumentCount": 61,
    "expected": "invalid",
    "phase": "declaration",
    "diagnostics": [
      {
        "code": "VBA_TOO_MANY_ARGUMENTS",
        "severity": "error"
      }
    ]
  }
]
```

---

## ARG_LIMIT_002 valid: ParamArray as final argument

```vba
Option Explicit

Public Sub Demo(ByVal label As String, ParamArray values() As Variant)
End Sub
```

Expected: valid.

---

## ARG_LIMIT_003 invalid: ParamArray not final

```vba
Option Explicit

Public Sub Demo(ParamArray values() As Variant, ByVal label As String)
End Sub
```

Expected: invalid.

---

# G. Array dimension limits

## ARRAY_LIMIT_001 generated boundary: 60 dimensions

Office VBA docs state that arrays can have no more than 60 dimensions.

```ts
function makeArrayDimensionCase(dimensions: number): string {
  const subscripts = Array.from({ length: dimensions }, () => "1 To 1").join(", ");
  return [
    "Option Explicit",
    "",
    "Public Sub Demo()",
    `    Dim values(${subscripts}) As Long`,
    "End Sub"
  ].join("\n");
}
```

Required generated tests:

```jsonc
[
  {
    "id": "ARRAY_LIMIT_001A",
    "title": "Array with 60 dimensions",
    "dimensionCount": 60,
    "expected": "valid",
    "phase": "declaration"
  },
  {
    "id": "ARRAY_LIMIT_001B",
    "title": "Array with 61 dimensions",
    "dimensionCount": 61,
    "expected": "invalid",
    "phase": "declaration",
    "diagnostics": [
      {
        "code": "VBA_TOO_MANY_ARRAY_DIMENSIONS",
        "severity": "error"
      }
    ]
  }
]
```

---

# H. Excel host-aware limit warnings

These are not pure VBA syntax errors. They should be host-aware warnings or analyzer diagnostics.

## EXCEL_LIMIT_001 warning: formula text approaching or exceeding Excel formula limit

Excel specifications list formula contents at 8,192 characters and internal formula length at 16,384 bytes.

```vba
Option Explicit

Public Sub Demo()
    Range("A1").Formula = "=SUM(1,2,3)"
End Sub
```

Generated variants:

```jsonc
[
  {
    "id": "EXCEL_LIMIT_001A",
    "title": "Formula string length 8192",
    "formulaLength": 8192,
    "expected": "valid",
    "phase": "host"
  },
  {
    "id": "EXCEL_LIMIT_001B",
    "title": "Formula string length 8193",
    "formulaLength": 8193,
    "expected": "warning",
    "phase": "host",
    "diagnostics": [
      {
        "code": "EXCEL_FORMULA_TOO_LONG",
        "severity": "warning"
      }
    ]
  }
]
```

Important: this only applies when the analyzer can confidently identify a formula assignment. Do not flag arbitrary long VBA strings as Excel formulas.

---

## EXCEL_LIMIT_002 warning: worksheet function argument count

Excel specifications list 255 arguments in a worksheet function.

This matters for strings assigned to `.Formula`, not VBA procedure calls.

Generated variants:

```jsonc
[
  {
    "id": "EXCEL_LIMIT_002A",
    "title": "Formula function call with 255 arguments",
    "expected": "valid",
    "phase": "host"
  },
  {
    "id": "EXCEL_LIMIT_002B",
    "title": "Formula function call with 256 arguments",
    "expected": "warning",
    "phase": "host",
    "diagnostics": [
      {
        "code": "EXCEL_FORMULA_FUNCTION_TOO_MANY_ARGUMENTS",
        "severity": "warning"
      }
    ]
  }
]
```

---

# I. Realtime-specific acceptance rules

During active typing, these should generally be `incomplete`, not hard invalid:

```vba
Public Sub Demo()
    Debug.Print "hello" & _
```

```vba
Public Sub Demo(
```

```vba
Public Function CalculateValue(
```

```vba
Public Sub Demo()
    ThisWorkbook.Worksheets("Sheet1") _
```

```vba
Public Sub Demo()
    Range("A1").Formula = "=
```

Strict mode may promote these to invalid at EOF, on save, or on explicit analysis.

---

# J. Required implementation notes for the agent

1. Tokenize physical lines first so the lexer can enforce 1023-character source lines before higher-level parsing.
2. Treat `_` continuation as a physical-line feature, not an expression-level operator.
3. Track continuation depth on logical lines.
4. Keep string-literal state before checking `_`. A `_` inside a string is ordinary text, not continuation.
5. Treat runtime `String` length and source line length as separate systems.
6. Treat Excel formula limits as host-aware diagnostics, never parser errors.
7. Apply identifier rules to base identifiers before type-declaration suffixes.
8. Enforce case-insensitive identity for duplicate-name checks, while preserving declared capitalization for symbol display and completion.
9. Separate exported source parsing from live editor parsing for `Attribute VB_Name` and member `Attribute` metadata.
10. Use generated tests for boundaries. Do not rely on hand-written examples for 1023-character, 255-name, 60-argument, or 60-dimension cases.

---

# Minimal pass checklist

A analyzer passes this addendum only if it can correctly classify:

- `_` continuation with and without required whitespace
- `_` inside strings
- trailing comments after `_`
- 24 vs 25 consecutive continuations
- 1023 vs 1024 physical source-line length
- runtime strings larger than 255 characters
- fixed-length strings and class-module restrictions
- 255 vs 256 character identifiers
- module names 31 vs 32 characters
- procedure argument count 60 vs 61
- array dimensions 60 vs 61
- Excel formula limits as host warnings only

