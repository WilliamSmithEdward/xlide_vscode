# XLIDE VBA Language Service Roadmap

## Purpose

Build the next layer of XLIDE: active linting, IntelliSense, symbol navigation, and safe formatting for VBA inside VS Code.

XLIDE already has the differentiated workbook/module IO layer. This roadmap is for the semantic layer above that IO: a VBA language service that understands modules, procedures, declarations, scopes, symbols, diagnostics, and keyword casing.

This document is written for another LLM or implementation agent. Follow it as a hard engineering plan, not as a loose idea list.

---

## Non-Negotiable Requirements

### 1. Verify against `MS-VBAL.pdf`

All lexer, parser, grammar, scoping, declaration, keyword, operator, literal, and syntax behavior must be verified against the official Microsoft VBA Language Specification, commonly referenced as `[MS-VBAL]: VBA Language Specification`.

Use the current Microsoft-published PDF as the canonical source:

- Microsoft Learn landing page: <https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74>
- PDF should be downloaded from the current Published Version row on that page.
- Store the checked PDF locally as `docs/spec/MS-VBAL.pdf`.
- Record the protocol revision and publication date in `docs/spec/MS-VBAL.version.md`.

Do not rely on memory, Visual Basic .NET behavior, internet snippets, Rubberduck behavior, or host-specific Excel behavior as the canonical definition of VBA syntax.

Every implemented grammar rule should have a nearby source note in code comments or in a spec mapping file.

Example:

```ts
// Verified against MS-VBAL.pdf, section: <section number/title>
parseSubStmt(): SubStmtNode { ... }
```

If the exact section number is not known yet, mark it explicitly:

```ts
// MS-VBAL verification required before widening this rule.
```

No speculative parser broadening should be accepted without a matching fixture and spec note.

### 2. Proper capitalization of VBA keywords is required

VBA is case-insensitive, but XLIDE must present and optionally normalize keywords using canonical VBA/VBE-style capitalization.

Examples:

```vba
Option Explicit
Private Sub Example()
    Dim value As Long
    If value > 0 Then
        Debug.Print value
    End If
End Sub
```

The language service must treat these as equivalent for parsing:

```vba
option explicit
OPTION EXPLICIT
OpTiOn ExPlIcIt
```

But completions, code actions, formatting, generated snippets, and auto-fixes must emit:

```vba
Option Explicit
```

Keyword capitalization must apply only to real keyword tokens. Do not alter text inside comments, strings, date literals, external declarations, or user-defined identifiers unless a separate identifier-case-sync feature is intentionally implemented.

### 3. Low-noise diagnostics only

Do not ship noisy linting. Prefer fewer diagnostics with high certainty.

A diagnostic is acceptable only when at least one of the following is true:

- It is directly supported by `MS-VBAL.pdf`.
- It is verified by a focused fixture against the real VBE/VBA runtime.
- It is clearly labeled as an optional style inspection and disabled by default.

### 4. Determinism over cleverness

The language service must be deterministic. The same project text must produce the same tokens, AST, symbol graph, diagnostics, and completions every run.

Do not use fuzzy heuristics where a simple explicit rule can work.

### 5. Separate VBA language from host object models

The core VBA language service must not confuse these layers:

- VBA language syntax and semantics: verify against `MS-VBAL.pdf`.
- Office/Excel/Word/Access/PowerPoint object models: verify against Microsoft Office VBA object model references or generated COM type-library metadata.
- XLIDE workbook IO behavior: verify against XLIDE tests and workbook round-trips.

Host objects are not the VBA language itself.

---

## Source-of-Truth Hierarchy

Use this hierarchy whenever sources conflict:

1. `MS-VBAL.pdf` for core language syntax and behavior.
2. Real VBE/VBA behavior tests for implementation quirks not obvious from the spec.
3. Microsoft Office VBA language reference for explanatory examples.
4. Microsoft Office object model docs or generated type-library metadata for host APIs.
5. Existing open-source projects as design references only.
6. Internet snippets, Stack Overflow examples, and LLM memory only as leads, never as authority.

---

## Target Capability

The target is not to build a full compiler immediately.

The target is a useful VS Code language service for workbook-backed VBA projects.

Minimum impressive feature set:

- Active syntax diagnostics.
- Project-wide symbol index.
- Go to definition for local procedures and module-level declarations.
- Document symbols and workspace symbols.
- Completion for local variables, parameters, procedures, modules, classes, enums, constants, and keywords.
- Canonical keyword completion and keyword capitalization fixes.
- Hover for known declarations.
- Signature help for procedures/functions.
- Basic member completion for typed variables.
- Static metadata-backed completion for Excel/VBA built-ins.

---

## Architectural Recommendation

Start with direct VS Code providers before building a full Language Server Protocol process.

Recommended sequence:

1. Build the analyzer as a pure TypeScript library.
2. Wire it into VS Code using direct `vscode.languages.*` providers.
3. Stabilize tokens, AST, symbol graph, diagnostics, and completions.
4. Add incremental parsing/indexing.
5. Only then wrap it as an LSP if cross-editor reuse or process isolation becomes necessary.

Suggested package layout:

```text
src/
  extension.ts
  analyzer/
    index.ts
    lexer/
      tokenize.ts
      tokenKinds.ts
      keywordTable.ts
      trivia.ts
    parser/
      parseModule.ts
      nodes.ts
      parserState.ts
      recovery.ts
    symbols/
      projectIndex.ts
      declarationIndex.ts
      scopeResolver.ts
      referenceResolver.ts
    diagnostics/
      diagnosticRules.ts
      syntaxDiagnostics.ts
      semanticDiagnostics.ts
    completions/
      completionProvider.ts
      keywordCompletions.ts
      symbolCompletions.ts
      memberCompletions.ts
    formatting/
      keywordCasing.ts
      safeFormat.ts
    metadata/
      vbaRuntime.json
      excelObjectModel.json
    spec/
      specMap.ts
      verificationStatus.ts
test/
  fixtures/
    lexer/
    parser/
    diagnostics/
    completions/
    casing/
docs/
  spec/
    MS-VBAL.pdf
    MS-VBAL.version.md
    MS-VBAL.verification-map.md
```

---

## Roadmap Phases

## Phase 0: Spec Acquisition and Verification Map

### Goal

Establish `MS-VBAL.pdf` as the canonical verification source before implementing grammar behavior.

### Tasks

- Download the current official `MS-VBAL.pdf` from the Microsoft Learn published version page.
- Save it at `docs/spec/MS-VBAL.pdf`.
- Create `docs/spec/MS-VBAL.version.md` containing:
  - Download date.
  - Microsoft publication date.
  - Protocol revision.
  - PDF filename or source URL.
- Create `docs/spec/MS-VBAL.verification-map.md`.
- Add a verification map table:

```md
| Feature | Implementation File | Fixture | MS-VBAL Section | Status |
|---|---|---|---|---|
| Line comments | src/analyzer/lexer/tokenize.ts | test/fixtures/lexer/comments.bas | TBD | Pending |
| String literals | src/analyzer/lexer/tokenize.ts | test/fixtures/lexer/strings.bas | TBD | Pending |
```

### Acceptance Criteria

- No parser feature is considered complete unless it has a corresponding verification-map row.
- Any unknown grammar behavior is marked `Pending`, not guessed.

---

## Phase 1: Lexer / Tokenizer

### Goal

Create a loss-aware tokenizer for VBA modules.

The lexer must preserve enough trivia to support diagnostics, formatting, keyword casing, source spans, and safe code actions.

### Must Recognize

- Keywords, case-insensitively.
- Identifiers.
- Bracketed identifiers, if verified by `MS-VBAL.pdf`.
- Numeric literals.
- String literals.
- Date literals, if verified by `MS-VBAL.pdf`.
- Comments using apostrophe.
- `Rem` comments, with exact behavior verified against `MS-VBAL.pdf`.
- Line continuations using underscore.
- Statement separators using colon.
- Operators.
- Punctuation.
- Newlines.
- Whitespace/trivia.
- Compiler directives, including `#If`, `#Else`, `#ElseIf`, `#End If`, and `#Const`, after spec verification.
- Attribute lines, such as `Attribute VB_Name = "Module1"`.

### Important Edge Cases

Verify these against `MS-VBAL.pdf` before finalizing behavior:

- Colon-separated statements.
- Apostrophe comments after code.
- `Rem` used as a comment versus `Rem` used near identifiers.
- Line continuation before comments.
- Line continuation inside argument lists.
- Multi-line procedure declarations.
- String literal escaping using doubled quotes.
- Date literal syntax.
- Identifier characters and reserved words.
- Labels and line numbers.

### Output Shape

Token objects should include:

```ts
interface VbaToken {
  kind: TokenKind;
  rawText: string;
  canonicalText?: string;
  start: number;
  end: number;
  line: number;
  character: number;
  leadingTrivia?: Trivia[];
  trailingTrivia?: Trivia[];
}
```

### Acceptance Criteria

- Tokenization is stable and round-trippable.
- Keyword tokens include canonical capitalization.
- Comments and strings are never keyword-normalized.
- All lexer fixtures include expected tokens.
- Lexer behavior has spec-map entries.

---

## Phase 2: Canonical Keyword Table

### Goal

Create a complete, verified keyword table with canonical capitalization.

### Rules

- All keyword matching must be case-insensitive.
- All emitted keyword text must use canonical casing.
- Completion labels must use canonical casing.
- Snippets must use canonical casing.
- Code actions must use canonical casing.
- The formatter must only alter tokens known to be keywords.

### Seed Canonical Keyword Table

This seed table is not a substitute for spec verification. Complete and correct it against `MS-VBAL.pdf` before marking Phase 2 complete.

```ts
export const VBA_KEYWORDS: Record<string, string> = {
  "addressof": "AddressOf",
  "alias": "Alias",
  "and": "And",
  "any": "Any",
  "as": "As",
  "base": "Base",
  "binary": "Binary",
  "boolean": "Boolean",
  "byref": "ByRef",
  "byte": "Byte",
  "byval": "ByVal",
  "call": "Call",
  "case": "Case",
  "cbool": "CBool",
  "cbyte": "CByte",
  "ccur": "CCur",
  "cdate": "CDate",
  "cdbl": "CDbl",
  "cdec": "CDec",
  "cint": "CInt",
  "clng": "CLng",
  "clnglng": "CLngLng",
  "clngptr": "CLngPtr",
  "compare": "Compare",
  "const": "Const",
  "csng": "CSng",
  "cstr": "CStr",
  "currency": "Currency",
  "cvar": "CVar",
  "cvdate": "CVDate",
  "decimal": "Decimal",
  "declare": "Declare",
  "defbool": "DefBool",
  "defbyte": "DefByte",
  "defcur": "DefCur",
  "defdate": "DefDate",
  "defdbl": "DefDbl",
  "defdec": "DefDec",
  "defint": "DefInt",
  "deflng": "DefLng",
  "deflnglng": "DefLngLng",
  "deflngptr": "DefLngPtr",
  "defobj": "DefObj",
  "defsng": "DefSng",
  "defstr": "DefStr",
  "defvar": "DefVar",
  "dim": "Dim",
  "do": "Do",
  "double": "Double",
  "each": "Each",
  "else": "Else",
  "elseif": "ElseIf",
  "empty": "Empty",
  "end": "End",
  "enum": "Enum",
  "eqv": "Eqv",
  "erase": "Erase",
  "error": "Error",
  "event": "Event",
  "explicit": "Explicit",
  "false": "False",
  "for": "For",
  "friend": "Friend",
  "function": "Function",
  "get": "Get",
  "global": "Global",
  "gosub": "GoSub",
  "goto": "GoTo",
  "if": "If",
  "imp": "Imp",
  "implements": "Implements",
  "in": "In",
  "input": "Input",
  "integer": "Integer",
  "is": "Is",
  "let": "Let",
  "lib": "Lib",
  "like": "Like",
  "lock": "Lock",
  "long": "Long",
  "longlong": "LongLong",
  "longptr": "LongPtr",
  "loop": "Loop",
  "lset": "LSet",
  "me": "Me",
  "mod": "Mod",
  "module": "Module",
  "new": "New",
  "next": "Next",
  "not": "Not",
  "nothing": "Nothing",
  "null": "Null",
  "object": "Object",
  "on": "On",
  "open": "Open",
  "option": "Option",
  "optional": "Optional",
  "or": "Or",
  "output": "Output",
  "paramarray": "ParamArray",
  "preserve": "Preserve",
  "print": "Print",
  "private": "Private",
  "property": "Property",
  "public": "Public",
  "put": "Put",
  "raiseevent": "RaiseEvent",
  "random": "Random",
  "read": "Read",
  "redim": "ReDim",
  "rem": "Rem",
  "resume": "Resume",
  "return": "Return",
  "rset": "RSet",
  "select": "Select",
  "set": "Set",
  "single": "Single",
  "static": "Static",
  "step": "Step",
  "stop": "Stop",
  "string": "String",
  "sub": "Sub",
  "text": "Text",
  "then": "Then",
  "to": "To",
  "true": "True",
  "type": "Type",
  "typeof": "TypeOf",
  "until": "Until",
  "variant": "Variant",
  "wend": "Wend",
  "while": "While",
  "with": "With",
  "withevents": "WithEvents",
  "write": "Write",
  "xor": "Xor"
};
```

### Special Casing Requirements

These compound forms must be emitted correctly when generated by snippets, fixes, or formatter logic:

```vba
Option Explicit
Option Base 1
Option Compare Binary
Option Compare Text
Private Sub
Public Sub
Private Function
Public Function
Friend Function
Property Get
Property Let
Property Set
End Sub
End Function
End Property
End If
End Select
End With
End Type
End Enum
Do While
Do Until
Loop While
Loop Until
For Each
On Error GoTo
On Error Resume Next
#If
#ElseIf
#Else
#End If
#Const
```

### Acceptance Criteria

- There is a test fixture where all-lowercase VBA is normalized to canonical keyword casing.
- There is a fixture proving strings and comments are untouched.
- The keyword table is verified against `MS-VBAL.pdf` before being marked complete.

---

## Phase 3: Parser and AST

### Goal

Build an error-tolerant parser that understands the top-level structure of VBA modules and enough statement/expression structure to power diagnostics and IntelliSense.

### First Parser Scope

Parse these constructs first:

- Module attributes.
- `Option Explicit`.
- `Option Base`.
- `Option Compare`.
- `Declare` statements.
- `Const` declarations.
- `Dim` declarations.
- `Private`, `Public`, `Friend`, `Global`, and `Static` declarations after spec verification.
- `Type ... End Type`.
- `Enum ... End Enum`.
- `Sub ... End Sub`.
- `Function ... End Function`.
- `Property Get ... End Property`.
- `Property Let ... End Property`.
- `Property Set ... End Property`.
- Procedure parameters.
- Basic block statements:
  - `If ... Then ... Else ... End If`
  - `Select Case ... End Select`
  - `For ... Next`
  - `For Each ... Next`
  - `Do ... Loop`
  - `While ... Wend`
  - `With ... End With`

### Expression Parsing

Start with enough expression parsing to support:

- Function/procedure calls.
- Member access via `.`.
- Unary and binary operators.
- Parenthesized expressions.
- Literals.
- Identifier references.
- Named arguments, after spec verification.

Do not attempt to perfectly execute or type-evaluate VBA expressions in the first iteration.

### Error Recovery

The parser must survive broken code.

VS Code users edit incomplete code constantly. Diagnostics and completions must still work inside partially written procedures.

Recovery rules:

- Recover at newline boundaries.
- Recover at colon statement separators.
- Recover at known block terminators.
- Recover at procedure terminators.
- Preserve malformed nodes with spans.

### Acceptance Criteria

- Parser never crashes on malformed input.
- AST includes source spans for every node.
- Procedure declarations and parameters are reliably extracted even when procedure bodies contain errors.
- Block mismatch diagnostics can be produced from AST/recovery data.
- Parser fixtures are linked to `MS-VBAL.pdf` verification rows.

---

## Phase 4: Project-Wide Symbol Graph

### Goal

Build a workbook/project-aware symbol index.

XLIDE’s advantage is that it is not just editing isolated `.bas` files. It can see the workbook-backed VBA project.

### Symbol Kinds

Index:

- Project.
- Module.
- Standard module.
- Class module.
- Document module.
- UserForm module.
- Procedure.
- Function.
- Property.
- Parameter.
- Local variable.
- Module variable.
- Constant.
- Enum.
- Enum member.
- User-defined type.
- UDT field.
- Event.
- Declare statement.

### Scope Model

Implement at least:

- Project scope.
- Module scope.
- Procedure scope.
- Block/local scope only where verified and useful.
- Class/member scope.
- Enum scope.
- UDT scope.

### Name Resolution

Implement conservative name resolution first:

1. Local variables and parameters.
2. Procedure-level declarations.
3. Module-level declarations.
4. Current module procedures/properties/functions.
5. Public declarations in other modules.
6. Class members where type is known.
7. Built-in VBA runtime symbols.
8. Host object model symbols.

Where VBA name resolution has nuanced rules, verify against `MS-VBAL.pdf` and/or VBE behavior tests before expanding.

### Acceptance Criteria

- Document symbols work for all module types.
- Workspace symbols work across the loaded workbook/project.
- Go to definition works for local procedures and module declarations.
- Duplicate declaration diagnostics are possible.
- Symbol graph is deterministic.

---

## Phase 5: Active Diagnostics

### Goal

Ship useful, high-confidence active linting.

### First Diagnostics

Enable these first:

- Unclosed string literal.
- Invalid line continuation.
- Missing `End Sub`.
- Missing `End Function`.
- Missing `End Property`.
- Missing `End If`.
- Missing `End Select`.
- Missing `End With`.
- Missing `End Type`.
- Missing `End Enum`.
- Unexpected block terminator.
- Duplicate procedure name in same module.
- Duplicate local variable in same procedure, if spec-confirmed.
- `Option Explicit` missing, configurable severity.
- Variable used but not declared when `Option Explicit` is active and confidence is high.
- Unknown procedure call when confidence is high.
- Assignment to `Const` when confidence is high.

### Disabled by Default Initially

Keep these off until the analyzer is mature:

- Unused variable.
- Implicit Variant warning.
- Procedure too long.
- Naming conventions.
- Hungarian notation checks.
- Complexity warnings.
- Style formatting warnings.

### Diagnostic Metadata

Every diagnostic rule should include:

```ts
interface DiagnosticRuleMetadata {
  code: string;
  title: string;
  defaultSeverity: "error" | "warning" | "information" | "hint";
  source: "XLIDE";
  specReference?: string;
  requiresWholeProject?: boolean;
  confidence: "high" | "medium" | "low";
}
```

Do not ship low-confidence diagnostics by default.

### Acceptance Criteria

- Diagnostics update on document change.
- Diagnostics clear after correction.
- Diagnostics do not require saving the workbook.
- Diagnostics work on virtual XLIDE module documents.
- Every active diagnostic has tests.

---

## Phase 6: IntelliSense and Completions

### Goal

Make the editor feel alive before perfecting all language semantics.

### Completion Sources

Provide completions from:

- Canonical keyword table.
- Current procedure locals.
- Current procedure parameters.
- Module-level declarations.
- Current module procedures.
- Public project procedures.
- Classes.
- Enums and enum members.
- UDTs and fields where known.
- Built-in VBA runtime functions.
- Host object model metadata.
- XLIDE-provided workbook context, where safe.

### Trigger Contexts

Implement completions for:

- Empty line / statement start.
- After access modifiers: `Private`, `Public`, `Friend`, `Static`.
- After `As` for known types.
- After `New` for creatable classes.
- After `.` for known member access.
- Inside call argument lists for signature help.
- After `Option` for `Explicit`, `Base`, `Compare`.
- After `On Error` for `GoTo` and `Resume Next` patterns.

### Completion Output Rules

- Keywords must use canonical capitalization.
- Snippets must use canonical capitalization.
- Insert text must not randomly alter nearby identifiers.
- Sort symbols before broad snippets when context is specific.
- Avoid suggesting invalid keywords in narrow contexts where parser state is known.

### Acceptance Criteria

- `dim x as ` suggests canonical VBA types and known classes.
- `private ` suggests `Sub`, `Function`, `Property`, declarations where valid.
- `Option ` suggests `Explicit`, `Base`, `Compare`.
- `End ` suggests valid block endings based on context.
- `ws.` suggests members when `ws As Worksheet` can be resolved through metadata.
- Completion tests cover casing.

---

## Phase 7: Hover, Definition, References, and Signature Help

### Goal

Add high-value navigation and explanation features.

### Hover

Show:

- Symbol kind.
- Declaration signature.
- Type, if known.
- Module/class origin.
- Visibility.
- Optional source note for built-ins/host objects.

Example:

```text
Function GetCustomer(id As Long) As Customer
Declared in Module: CustomerApi
Visibility: Public
```

### Go to Definition

Support first:

- Local variables.
- Parameters.
- Module variables.
- Procedures in the same module.
- Procedures in other modules.
- Enums and enum members.
- UDTs and fields.
- Class members where known.

### References

Support conservatively:

- Exact symbol references in same procedure.
- Same-module references.
- Cross-module references once resolver is stable.

### Signature Help

Support:

- Project procedures/functions.
- Built-in VBA runtime functions from metadata.
- Host object model methods from metadata.

### Acceptance Criteria

- Go to definition and hover work without saving.
- Results are based on current editor contents plus loaded project index.
- Ambiguous symbols are handled explicitly, not guessed silently.

---

## Phase 8: Keyword Capitalization Code Action / Formatter

### Goal

Provide safe keyword normalization without becoming an intrusive formatter.

### Features

- Command: `XLIDE: Normalize VBA Keyword Capitalization`.
- Code action: `Normalize keyword capitalization in document`.
- Optional on-save setting:

```json
{
  "xlide.vba.normalizeKeywordCasingOnSave": false
}
```

Default must be `false` initially.

### Safety Rules

The capitalization pass may alter only tokens where:

- Token kind is `Keyword`.
- Token span is not inside comment/string/date literal.
- Token was produced by the lexer as a keyword token.
- Replacement is exactly the canonical keyword spelling.

Do not modify:

- Identifiers.
- Procedure names.
- Variable names.
- Module names.
- Class names.
- String contents.
- Comments.
- Attribute values.
- External declaration aliases.

### Example

Input:

```vba
option explicit
private sub test()
    dim x as long
    if x = 0 then debug.print "if then else"
end sub
```

Output:

```vba
Option Explicit
Private Sub test()
    Dim x As Long
    If x = 0 Then Debug.Print "if then else"
End Sub
```

### Acceptance Criteria

- Safe on ugly mixed-case code.
- Does not change comments or strings.
- Does not change identifiers.
- Has snapshot tests.
- Uses spec-verified keyword table.

---

## Phase 9: Built-In VBA Runtime Metadata

### Goal

Provide completions, hover, and signature help for built-in VBA functions and types.

### Metadata File

Create:

```text
src/analyzer/metadata/vbaRuntime.json
```

Shape:

```json
{
  "functions": {
    "MsgBox": {
      "name": "MsgBox",
      "signature": "MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context]) As VbMsgBoxResult",
      "returns": "VbMsgBoxResult",
      "source": "verified"
    }
  },
  "types": {
    "Long": { "name": "Long", "kind": "intrinsic" },
    "String": { "name": "String", "kind": "intrinsic" }
  },
  "constants": {},
  "enums": {}
}
```

### Verification

Core language built-ins must be verified against `MS-VBAL.pdf` or Microsoft VBA documentation.

Do not invent signatures from memory.

### Acceptance Criteria

- Built-ins appear in completions.
- Signature help works for at least a small verified set.
- Metadata entries include source status.

---

## Phase 10: Host Object Model Metadata

### Goal

Provide useful Excel/Office IntelliSense without pretending host APIs are part of core VBA.

### Strategy

Use static JSON metadata first.

Later, optionally generate metadata from COM type libraries on Windows.

Suggested files:

```text
src/analyzer/metadata/excelObjectModel.json
src/analyzer/metadata/officeObjectModel.json
src/analyzer/metadata/wordObjectModel.json
src/analyzer/metadata/accessObjectModel.json
src/analyzer/metadata/powerPointObjectModel.json
```

### Metadata Shape

```json
{
  "types": {
    "Excel.Worksheet": {
      "displayName": "Worksheet",
      "members": {
        "Range": {
          "kind": "property",
          "returns": "Excel.Range"
        },
        "Activate": {
          "kind": "method",
          "signature": "Activate()"
        }
      }
    }
  },
  "aliases": {
    "Worksheet": "Excel.Worksheet",
    "Range": "Excel.Range",
    "Workbook": "Excel.Workbook"
  }
}
```

### Member Completion Example

Given:

```vba
Dim ws As Worksheet
ws.
```

Resolve:

```text
ws -> Worksheet -> Excel.Worksheet -> members
```

Then suggest verified members.

### Acceptance Criteria

- `Application.` provides useful completions.
- `Workbook`, `Worksheet`, and `Range` provide useful completions.
- Metadata is versioned.
- Metadata source is documented.
- Host metadata never overrides core language rules.

---

## Phase 11: VS Code Integration

### Goal

Expose the analyzer through VS Code APIs.

### Providers

Implement:

- `DocumentSymbolProvider`.
- `WorkspaceSymbolProvider`.
- `DefinitionProvider`.
- `ReferenceProvider`.
- `CompletionItemProvider`.
- `HoverProvider`.
- `SignatureHelpProvider`.
- `CodeActionProvider`.
- Diagnostics through `DiagnosticCollection`.
- Optional `DocumentFormattingEditProvider` only for very safe formatting.

### Update Model

- Re-analyze current document on text change with debounce.
- Rebuild project index when XLIDE virtual modules are loaded, saved, renamed, added, or deleted.
- Use incremental invalidation by module when possible.
- Never block the UI thread with full project analysis on every keystroke.

### Acceptance Criteria

- Active diagnostics update while typing.
- Project symbol index updates after module save/load.
- IntelliSense works inside virtual workbook-backed module files.
- No corruption risk to workbook save-back.

---

## Phase 12: Testing Strategy

### Required Test Types

- Lexer fixtures.
- Parser fixtures.
- AST snapshots.
- Diagnostic fixtures.
- Completion fixtures.
- Keyword casing fixtures.
- Project index fixtures.
- Workbook-backed integration fixtures.
- VBE behavior comparison fixtures where needed.

### Fixture Format

Recommended:

```text
test/fixtures/parser/simple-sub/input.bas
test/fixtures/parser/simple-sub/expected.ast.json
test/fixtures/parser/simple-sub/spec.md
```

Each `spec.md` should say:

```md
# Fixture: simple-sub

Verified against: MS-VBAL.pdf
Section: TBD
Notes: Basic Sub block parse.
```

### Golden Rule

No grammar feature is “done” without a test and a spec-map row.

---

## Phase 13: Settings

Add settings slowly.

Suggested initial settings:

```json
{
  "xlide.vba.diagnostics.enabled": true,
  "xlide.vba.diagnostics.optionExplicit": "warning",
  "xlide.vba.completions.keywords": true,
  "xlide.vba.completions.projectSymbols": true,
  "xlide.vba.completions.hostObjectModel": true,
  "xlide.vba.normalizeKeywordCasingOnSave": false,
  "xlide.vba.specStrictMode": true
}
```

`specStrictMode` means the analyzer avoids unverified broad grammar interpretations.

---

## Phase 14: Do-Not-Do List

Do not:

- Build a full compiler first.
- Treat VB.NET as equivalent to VBA.
- Treat Excel object model behavior as VBA language behavior.
- Add noisy style warnings early.
- Auto-format entire files aggressively.
- Normalize identifiers without a separate user setting.
- Change comments or strings during keyword capitalization.
- Trust LLM-generated grammar without `MS-VBAL.pdf` verification.
- Guess function signatures from memory.
- Require COM for the base language service.
- Block VS Code typing with whole-project analysis on each keystroke.

---

## Minimal Vertical Slice

Implement this first to prove the architecture:

### Input

```vba
option explicit

private sub Example()
    dim ws as Worksheet
    ws.
end sub
```

### Expected Behavior

- Diagnostics: none.
- Code action: normalize keyword casing.
- Completion after `ws.` suggests `Worksheet` members if Excel metadata is loaded.
- Document symbols show `Example`.
- Hover over `ws` shows local variable, type `Worksheet`.
- Keyword normalization produces:

```vba
Option Explicit

Private Sub Example()
    Dim ws As Worksheet
    ws.
End Sub
```

This vertical slice proves lexer, keyword table, parser, symbols, completions, diagnostics, casing, and VS Code wiring.

---

## Implementation Prompt for Another LLM

Use this prompt when handing work to another LLM:

```text
You are implementing the VBA language-service layer for the XLIDE VS Code extension.

Hard requirements:
1. Verify all VBA language grammar, keyword, operator, literal, declaration, scoping, and syntax behavior against the official Microsoft [MS-VBAL]: VBA Language Specification PDF, stored locally as docs/spec/MS-VBAL.pdf.
2. Do not treat VB.NET, Excel object model examples, Rubberduck behavior, Stack Overflow examples, or LLM memory as canonical for core VBA syntax.
3. Proper capitalization of VBA keywords is required. VBA is case-insensitive, but completions, snippets, code actions, and formatting must emit canonical casing such as Option Explicit, Private Sub, Dim, As, If, Then, End If, End Sub.
4. Keyword capitalization must not alter comments, strings, date literals, identifiers, procedure names, variable names, module names, aliases, or attribute values.
5. Keep diagnostics conservative and high-confidence.
6. Prefer deterministic parsing and symbol indexing over fuzzy heuristics.
7. Maintain docs/spec/MS-VBAL.verification-map.md so each implemented language feature maps to a spec section and test fixture.

Build in this order:
1. Lexer with canonical keyword table and trivia preservation.
2. Parser for modules, options, declarations, procedures, properties, types, enums, and major block statements.
3. Project-wide symbol graph.
4. High-confidence diagnostics.
5. Keyword completions and symbol completions.
6. Safe keyword capitalization code action.
7. Hover, go-to-definition, references, and signature help.
8. Built-in VBA metadata.
9. Host object model metadata for Excel, clearly separated from core VBA language rules.
10. VS Code providers and tests.

Do not mark a feature complete unless it has tests and a corresponding MS-VBAL verification note.
```

---

## Definition of Done

The VBA language-service layer is ready for a first public preview when:

- The current `MS-VBAL.pdf` has been downloaded and version-recorded.
- The verification map exists and covers all implemented grammar features.
- Keyword table is spec-verified.
- Keyword capitalization works safely.
- Lexer and parser are stable on malformed code.
- Active diagnostics work without saving.
- Completions include keywords, local symbols, project symbols, and at least basic Excel metadata.
- Document symbols and go-to-definition work for common cases.
- Tests cover all shipped features.
- No automatic formatting feature can corrupt user code.
- All behavior that is not spec-verified is explicitly labeled experimental or disabled by default.
