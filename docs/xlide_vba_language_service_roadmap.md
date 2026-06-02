# XLIDE VBA Language Service Roadmap

## Purpose

Build the next layer of XLIDE: active analysis, IntelliSense, symbol navigation, and safe formatting for VBA inside VS Code.

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

Do not ship noisy analysis. Prefer fewer diagnostics with high certainty.

A diagnostic is acceptable only when at least one of the following is true:

- It is directly supported by `MS-VBAL.pdf`.
- It is verified by a focused fixture against the real VBE/VBA runtime.
- It is clearly labeled as an optional style inspection and disabled by default.

### 4. Determinism over cleverness

The language service must be deterministic. The same project text must produce the same tokens, AST, symbol graph, diagnostics, and completions every run.

Do not use fuzzy heuristics where a simple explicit rule can work.

### 5. Unified business rules across surfaces

XLIDE must not have separate interpretations of the same VBA rule for different
UI surfaces. Unless a behavior is explicitly documented as a corner case, the
same semantic rule must drive completions, inserted text, hover, signature help,
diagnostics, go-to-definition, find references, rename, tree actions, semantic
coloring, formatter logic, snippets, and code actions.

When adding or changing a rule, implement it in a shared analyzer/provider helper
where practical and add regression coverage for every affected surface, rather
than duplicating similar logic in separate UI paths.

Prefer the broadest deterministic implementation that can remain low-noise. A
targeted fix is acceptable only when the corresponding broad rule would require
guessing, incomplete metadata, or user-visible false positives; when the broader
rule becomes provable, consolidate the narrow path into it.

### 6. Separate VBA language from host object models

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

> Status: DONE. Spec stored at `docs/[MS-VBAL].pdf` (v20250520); version recorded
> in `docs/spec/MS-VBAL.version.md`; verification map at
> `docs/spec/MS-VBAL.verification-map.md`.

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

> Status: DONE (core). Implemented in `src/analyzer/lexer/{tokenize,trivia,
> tokenKinds}.ts`, fixtures in `tests/vbaLexer.test.ts`. Round-trippable and
> spec-cited. Date-token inner grammar, non-Latin codepage ranges, and directive
> block parsing remain Partial (see verification map).

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

> Status: DONE. Implemented in `src/analyzer/lexer/keywordTable.ts`, fixtures in
> `tests/vbaKeywordTable.test.ts`. The seed table below was completed and
> corrected against MS-VBAL 3.3.5.2 (added missing reserved-names such as
> `CVErr`, `DoEvents`, `Abs`, `Fix`, `LenB`, special-forms, etc.) and split into
> spec reserved identifiers vs VBE-convention contextual keywords.

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

> Status: DONE (structural). Implemented in `src/analyzer/parser/{nodes,
> parserState,parseModule}.ts`, fixtures in `tests/vbaParser.test.ts`, verification
> rows in `docs/spec/MS-VBAL.verification-map.md`. Builds a
> `ModuleNode` AST (attributes, options, conditional-compilation directives,
> declarations, `Type`/`Enum`, procedures + parameters, nested block statements)
> with absolute source spans;
> never throws on malformed input; emits block-mismatch diagnostics. Deferred:
> full expression AST (calls/member-access/operators, section 5.6) and `If`
> branch modeling — captured as raw `Statement` nodes for now and tracked as
> Pending in the verification map.

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

> Status: DONE (structural). Implemented in `src/analyzer/symbols/{symbolModel,
> buildModuleSymbols,projectIndex}.ts`: a pure AST -> symbol projection plus a
> `ProjectIndex` that answers document symbols, workspace symbols, conservative
> go-to-definition name resolution (locals/params -> same-module -> exported
> cross-module), and duplicate-procedure detection. Cross-module visibility
> follows MS-VBAL (default-Public procedures exported; Private/Dim/Friend module
> members stay private; enum members inherit the visibility of their containing
> `Enum`). Covered by `tests/vbaSymbolGraph.test.ts` (59 tests).
> Verification-map rows added. The AST index now drives the live Go to
> Definition, Find All References, and Rename providers in
> `src/vbaLanguageProviders.ts` via `resolveDefinition`,
> `resolveQualifiedDefinition`, and `referenceScope` (scope-restricted
> occurrence search). Remaining: wiring the AST index into the live VS Code
> Document/Workspace symbol providers (still served by the interim regex index
> in `src/vbaSymbolIndex.ts`), and richer block/UDT/class member scope
> resolution.

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

> Second-pass note (per addendum): in addition to the pure MS-VBAL language
> symbols above, maintain a separate **host-context symbol layer** for globals
> the host injects rather than the user declaring them — `ThisWorkbook`,
> `Application`, `ActiveWorkbook`, `ActiveSheet`, worksheet code names
> (`Sheet1`), and `Me` (resolved by module kind). These resolve to host
> object-model types (Phase 10), must never override core-language resolution,
> and worksheet code names must come from the actual workbook project structure,
> not the visible sheet tab name.

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

> Status: IN PROGRESS (core shipped). A pure, vscode-free analyzer engine
> `src/analyzer/diagnostics/{ruleMetadata,analyzeModule}.ts` (barrel-exported
> as `analyzeModule`, `DIAGNOSTIC_RULES`) computes high-confidence semantic
> diagnostics directly from editor text - no save, no Python round-trip. It is
> merged with the existing structural block-balance analyzer (`analyzeVbaStructure`,
> which already covers every "Missing End .../Unexpected block terminator"
> case) inside `registerVbaDiagnostics` in `src/vbaLanguageProviders.ts`, runs
> on open and debounced (300 ms) on every edit, and works on virtual
> `xlide-vba` module documents. Settings: `xlide.diagnostics.enabled`
> (default true) and `xlide.diagnostics.optionExplicit`
> (off/hint/information/warning/error, default warning); both re-run open
> documents on change. Covered by `tests/vbaDiagnostics.test.ts`.
>
> Shipped semantic rules (all high-confidence, spec-referenced in
> `ruleMetadata.ts` and `docs/spec/MS-VBAL.verification-map.md`):
> - `unterminated-string` - odd-quote-count detection, handles `""` escapes.
> - `duplicate-procedure` - allows Property Get/Let/Set to share a name.
> - `duplicate-declaration` - param/local collisions, flat procedure scope.
> - `duplicate-module-variable` - module-level redeclaration.
> - `const-assignment` - bare `name =` to a Const (excludes `.member`,
>   `index(...)`, `Set`, comparisons).
> - `option-explicit-missing` - configurable; silent on empty/attribute-only
>   modules.
> - `unknown-call` ("Sub or Function not defined") - a call statement whose
>   callee is a bare (non-member) identifier and resolves to nothing: no project
>   procedure visible as a bare call from the current module (the provider passes
>   `ProjectIndex.visibleProcedureNames(moduleName)` as `knownProcedures`), no runtime
>   function/statement, no host global or `Application` member, and no in-scope
>   declaration. `bareCallStatementTarget` accepts the three unambiguous forms - a
>   lone identifier, a parenless call with arguments (`MsgBox ""`), and an
>   explicit `Call name`. Assignments (any top-level `=`), member calls, line
>   labels, and the implicit-host-member form `Cells(1, 1)` / `Range("A1")` are
>   deliberately not touched. The rule is skipped entirely when
>   `knownProcedures` is absent so a module is never analysed in isolation.
> - `invalid-proc-header` ("Invalid procedure declaration") - a malformed
>   `Sub`/`Function`/`Property` header where a token other than `(` (or `As` for
>   a `Function`/`Property Get`) follows the procedure name, e.g. `Sub My Sub`
>   or `Function Calc Total()`. Valid parameterless subs, parameterised subs,
>   return-typed functions, and `Property Get ... As` headers stay clean.
> - `unbalanced-parens` - a `(` left open at a statement boundary or a `)` with
>   no matching `(`, within one logical statement. Token-stream depth scan that
>   resets at each statement boundary (newline / depth-0 `:`); parentheses in
>   strings, comments, date literals and `[bracketed]` names are distinct token
>   kinds so they never miscount. At most one diagnostic per statement.
> - `argument-count` ("Wrong number of arguments") - same-module procedures,
>   unique exported project procedures, module-qualified standard-module
>   procedures, verified runtime functions, and known source-backed/host member
>   calls share one callable-signature path. It handles parenthesized calls,
>   explicit `Call`, and parenless call statements, honoring
>   `Optional`/`ParamArray` and named arguments while skipping ambiguous names.
> - `call-requires-parens` - explicit `Call` statements with arguments must wrap
>   the argument list in parentheses.
> - `invalid-explicit-call-target` - runtime entries can opt out of explicit
>   `Call` through shared metadata. The current oracle-backed case is
>   `DoEvents`: `Call DoEvents` and `Call DoEvents()` are rejected, while bare
>   `DoEvents` and expression `DoEvents()` stay valid.
> - `call-statement-forbids-parens` - standalone zero-argument calls such as
>   `myFunction()`, `DoEvents()`, `Application.Calculate()`, and
>   `ThisWorkbook.CanCheckIn()` are rejected when they use empty parentheses
>   without valid `Call` syntax or expression context; required-argument calls
>   stay on `argument-count`.
> - `With` receiver member binding now uses the shared member-access resolver for
>   leading-dot diagnostics: unknown source-backed class/UDT members, read-only
>   assignments, member assignment type checks, class/host member argument count
>   checks, argument type checks, and standalone empty-parentheses call-statement
>   checks all resolve `.Member` against the active `With` object.
> - `invalid-expression-syntax` - narrow expression syntax coverage for
>   impossible operator sequences (`***`) and statements ending in a binary
>   operator.
>
> Deliberately deferred (would require a fuller expression binder + broader host
> catalogue, and per the project's no-false-positive rule must not ship until
> they can be proven safe): full flow-aware identifier binding, ambiguous
> external-reference behavior, and the broad arbitrary-expression form of
> `unknown-call`. The shipped `undeclared-variable` slice now covers
> project-backed `Option Explicit` write/read positions such as assignment
> targets, RHS/call-argument reads, control-flow headers, member receivers, and
> indexed bases, while skipping type-name, label, named-argument, and unresolved
> external-style call positions.

### Goal

Ship useful, high-confidence active analysis.

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

> Status: IN PROGRESS. Member access after `.` (host-context member completion,
> see addendum), type completion after `As` / `As New` / expression `New`, and bare-identifier
> completion (host globals + worksheet/document code names + in-scope
> declarations) are implemented:
> `src/analyzer/completion/{memberAccess,typeCompletion,identifierCompletion}.ts`
> (pure) wired through `src/vbaMemberCompletion.ts`. Type completion offers VBA
> built-in types, Excel host types, and project-defined types from
> `ProjectIndex.visibleTypeNames()` (object-module names plus visible
> `Type`/`Enum` declarations, with duplicate names preserved as ambiguous).
> Identifier completion offers
> host-injected globals, code names, and the enclosing procedure's
> params/locals plus module-level vars/consts/procs/Declares/enums/types, exported
> standard-module project declarations from
> `ProjectIndex.visibleIdentifierSymbols()` (globals, enums, enum members,
> types, and procedures) with procedure signatures still coming from the shared
> project-signature surface, runtime
> constants, and generated Excel enum constants; it is
> suppressed after `.`, after `As`, and in declaration-name positions. Exported
> project procedure/Declare completions carry full callable signatures, declaring-module
> detail, parameter defaults where known, and inline `'''` documentation for the
> IntelliSense preview. Runtime completion uses the shared explicit-`Call`
> compatibility metadata, so invalid targets such as `DoEvents` are not offered
> after `Call`. Member completion/hover/diagnostics/navigation consume
> `ProjectIndex.projectMemberSurfaces(moduleName)`: object-module members plus
> visible UDT fields. UDT fields are exhaustive, writable, definition-backed
> property-like members, so `Dim p As TPoint : p.` and `With p : .X` use the
> same member resolver as workbook class members. The shared member-access
> resolver also resolves leading-dot chains inside an active `With ... End With`
> block against the `With` receiver; completion, hover, signature help
> (including parenless `.Method arg, ...` call tips), canonical casing,
> call-completion parenthesis insertion, and diagnostics now use that same path.
> Parenthesized calls, parenless calls, and
> explicit `Call` contexts are now classified in
> `src/analyzer/call/callContext.ts` so signature help, callable completion
> insertion, and bare-call diagnostics share one rule set. Resolved
> type names (project, host, and primitive) also get semantic tokens and
> type-position hover via `src/analyzer/semantic/typeSemanticTokens.ts`, covered by
> `tests/vbaSemanticTokens.test.ts`. The completion slices are covered by
> `tests/vba{MemberCompletion,TypeCompletion,IdentifierCompletion}.test.ts`.
> Keyword/snippet completion is now implemented by
> `src/analyzer/completion/keywordCompletion.ts`: additive block snippets at
> statement start (`If`, `With`, `For`, `For Each`, `Do`, `While`,
> `Select Case`, procedure blocks), exclusive grammar completions after
> `Option`, `End`, `On Error`, access modifiers, and `#` conditional
> compilation markers, with innermost-block-aware close suggestions. Remaining:
> optional enter-time auto-block insertion after manually typed block headers.
> Current next slice: branch-aware conditional compilation for `VBA7`,
> `Win64`, `Win32`, `PtrSafe`, `LongPtr`, and related 32/64-bit Office patterns
> on top of the shared `Declare` metadata now in place.

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
- Built-in VBA runtime functions and constants.
- Host object model metadata.
- Generated host enum constants.
- Workbook-defined class/document/UserForm members when the receiver type is
  known.
- External object/member metadata for explicitly declared referenced APIs.
- XLIDE-provided workbook context, where safe.

### Trigger Contexts

Implement completions for:

- Empty line / statement start.
- After access modifiers: `Private`, `Public`, `Friend`, `Static`.
- After `As` for known types.
- After `New` for creatable classes.
- After `.` for known member access.
- After leading `.` inside `With ... End With`, resolving against the active
  `With` receiver.
- Inside call argument lists for signature help.
- After `Option` for `Explicit`, `Base`, `Compare`.
- After `On Error` for `GoTo` and `Resume Next` patterns.
- After `End` for the innermost open block's valid terminator.
- After `#` for conditional compilation directives and, later, verified
  32/64-bit external declaration templates.

### Completion Output Rules

- Keywords must use canonical capitalization.
- Snippets must use canonical capitalization.
- Insert text must not randomly alter nearby identifiers.
- Callable insert text must respect VBA call-statement syntax: standalone
  procedure/method completions insert only the canonical name, while expression
  and explicit `Call` contexts may insert `(...)`; runtime entries that are not
  valid explicit-`Call` targets must be filtered from that target context.
- Multi-line snippets that repeat a logical name must use linked placeholders.
  Example: changing the iterator in a generated `For ... Next` block must also
  change the `Next` variable name.
- VS Code snippet mode must not trap normal navigation. If a user clicks away
  or uses keyboard navigation after accepting a VBA keyword snippet, XLIDE should
  leave snippet mode instead of letting a later `Tab` / `Enter` unexpectedly
  continue an old snippet session.
- Sort symbols before broad snippets when context is specific.
- Avoid suggesting invalid keywords in narrow contexts where parser state is known.

### Conditional Compilation and External Declare Roadmap Slice

Status: IN PROGRESS. The shared metadata foundation is implemented: `Declare`
statements now parse visibility, `PtrSafe`, `Sub`/`Function`, name, `Lib`,
`Alias`, parameters, and return type. That metadata is indexed as a
bare-callable signature and feeds same-module/project completion, hover,
signature help, argument-count diagnostics, and argument-type diagnostics.
Conditional-compilation directives now parse into the AST at module and
procedure scope, and `src/analyzer/conditional/conditionalCompilation.ts`
collects them in source order, indexes `#Const` definitions, evaluates
high-confidence compiler-constant expressions when a caller supplies the target
compiler constants, and reports branch activity as `active` / `inactive` /
`unknown`. `createConditionalActivityTracker` is now wired into the shared
module symbol graph, `ProjectIndex`, and active diagnostics, so completions,
hover/signature surfaces that consume project signatures, call validation,
duplicate-declaration checks, and type diagnostics all skip only branches proven
inactive. A first low-noise Win64 diagnostic is enabled: active `Declare`
statements missing `PtrSafe` are flagged only when supplied compiler constants
prove `Win64`.

This slice covers the VBA `#` syntax used for 32/64-bit Office compatibility and
Win32 API declarations. It belongs partly to parser/diagnostics and partly to
completion, but the user-facing experience should be cohesive.

Authoritative references:

- Microsoft compiler constants:
  <https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/compiler-constants>
- Microsoft conditional compilation overview:
  <https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-conditional-compilation>
- Microsoft `PtrSafe` reference:
  <https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ptrsafe-keyword>

Required behavior:

- [Done] Parse and index conditional-compilation directives: `#Const`, `#If`,
  `#ElseIf`, `#Else`, and `#End If`.
- [Done foundation] Evaluate high-confidence directive expressions for supplied compiler constants:
  `VBA7`, `Win64`, `Win32`, `Mac`, and project `#Const` values. Treat `Win32`
  carefully because Microsoft's compiler-constants documentation says it is
  true in both 32-bit and 64-bit Windows development environments; prefer
  examples that test `Win64` before falling back to `Win32`.
- [Done] Parse full external declaration metadata: visibility, `Declare`, `PtrSafe`,
  `Sub` / `Function`, procedure name, `Lib`, optional `Alias`, parameters, and
  return type.
- [Done first slice] Add low-noise diagnostics for verified cases only:
  missing `PtrSafe` in active `Win64` branches, and duplicate declarations only
  when they are active in the same conditional branch. Remaining compatibility
  work: obviously pointer-sized parameters/returns that should use `LongPtr` or
  `LongLong`, and malformed directive blocks.
- [Done first slice] Make branch-aware analysis suppress false positives from
  mutually exclusive 32-bit and 64-bit declarations across the shared symbol
  graph and active diagnostics.
- Provide snippets for common patterns such as `#If VBA7 Then ... #Else ...
  #End If` and `#If Win64 Then ... #Else ... #End If` with `Declare PtrSafe`
  templates.
- Surface `Declare` hovers and signature help from parsed metadata, including
  `Lib` / `Alias` details when present.
- Record MS-VBAL and Microsoft Learn verification notes before enabling new
  diagnostics by default.

### Acceptance Criteria

- `dim x as ` suggests canonical VBA types and known classes.
- `private ` suggests `Sub`, `Function`, `Property`, declarations where valid.
- `Option ` suggests `Explicit`, `Base`, `Compare`.
- `End ` suggests valid block endings based on context.
- `ws.` suggests members when `ws As Worksheet` can be resolved through metadata.
- Completion tests cover casing.

---

## Phase 7: Hover, Definition, References, and Signature Help

> Status: IN PROGRESS. Hover is DONE - `src/analyzer/hover/resolveHover.ts`
> (pure) describes the identifier under the cursor (host members via the
> exported `resolveReceiverTypeAt`, host globals, worksheet code names, and live
> user declarations from the module symbol graph (procedure/Declare signatures,
> variables/parameters/constants with `As` type, enums/members, types/fields),
> exported standard-module project procedures/Declares visible as bare calls,
> and built-in constants,
> wired through the `HoverProvider` in `src/vbaMemberCompletion.ts` and covered by
> `tests/vbaHover.test.ts`. Signature Help is DONE -
> `src/analyzer/signature/signatureHelp.ts` (pure) returns the active call tip
> for host members (verified `Workbooks.Open`, `Range.Offset`, ...), user
> procedures and external Declares (current-module AST plus exported standard-module project
> `Sub`/`Function`/`Declare` signatures), and runtime built-ins, with explicit-`Call`
> runtime incompatibilities suppressed through the same shared metadata used by
> completion and diagnostics, wired through the
> `SignatureHelpProvider` in `src/vbaMemberCompletion.ts` (triggers `(` `,`
> space) and covered by `tests/vbaSignatureHelp.test.ts`. Go to Definition,
> Find All References, and Rename are DONE - the live providers in
> `src/vbaLanguageProviders.ts` now build an AST `ProjectIndex`
> (`src/analyzer/symbols/projectIndex.ts`) per query: definitions use
> scope-aware `resolveDefinition`/`resolveQualifiedDefinition`; references and
> rename use `referenceScope` to restrict the textual occurrence search to the
> binding scope (local procedure, owning module, or whole project minus
> privately-shadowing modules and locals). Exported enum members use the same
> module-level binding rule here as Option Explicit visibility, while private
> enum members shadow only their own module. Covered by the `referenceScope`,
> `resolveDefinition`, and `resolveQualifiedDefinition` cases in
> `tests/vbaSymbolGraph.test.ts`.

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

> Status: DONE. `src/analyzer/signature/signatureHelp.ts` resolves the active
> call tip from module text (paren and parenless call statements), sourcing the
> signature from verified host-member signatures, user-procedure AST, or runtime
> built-ins, with the active parameter tracked across commas. Runtime entries
> that are not legal explicit-`Call` targets do not surface call tips in that
> invalid context.

Support:

- Project procedures/functions.
- Built-in VBA runtime functions from metadata.
- Built-in VBA and host enum constants from metadata.
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

### Status: IN PROGRESS (hover + completion shipped)

Implemented in `src/analyzer/runtime/vbaRuntime.ts`.

- **Deviation from the JSON suggestion above.** The metadata lives in a typed
  TS module, not `src/analyzer/metadata/vbaRuntime.json`. This matches the
  existing host-model precedent (`src/analyzer/host/excelObjectModel.ts` is a TS
  module, not JSON) and gives compile-time checking of every entry. The
  `VbaRuntimeFunction` interface carries `name`, `signature`, optional
  `returns`, a `kind` of `function | statement`, optional explicit-`Call`
  compatibility metadata, and `source: 'verified'`.
- **Coverage.** ~85 verified intrinsic functions/statements: interaction
  (`MsgBox`, `InputBox`, `Shell`, `CreateObject`...), strings (`Left`, `Mid`,
  `Replace`, `InStr`, `Split`, `Format`...), conversions (`CLng`, `CStr`,
  `CDate`...), math (`Abs`, `Round`, `Sqr`...), date/time (`Now`, `DateAdd`,
  `Year`...), arrays/inspection (`Array`, `UBound`, `IsNumeric`, `TypeName`...),
  and `RGB`/`IIf`/`Choose`/`Switch`.
- **Deliberate omissions.** Names that collide with intrinsic data types or are
  otherwise context-ambiguous (`Date`, `Time`, `String`, `Error`) are excluded
  so type-position hover handles them as types instead of runtime calls.
- **Constants.** `VBA_RUNTIME_CONSTANTS` and generated Excel enum constants are
  resolved case-insensitively for completion, hover, and high-confidence
  `Option Explicit` diagnostics.
- **Hover.** `resolveHover` resolves built-ins *after* user symbols, host
  globals, and code names, so a user declaration of the same name correctly
  shadows the built-in. Returns the signature plus a `VBA runtime function` /
  `VBA runtime statement` / constant detail line.
- **Completion.** `resolveIdentifierCompletions` offers built-ins at bare-
  identifier positions (new `runtime` completion kind, `Function` icon),
  gated by `includeRuntime` (default true), and filters runtime entries that are
  VBE-oracle-verified as invalid explicit-`Call` targets when the cursor is at
  `Call <target>`.
- **Explicit Call compatibility.** `runtimeAllowsExplicitCall` centralizes
  runtime-specific call-statement behavior. `DoEvents` currently opts out based
  on focused VBE oracle cases: `Call DoEvents` and `Call DoEvents()` reject as
  Syntax error, while bare `DoEvents` and expression `DoEvents()` compile.
- **Verification.** Signatures transcribed from
  learn.microsoft.com/office/vba/language and MS-VBAL; never LLM-invented.
- **Signature help.** DONE - the verified runtime signatures (and host-member
  and user-procedure signatures) now drive the parameter-info call tip via
  `src/analyzer/signature/signatureHelp.ts`; see Phase 7 Signature Help.
- Tests: `tests/vbaRuntime.test.ts` + runtime cases in `tests/vbaHover.test.ts`
  + `tests/vbaSignatureHelp.test.ts`.

---

## Phase 10: Host Object Model Metadata

> Second-pass note: the **Addendum: Host-Context Member Completion** (end of this
> document) is the concrete first milestone of this phase. It additionally
> requires a *host-context symbol layer* (see Phase 4) that resolves host-
> injected globals such as `ThisWorkbook`, `Application`, `ActiveSheet`,
> `Sheet1` (worksheet code name), and `Me` to their object-model types before
> member completion can run. Treat the addendum's resolution table and Excel
> member metadata as the deliverables that satisfy this phase's acceptance
> criteria.

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
- `Dim p As Person: p.` suggests public members from the workbook-defined
  `Person` class module.
- External object/member metadata can add deterministic completions for
  referenced APIs that are not present in workbook source.
- Metadata is versioned.
- Metadata source is documented.
- Host metadata never overrides core language rules.
- Downstream developer documentation explains the metadata schema, examples,
  reload behavior, precedence, and troubleshooting before the workflow ships.

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
- `DocumentSemanticTokensProvider`.
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

### Configuration Scoping Contract

Workbook-facing GUIs must use workbook-scoped settings for workbook-specific
choices, with global extension settings acting only as defaults until a workbook
establishes its own value.

Use this split across analysis, import/export, diff, test, and future workbook
workflow panels:

- **Global extension settings** are for environment and default behavior:
  Python/backend setup, editor typing behavior, default block layout, default
  analysis profile, default import/export modes, documentation metadata glob,
  and sidebar/editor preferences.
- **Workbook-scoped settings** are for decisions tied to one workbook:
  import/export folder and mode, analysis severity visibility, tracked/untracked
  analysis rules, suppression visibility, per-workbook diff preferences,
  selected module state where useful, and future workbook test/run preferences.
- Some workbook-scoped settings are workbook-only and do not have a global
  fallback. Workbook-specific facts such as a selected sync folder belong only
  to that workbook.
- A workbook GUI action must not silently mutate a global setting. If the action
  is performed in the context of a workbook, persist it beside or for that
  workbook.
- Workbook-scoped settings inherit global defaults until the workbook has an
  explicit value. Add reset actions where useful so a workbook can return to the
  current global default.
- Refreshing an already-open workbook GUI must preserve active UI state and must
  not replay "on open" defaults. A fresh open may re-read workbook settings and
  global defaults.
- Displayed settings and runtime behavior must come from the same resolver.
  Avoid separate GUI-only, command-only, and agent-only configuration pipelines.
- Configuration writes should be debounced and deterministic. Malformed
  workbook-local settings should be reported with a clear recovery path rather
  than guessed. This is implemented for `<workbook>.xlide_settings.json`:
  invalid JSON, unknown keys, invalid sync modes, and invalid analysis settings
  surface as explicit workbook settings errors.

Implementation priority:

1. [x] Move the Analysis GUI's severity visibility and rule tracking/untracking
   to workbook-scoped analysis settings with global fallback. The Analysis
   Settings dialog now shows source provenance and can reset workbook overrides
   back to global defaults without mutating VS Code settings.
2. [x] Keep Import/Export settings in the workbook sidecar and ensure
   descriptions, command behavior, GUI state, and agent tools use the same
   planner/resolver.
3. [x] Keep workbook sidecar path, strict schema validation, and persistence in
   `src/workbookSettings.ts`, consumed by import/export, analysis settings,
   live diagnostics, webviews, and agent tools instead of feature-local
   settings helpers.
4. [ ] Apply the same resolver pattern to the future test runner GUI and any
   workbook diff/sync surface.

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

## Addendum: Host-Context Member Completion

> Status: DONE (first milestone). Implemented as a pure analyzer host layer
> (`src/analyzer/host/excelObjectModel.ts` verified member metadata +
> `src/analyzer/host/hostModel.ts` resolver +
> `src/analyzer/completion/memberAccess.ts` receiver-chain resolution) and a VS
> Code provider (`src/vbaMemberCompletion.ts`, trigger `.`). Covered by
> `tests/vbaMemberCompletion.test.ts` (23 tests). Resolves host globals,
> worksheet code names, `Me` by module kind, typed local/param/module variables,
> and member-access chains through return types. Member metadata is transcribed
> from the official Office VBA object-model reference (verified 2026-05-30), not
> LLM memory. Remaining work for later milestones: collection element typing
> (e.g. `Worksheets(1).`), UserForm/control members, and class-module `Me`
> members.
>
> Second-pass reconciliation: this addendum is not a standalone phase. It spans
> and refines several existing phases:
> - **Phase 4** gains a host-context symbol layer (host-injected globals +
>   worksheet code names + `Me`), kept separate from MS-VBAL language symbols.
> - **Phase 10** owns the Excel object-model member metadata and the
>   global -> type resolution table it describes below.
> - **Phases 6/7** wire the actual `.`-triggered completion / member listing UI.
> - **Phase 3 (done)** must be extended with member-access expression parsing
>   (currently Pending in the verification map) before `obj.member` can be
>   resolved positionally.
> Verification rule is unchanged: core VBA grammar is verified against MS-VBAL;
> host members are verified against Office VBA object-model docs, generated COM
> type-library metadata, or recorded VBE behavior fixtures — never LLM memory.
> Progress is tracked in the "Addendum - Host-Context Member Completion" table in
> `docs/spec/MS-VBAL.verification-map.md`.

The VBA language service must support member-completion popups for host-provided global objects, not only user-declared variables.

This is required for Excel VBA scenarios such as:

```vba
ThisWorkbook.
Application.
ActiveWorkbook.
ActiveSheet.
Sheet1.
Me.
```

These identifiers are not all ordinary local declarations. Some are injected by the host environment or generated from the workbook/project structure. The analyzer must therefore maintain a host-context symbol layer in addition to the pure VBA language symbol graph.

### Required behavior

When the user types:

```vba
ThisWorkbook.
```

The extension must open a completion popup showing verified members of the Excel `Workbook` object.

Example expected completions include, subject to verification:

```text
Worksheets
Sheets
Save
SaveAs
Close
Name
FullName
Path
VBProject
```

The type resolution rule is:

```text
ThisWorkbook -> Excel.Workbook -> Workbook members
```

When the user types:

```vba
Application.
```

The completion popup must resolve:

```text
Application -> Excel.Application -> Application members
```

When the user types:

```vba
ActiveSheet.
```

The completion popup must resolve to the best known type. In Excel this is usually a worksheet-like object, but it may represent different sheet types. If the type is ambiguous, completions must either show the common verified members or mark the result as ambiguous internally.

When the user types:

```vba
Sheet1.
```

The completion popup must resolve `Sheet1` from the workbook’s VBA project components, using the component’s code name, not merely the visible worksheet tab name.

When the user types:

```vba
Me.
```

The completion popup must resolve based on the current module context:

```text
Worksheet module -> worksheet members
ThisWorkbook module -> workbook members
UserForm module -> form/control members
Class module -> class members
```

### Verification requirement

All host object members must be verified against authoritative sources.

The VBA language grammar and semantics must still be verified against `MS-VBAL.pdf`.

Excel object model members must be verified against one of the following:

```text
1. Microsoft Office VBA object model documentation
2. Generated metadata from local COM type libraries
3. Empirical VBE behavior tests, recorded in fixtures
```

LLM-generated member lists must never be treated as authoritative.

### Metadata requirement

The implementation should maintain a host metadata table similar to:

```json
{
  "Excel.ThisWorkbook": {
    "resolvesTo": "Excel.Workbook"
  },
  "Excel.Application": {
    "resolvesTo": "Excel.Application"
  },
  "Excel.ActiveWorkbook": {
    "resolvesTo": "Excel.Workbook"
  },
  "Excel.ActiveSheet": {
    "resolvesTo": "Excel.SheetLike"
  }
}
```

Workbook-specific symbols must be derived from the actual project structure whenever possible.

For example:

```text
Component code name: Sheet1
Visible sheet name: Customers
Resolved symbol: Sheet1
Resolved type: Excel.Worksheet
```

### Acceptance criteria

The following must work before host-context completion is considered complete for the first milestone:

```vba
Sub Test()
    ThisWorkbook.
End Sub
```

Typing after the dot must show a completion popup.

The popup must include verified `Workbook` members.

The popup must use proper capitalization.

The popup must not invent unverified members.

The analyzer must distinguish between:

```text
ThisWorkbook
ActiveWorkbook
Workbook variables declared by the user
Worksheet code names such as Sheet1
The current object represented by Me
```

This feature is required because workbook-aware IntelliSense is one of XLIDE’s core advantages over generic VBA syntax extensions.
